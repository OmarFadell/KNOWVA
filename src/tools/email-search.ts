import { Client as GraphClient } from "@microsoft/teams.graph";
import type { ILogger } from "@microsoft/teams.common";
import config from "../../config";
import type { AgentTool, AgentToolResult } from "../agent/loop";
import { listAttachmentNames, searchMailbox, type MailMessageRecord } from "../graph/mail";
import {
  classifyMailError,
  isToolError,
  type EmailAccess,
  type ResolvedMailbox,
} from "./email-access";

/**
 * The `search_emails` tool: finds messages in Outlook.
 *
 *   GET /me/mailFolders/{inbox|sentitems}/messages?$search="..."
 *
 * The most sensitive source Knowva has been connected to. Three boundaries hold
 * it in, and each is enforced somewhere different on purpose:
 *
 *   1. WHOSE MAIL -- the delegated token, so only the asking user's mailbox (or
 *      a shared one Exchange already lets them open). Enforced by Graph.
 *   2. WHICH FOLDERS -- Inbox and Sent Items only, never Drafts or Deleted
 *      Items. Enforced by the URL path in src/graph/mail.ts.
 *   3. WHETHER AT ALL -- the per-user /disable-email opt-out. Enforced by the
 *      guard in src/tools/email-access.ts, before any Graph call is built.
 *
 * What it deliberately does NOT do is read attachments. Filenames are listed;
 * content is never fetched. See ATTACHMENT_SELECT in src/graph/mail.ts, where
 * that is a property of the request rather than a promise made in a comment.
 *
 * ---------------------------------------------------------------------------
 * CATEGORY NARROWING: WHY NOT THE $filter THE BRIEF SPECIFIED
 *
 * The intended narrowing was `$filter=categories/any(c:c eq '{category}')`
 * alongside the `$search`. Microsoft Graph does not allow that combination:
 *
 *   "$search cannot be combined with $filter or $orderby" on message
 *   collections. (https://learn.microsoft.com/en-us/graph/search-query-parameter)
 *
 * Using the $filter form would therefore mean giving up keyword search entirely
 * and returning "the most recent messages tagged X", which is a different tool.
 * So the narrowing is done two ways at once instead, and the redundancy is the
 * point:
 *
 *   (a) `category:"X"` is added to the KQL inside $search. Exchange's KQL does
 *       document a `category` searchable property, so this should narrow
 *       server-side, preserving both recall and ranking.
 *       (https://learn.microsoft.com/en-us/purview/ediscovery-keyword-queries-and-search-conditions)
 *       It is NOT in Graph's own table of supported message search properties,
 *       though, which is why it is not trusted on its own.
 *   (b) Every returned message is checked against its own `categories` array in
 *       this process, and anything not actually carrying the category is
 *       dropped.
 *
 * If (a) works, (b) is a no-op. If Exchange quietly ignores `category:` and
 * treats it as free text, (b) turns what would have been a silent lie -- a pile
 * of untagged mail presented as the tagged results -- into an honest empty
 * result, which is exactly the case the system prompt tells Claude to handle by
 * retrying without a category. Failing visibly beats narrowing invisibly.
 * ---------------------------------------------------------------------------
 */

/** Messages handed to the model. Emails are long; a dozen would crowd out the answer. */
const MAX_RESULTS = 8;

/**
 * Messages whose attachment filenames are fetched, per search.
 *
 * Each costs its own round trip, so this is capped rather than run over every
 * hit. Results past the cap still report *that* they have attachments -- they
 * just do not name the files, and the tool says so rather than implying there
 * were none.
 */
const MAX_ATTACHMENT_LOOKUPS = 5;

/** Longest query accepted. Long enough for a real question, short enough to stay a search. */
const MAX_QUERY_LENGTH = 400;

export function createEmailSearchTool(
  graph: GraphClient,
  access: EmailAccess,
  log: ILogger
): AgentTool {
  return {
    definition: {
      name: "search_emails",
      description:
        "Search the user's Outlook email and return matching messages with their sender, " +
        "recipients, date, a short snippet, and the filenames of any attachments. Searches the " +
        "Inbox and Sent Items only -- never drafts or deleted mail. By default it searches the " +
        "mailbox of the person you are talking to. Attachment contents are never read, only " +
        "filenames. Use this when the user asks what someone emailed, what was agreed or sent " +
        "over email, or to find a specific message. Call get_email_content afterwards if you " +
        "need the full text of one message.",
      inputSchema: {
        type: "object",
        properties: {
          query: {
            type: "string",
            description:
              "What to look for. Plain keywords work best. You may also use Outlook search " +
              'syntax such as from:priya, subject:budget, or hasAttachments:true. At most ' +
              `${MAX_QUERY_LENGTH} characters.`,
          },
          mailbox: {
            type: "string",
            description:
              describeMailboxArgument(),
          },
          category: {
            type: "string",
            description:
              "Optional Outlook category to narrow to, e.g. 'Project Alpha'. Prefer this for " +
              "project-scoped questions. Only messages actually tagged with this category are " +
              "returned, so an empty result means nothing tagged matched -- retry without it.",
          },
        },
        required: ["query"],
        additionalProperties: false,
      },
    },

    async run(input: Record<string, unknown>): Promise<AgentToolResult> {
      // FIRST, before anything is built or any request is shaped: has this user
      // opted out? Nothing below this line may move up above it.
      const blocked = await access.guard();
      if (blocked) return blocked;

      const rawQuery = typeof input.query === "string" ? input.query : "";
      const query = rawQuery.trim();
      if (!query) {
        return {
          content:
            "No search terms were provided. Call search_emails again with the keywords to look for.",
          isError: true,
        };
      }
      if (query.length > MAX_QUERY_LENGTH) {
        return {
          content:
            `That query is ${query.length} characters; email search accepts at most ` +
            `${MAX_QUERY_LENGTH}. Shorten it to the distinctive keywords and try again.`,
          isError: true,
        };
      }

      const mailbox = access.resolveMailbox(input.mailbox);
      if (isToolError(mailbox)) return mailbox;

      const category = typeof input.category === "string" ? input.category.trim() : "";
      const kql = buildKqlQuery(query, category);

      log.info(
        `email-search: searching ${mailbox.path} ` +
          `(queryChars=${query.length}${category ? `, category="${category}"` : ""})`
      );

      const { messages, foldersFailed, firstError } = await searchMailbox(
        graph,
        mailbox.path,
        kql,
        log
      );

      // Both folders failed, so nothing was searched at all. Report the actual
      // cause rather than "no results found", which would be a lie the model
      // would pass straight on to the user.
      if (foldersFailed === 2) {
        return classifyMailError(firstError, mailbox, log);
      }

      const deduped = dedupeById(messages);
      const narrowed = category ? deduped.filter((m) => hasCategory(m, category)) : deduped;
      const ranked = narrowed
        .sort((a, b) => Date.parse(b.timestamp) - Date.parse(a.timestamp))
        .slice(0, MAX_RESULTS);

      log.info(
        `email-search: ${deduped.length} message(s) matched, ` +
          `${narrowed.length} after category narrowing, ${ranked.length} returned ` +
          `(foldersFailed=${foldersFailed})`
      );

      if (ranked.length === 0) {
        return { content: noResultsMessage(query, category, mailbox, foldersFailed) };
      }

      const withAttachments = await attachFilenames(graph, mailbox.path, ranked, log);

      return {
        content: formatResults(withAttachments, mailbox, {
          category,
          foldersFailed,
          attachmentLookupsCapped: countAttachmentLookupsSkipped(ranked),
        }),
      };
    },
  };
}

/**
 * The mailbox argument's description is built from config rather than written
 * out, so the model is told the actual mailbox names this deployment allows
 * instead of being invited to guess at plausible ones.
 */
function describeMailboxArgument(): string {
  const configured = config.sharedMailboxes;
  if (configured.length === 0) {
    return (
      "Leave this out. No shared mailboxes are configured on this deployment, so only the " +
      "user's own mailbox can be searched."
    );
  }
  return (
    "Optional. Omit it to search the user's own mailbox, which is the default. The only other " +
    "mailboxes that can be searched are these configured shared mailboxes, named exactly: " +
    configured.map((mailbox) => `"${mailbox.name}"`).join(", ") +
    ". Never pass any other value -- a colleague's name or address will not work."
  );
}

/**
 * Builds the KQL string that goes inside $search="...".
 *
 * Double quotes are stripped rather than escaped. A stray quote inside the
 * search string breaks the surrounding KQL literal, and the failure mode is a
 * 400 that reads like a permissions problem; the model is passing natural
 * keywords, so losing a quote character costs nothing worth defending.
 */
function buildKqlQuery(query: string, category: string): string {
  const safeQuery = query.replace(/"/g, " ").replace(/\s+/g, " ").trim();
  if (!category) return safeQuery;

  const safeCategory = category.replace(/"/g, "").trim();
  return `${safeQuery} AND category:"${safeCategory}"`;
}

/**
 * A message can legitimately appear in both folder searches -- most obviously a
 * message the user sent to themselves. Graph message ids are stable within a
 * mailbox, so the id is the right key.
 */
function dedupeById(messages: MailMessageRecord[]): MailMessageRecord[] {
  const seen = new Map<string, MailMessageRecord>();
  for (const message of messages) {
    if (!seen.has(message.id)) seen.set(message.id, message);
  }
  return [...seen.values()];
}

/** Outlook category names are case-preserving but not case-sensitive in practice. */
function hasCategory(message: MailMessageRecord, category: string): boolean {
  const wanted = category.toLowerCase();
  return message.categories.some((c) => c.toLowerCase() === wanted);
}

function countAttachmentLookupsSkipped(messages: MailMessageRecord[]): number {
  const withAttachments = messages.filter((m) => m.hasAttachments).length;
  return Math.max(0, withAttachments - MAX_ATTACHMENT_LOOKUPS);
}

/**
 * Fetches attachment filenames for the results that have any, up to the cap.
 * Sequential rather than parallel, matching search_conversations: a burst of
 * simultaneous Graph reads is the shape that earns a 429.
 */
async function attachFilenames(
  graph: GraphClient,
  mailbox: string,
  messages: MailMessageRecord[],
  log: ILogger
): Promise<(MailMessageRecord & { attachmentNames?: string[] })[]> {
  const results: (MailMessageRecord & { attachmentNames?: string[] })[] = [];
  let lookups = 0;

  for (const message of messages) {
    if (message.hasAttachments && lookups < MAX_ATTACHMENT_LOOKUPS) {
      lookups++;
      results.push({
        ...message,
        attachmentNames: await listAttachmentNames(graph, mailbox, message.id, log),
      });
    } else {
      results.push(message);
    }
  }

  return results;
}

function noResultsMessage(
  query: string,
  category: string,
  mailbox: ResolvedMailbox,
  foldersFailed: number
): string {
  const partial =
    foldersFailed > 0
      ? " Note that one of the two folders could not be searched, so coverage was incomplete -- say so."
      : "";

  if (category) {
    // The fallback path the system prompt tells Claude to take. Naming the
    // retry explicitly here makes it far more likely to actually happen.
    return (
      `No emails tagged with the "${category}" category matched "${query}" in ${mailbox.label} ` +
      "(Inbox and Sent Items). Nothing is tagged that way, or nothing tagged matched. " +
      "Call search_emails again with the same query and NO category argument before concluding " +
      `there is nothing.${partial}`
    );
  }

  return (
    `No emails matching "${query}" were found in ${mailbox.label} (Inbox and Sent Items only). ` +
    "Tell the user nothing came up, and mention that drafts and deleted mail are never searched. " +
    "Do NOT answer from your own knowledge as though you had found an email." +
    partial
  );
}

/**
 * Renders results for the model.
 *
 * Every message carries sender, recipients and a UTC timestamp, because
 * src/llm/prompt.ts forbids citing an email without all three -- so the tool
 * must never hand back a result missing any of them. src/graph/mail.ts holds up
 * its end by never returning an empty sender.
 */
function formatResults(
  messages: (MailMessageRecord & { attachmentNames?: string[] })[],
  mailbox: ResolvedMailbox,
  context: { category: string; foldersFailed: number; attachmentLookupsCapped: number }
): string {
  const blocks = messages.map((message, i) => {
    const lines = [
      `[${i + 1}] "${message.subject}"`,
      `    from: ${message.sender}`,
      `    to: ${message.to.length > 0 ? message.to.join("; ") : "(no named recipients)"}`,
    ];

    if (message.cc.length > 0) lines.push(`    cc: ${message.cc.join("; ")}`);

    lines.push(`    date: ${message.timestamp} (UTC)`);
    lines.push(`    folder: ${message.folder}`);
    lines.push(`    messageId: ${message.id}`);

    if (message.categories.length > 0) {
      lines.push(`    categories: ${message.categories.join(", ")}`);
    }
    if (message.webLink) {
      lines.push(`    link: ${message.webLink}`);
    }

    if (message.attachmentNames && message.attachmentNames.length > 0) {
      lines.push(
        `    attachments (filenames only, contents NOT read): ${message.attachmentNames.join(", ")}`
      );
    } else if (message.hasAttachments) {
      lines.push("    attachments: this message has attachments; their filenames were not listed");
    }

    lines.push(`    snippet: ${message.snippet || "(no preview text)"}`);
    return lines.join("\n");
  });

  const caveats: string[] = [
    `Searched ${mailbox.label}: Inbox and Sent Items only. Drafts and deleted mail were not searched.`,
    "These are snippets, not whole emails. Call get_email_content with a messageId if you need the full text.",
  ];

  if (context.category) {
    caveats.push(
      `Narrowed to messages tagged "${context.category}"; untagged messages were excluded.`
    );
  }
  if (context.foldersFailed > 0) {
    caveats.push(
      `${context.foldersFailed} of the 2 folders could not be searched, so results may be incomplete -- say so.`
    );
  }
  if (context.attachmentLookupsCapped > 0) {
    caveats.push(
      `${context.attachmentLookupsCapped} further result(s) have attachments whose filenames were not looked up.`
    );
  }

  return (
    "Email search results. This is real correspondence between real people. Cite every claim " +
    "you take from it with the sender, the recipients where they matter, and the date. Apply " +
    "the sensitivity rules: paraphrase rather than quote anything to do with pay, HR, legal or " +
    "personal matters. Attachment contents were NOT read -- only filenames are listed, so never " +
    "describe what an attachment contains.\n\n" +
    blocks.join("\n\n") +
    "\n\n" +
    caveats.join(" ")
  );
}
