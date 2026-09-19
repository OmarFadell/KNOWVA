import { Client as GraphClient } from "@microsoft/teams.graph";
import type { ILogger } from "@microsoft/teams.common";
import { bodyToText } from "./chat-messages";

/**
 * Outlook mail reads, for the search_emails and get_email_content tools.
 *
 * Delegated only. Every call here goes out on the asking user's token from
 * Teams SSO -- Mail.Read for their own mailbox, Mail.Read.Shared for a shared
 * one. There is no app-only path and there must not be one: unlike the RSC chat
 * reads in src/graph/app-client.ts, an *application* Mail.Read grant is
 * tenant-wide over every mailbox in the organization, which is exactly the
 * blast radius this feature must not have. If a future change ever makes a mail
 * call from an app token, that is a bug, not an optimisation.
 *
 * ---------------------------------------------------------------------------
 * FOLDER SCOPING: WHY TWO REQUESTS, NOT ONE FILTERED REQUEST
 *
 * The scope is Inbox and Sent Items -- never Drafts, never Deleted Items, never
 * anything else. The obvious way to express that is a $filter on parentFolderId
 * alongside the $search. It is not available:
 *
 *   "$search cannot be combined with $filter or $orderby" on message
 *   collections. (https://learn.microsoft.com/en-us/graph/search-query-parameter)
 *
 * So the folder restriction moves into the URL path instead, where it costs
 * nothing and cannot be got wrong:
 *
 *   GET /me/mailFolders/inbox/messages?$search="..."
 *   GET /me/mailFolders/sentitems/messages?$search="..."
 *
 * Two round trips, merged in this process. That is a feature rather than a
 * concession: a path-scoped search *cannot* return a draft or a deleted item,
 * whereas a filter expression is one typo away from doing so. `inbox` and
 * `sentitems` are Graph well-known folder names, so no folder ids need
 * resolving and this code does not care how the mailbox is laid out.
 *
 * Results come back sorted by send date, which is what we want -- and just as
 * well, because $orderby is unavailable for the same reason $filter is.
 * ---------------------------------------------------------------------------
 *
 * Docs:
 *   https://learn.microsoft.com/en-us/graph/search-query-parameter
 *   https://learn.microsoft.com/en-us/graph/api/user-list-messages
 *   https://learn.microsoft.com/en-us/graph/api/message-list-attachments
 */

/** The only two folders search_emails may look in. Well-known Graph folder names. */
const SEARCHED_FOLDERS = ["inbox", "sentitems"] as const;

/** Messages requested per folder. Doubled across the two, then capped by the caller. */
const PAGE_SIZE = 15;

/**
 * Fields pulled back for a search hit.
 *
 * `body` is deliberately absent: a search must not drag whole message bodies
 * through the model's context, and get_email_content exists precisely so that
 * reading an entire email is a separate, deliberate act. `bodyPreview` is
 * Graph's own first-~255-character snippet, which is enough to judge a hit by.
 */
const SEARCH_SELECT = [
  "id",
  "subject",
  "from",
  "toRecipients",
  "ccRecipients",
  "receivedDateTime",
  "sentDateTime",
  "bodyPreview",
  "hasAttachments",
  "categories",
  "webLink",
].join(",");

/** Fields for the single-message fetch: the same list plus the body itself. */
const CONTENT_SELECT = SEARCH_SELECT + ",body";

/**
 * ATTACHMENT METADATA ONLY -- this $select is load-bearing, do not drop it.
 *
 * GET /messages/{id}/attachments with no $select returns `contentBytes` for
 * every fileAttachment: the whole file, base64-encoded, in the response body.
 * This milestone lists attachment filenames and nothing more, so naming these
 * four metadata fields is what keeps attachment *content* out of the process
 * entirely, rather than fetching it and then choosing not to look at it.
 */
const ATTACHMENT_SELECT = "id,name,contentType,size";

/** Attachment names listed per message. A thread with 40 files does not need enumerating. */
const MAX_ATTACHMENT_NAMES = 10;

/** How much of a body get_email_content hands the model, in characters. */
const MAX_BODY_CHARS = 6000;

interface GraphEmailAddress {
  name?: string;
  address?: string;
}

interface GraphRecipient {
  emailAddress?: GraphEmailAddress;
}

interface GraphMessage {
  id?: string;
  subject?: string;
  from?: GraphRecipient | null;
  toRecipients?: GraphRecipient[];
  ccRecipients?: GraphRecipient[];
  receivedDateTime?: string;
  sentDateTime?: string;
  bodyPreview?: string;
  hasAttachments?: boolean;
  categories?: string[];
  webLink?: string;
  body?: { contentType?: string; content?: string };
}

interface MessageListResponse {
  value?: GraphMessage[];
}

interface GraphAttachment {
  id?: string;
  name?: string;
  contentType?: string;
  size?: number;
}

interface AttachmentListResponse {
  value?: GraphAttachment[];
}

export type MailFolderLabel = "Inbox" | "Sent Items";

/** One search hit, flattened to what attribution and citation need. */
export interface MailMessageRecord {
  id: string;
  subject: string;
  /** Sender as "Name <address>", or whichever half Graph supplied. */
  sender: string;
  /** To: recipients, formatted the same way. */
  to: string[];
  /** Cc: recipients. */
  cc: string[];
  /** ISO 8601, UTC. Sent date where present, received date otherwise. */
  timestamp: string;
  /** Which of the two searched folders this came from. */
  folder: MailFolderLabel;
  /** Graph's own short preview of the body. */
  snippet: string;
  hasAttachments: boolean;
  /** Outlook categories on the message, verbatim. */
  categories: string[];
  /** Deep link that opens the message in Outlook on the web. */
  webLink?: string;
}

/** A whole message, as get_email_content returns it. */
export interface MailMessageContent extends MailMessageRecord {
  /** Body converted to plain text, and truncated at MAX_BODY_CHARS. */
  body: string;
  /** True when the body was cut short. */
  bodyTruncated: boolean;
  /** Filenames only. Never content -- see ATTACHMENT_SELECT. */
  attachmentNames: string[];
}

/**
 * The Graph path prefix for a mailbox.
 *
 * `/me` for the asking user's own mailbox; `/users/{address}` for a shared one.
 * Naming a shared mailbox does not open it: Exchange still evaluates
 * Mail.Read.Shared against *this caller's* delegate rights on that mailbox, so
 * a user who was never granted access gets a 403 no matter what is configured.
 */
export function mailboxPath(address?: string): string {
  return address ? "/users/" + encodeURIComponent(address) : "/me";
}

/**
 * Searches Inbox and Sent Items for one KQL query and merges the results.
 *
 * A failure in one folder does not fail the search: a mailbox with no Sent
 * Items folder, or a shared mailbox the caller can only partly read, should
 * still yield whatever the other folder had. The returned `foldersFailed` count
 * lets the caller tell the model coverage was incomplete, rather than quietly
 * presenting half a search as a whole one.
 */
export async function searchMailbox(
  graph: GraphClient,
  mailbox: string,
  kqlQuery: string,
  log: ILogger
): Promise<{ messages: MailMessageRecord[]; foldersFailed: number; firstError?: unknown }> {
  const messages: MailMessageRecord[] = [];
  let foldersFailed = 0;
  let firstError: unknown;

  for (const folder of SEARCHED_FOLDERS) {
    // $search's value is a quoted KQL string. The surrounding double quotes are
    // part of the syntax, so they go on before encoding rather than being added
    // by it.
    const path =
      mailbox +
      "/mailFolders/" +
      folder +
      "/messages" +
      "?$search=" +
      encodeURIComponent('"' + kqlQuery + '"') +
      "&$select=" +
      encodeURIComponent(SEARCH_SELECT) +
      "&$top=" +
      PAGE_SIZE;

    try {
      const response = await graph.http.get<MessageListResponse>(path);
      const label: MailFolderLabel = folder === "inbox" ? "Inbox" : "Sent Items";
      for (const message of response.data.value ?? []) {
        const record = toRecord(message, label);
        if (record) messages.push(record);
      }
    } catch (err) {
      foldersFailed++;
      if (firstError === undefined) firstError = err;
      log.warn(
        `mail: search of ${mailbox}/mailFolders/${folder} failed; continuing with the other folder`,
        { message: (err as { message?: string })?.message }
      );
    }
  }

  log.info(
    `mail: searched ${SEARCHED_FOLDERS.length - foldersFailed}/${SEARCHED_FOLDERS.length} ` +
      `folder(s) of ${mailbox} -- ${messages.length} message(s) returned`
  );

  return { messages, foldersFailed, firstError };
}

/**
 * Fetches one message in full, with its attachment filenames.
 *
 * Note the path: `{mailbox}/messages/{id}`, not a folder-scoped one. A message
 * id is opaque and already identifies its own folder, and get_email_content is
 * only ever called with an id search_emails returned -- which by construction
 * came from Inbox or Sent Items. Re-scoping the fetch by folder would add a
 * failure mode (guessing the wrong folder) without adding a guarantee.
 */
export async function getMailMessage(
  graph: GraphClient,
  mailbox: string,
  messageId: string,
  log: ILogger
): Promise<MailMessageContent> {
  const path =
    mailbox +
    "/messages/" +
    encodeURIComponent(messageId) +
    "?$select=" +
    encodeURIComponent(CONTENT_SELECT);

  const response = await graph.http.get<GraphMessage>(path);
  const message = response.data;

  // Which folder this sits in is unknown on a direct fetch and is not worth a
  // second request to establish, so it is inferred: a message the user sent has
  // a sentDateTime and no meaningful receivedDateTime.
  const record = toRecord(message, message.receivedDateTime ? "Inbox" : "Sent Items");
  if (!record) {
    throw new Error(`Graph returned a message with no id for ${mailbox}/messages/${messageId}`);
  }

  const text = emailBodyToText(message.body?.content ?? "", message.body?.contentType);
  const truncated = text.length > MAX_BODY_CHARS;

  const attachmentNames = record.hasAttachments
    ? await listAttachmentNames(graph, mailbox, messageId, log)
    : [];

  log.info(
    `mail: fetched message ${messageId} from ${mailbox} ` +
      `(${text.length} body chars, ${attachmentNames.length} attachment name(s))`
  );

  return {
    ...record,
    body: truncated ? text.slice(0, MAX_BODY_CHARS) + "..." : text,
    bodyTruncated: truncated,
    attachmentNames,
  };
}

/**
 * Attachment FILENAMES for one message. Never content -- see ATTACHMENT_SELECT.
 *
 * Best-effort: a message whose attachment list cannot be read is still a useful
 * result, so this returns an empty list rather than failing the whole call.
 */
export async function listAttachmentNames(
  graph: GraphClient,
  mailbox: string,
  messageId: string,
  log: ILogger
): Promise<string[]> {
  const path =
    mailbox +
    "/messages/" +
    encodeURIComponent(messageId) +
    "/attachments?$select=" +
    encodeURIComponent(ATTACHMENT_SELECT);

  try {
    const response = await graph.http.get<AttachmentListResponse>(path);
    return (response.data.value ?? [])
      .map((attachment) => attachment.name?.trim())
      .filter((name): name is string => Boolean(name))
      .slice(0, MAX_ATTACHMENT_NAMES);
  } catch (err) {
    log.warn(`mail: could not list attachments for message ${messageId}`, {
      message: (err as { message?: string })?.message,
    });
    return [];
  }
}

function toRecord(message: GraphMessage, folder: MailFolderLabel): MailMessageRecord | null {
  if (!message.id) return null;

  return {
    id: message.id,
    subject: message.subject?.trim() || "(no subject)",
    sender: formatRecipient(message.from),
    to: (message.toRecipients ?? []).map(formatRecipient),
    cc: (message.ccRecipients ?? []).map(formatRecipient),
    // Sent date first: it is the one date that means the same thing on a
    // message the user sent and one they received.
    timestamp: message.sentDateTime || message.receivedDateTime || "unknown",
    folder,
    snippet: (message.bodyPreview ?? "").replace(/\s+/g, " ").trim(),
    hasAttachments: message.hasAttachments === true,
    categories: (message.categories ?? []).filter(Boolean),
    webLink: message.webLink,
  };
}

/**
 * "Display Name <address>", degrading to whichever half Graph supplied.
 *
 * An address with no display name is still an attribution; an empty string is
 * not, and src/llm/prompt.ts requires every email citation to name a sender. So
 * this never returns "" -- it says the sender is unidentified, which the model
 * can relay honestly instead of quietly omitting the who.
 */
function formatRecipient(recipient: GraphRecipient | null | undefined): string {
  const name = recipient?.emailAddress?.name?.trim();
  const address = recipient?.emailAddress?.address?.trim();

  if (name && address) return `${name} <${address}>`;
  if (address) return address;
  if (name) return name;
  return "an unidentified sender";
}

/**
 * Email HTML to plain text.
 *
 * Delegates the tag and entity work to the chat-message converter, after
 * stripping what only email bodies contain: a <head> (title, meta, embedded
 * CSS), Outlook's conditional comments, and ordinary HTML comments. Left in,
 * those survive as visible noise and spend the model's context on markup nobody
 * wrote.
 */
export function emailBodyToText(content: string, contentType?: string): string {
  if (!content) return "";
  if (contentType !== "html") return content.trim();

  const stripped = content
    .replace(/<head\b[^>]*>[\s\S]*?<\/head>/gi, " ")
    .replace(/<!--[\s\S]*?-->/g, " ");

  return bodyToText(stripped, "html");
}
