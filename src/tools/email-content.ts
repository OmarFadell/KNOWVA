import { Client as GraphClient } from "@microsoft/teams.graph";
import type { ILogger } from "@microsoft/teams.common";
import type { AgentTool, AgentToolResult } from "../agent/loop";
import { getMailMessage, type MailMessageContent } from "../graph/mail";
import {
  classifyMailError,
  isToolError,
  type EmailAccess,
  type ResolvedMailbox,
} from "./email-access";

/**
 * The `get_email_content` tool: the full text of one message, on demand.
 *
 *   GET /me/messages/{id}?$select=...,body
 *
 * Separate from search_emails for the same reason get_document_metadata is
 * separate from search_documents: whole email bodies are large, most questions
 * are answerable from a snippet, and paying for a body only when the model has
 * decided which message matters keeps the turn affordable. Here it buys
 * something else too -- reading someone's entire email is a more consequential
 * act than seeing that it exists, and making it a distinct, deliberate call
 * means it shows up as one in the logs.
 *
 * The opt-out guard runs here exactly as it does in search_emails, and for the
 * same reason: this tool takes a messageId, and a messageId from an earlier
 * turn would otherwise be a way to keep reading mail after the user had said
 * stop. Every entry point to the mailbox is gated, not just the first one.
 *
 * Attachments are still filenames only. This tool fetches a body; it does not
 * fetch a single byte of any attached file (see ATTACHMENT_SELECT in
 * src/graph/mail.ts).
 *
 * Docs:
 *   https://learn.microsoft.com/en-us/graph/api/message-get
 */

export function createEmailContentTool(
  graph: GraphClient,
  access: EmailAccess,
  log: ILogger
): AgentTool {
  return {
    definition: {
      name: "get_email_content",
      description:
        "Read the full body of one email message. Pass the messageId exactly as search_emails " +
        "returned it. Use this when a snippet isn't enough to answer the question -- search " +
        "first, decide which message matters, then read that one. Returns the sender, " +
        "recipients, date, full text, and attachment filenames. Attachment contents are never " +
        "read.",
      inputSchema: {
        type: "object",
        properties: {
          messageId: {
            type: "string",
            description:
              "The messageId from a search_emails result, copied verbatim. Do not shorten, " +
              "reformat, or invent one.",
          },
          mailbox: {
            type: "string",
            description:
              "The same mailbox argument you passed to search_emails, if any. Omit it for the " +
              "user's own mailbox. A messageId only resolves in the mailbox it came from.",
          },
        },
        required: ["messageId"],
        additionalProperties: false,
      },
    },

    async run(input: Record<string, unknown>): Promise<AgentToolResult> {
      // Same gate, same position: before anything else happens.
      const blocked = await access.guard();
      if (blocked) return blocked;

      const messageId = typeof input.messageId === "string" ? input.messageId.trim() : "";
      if (!messageId) {
        return {
          content:
            "No messageId was provided. Call search_emails first, then call get_email_content " +
            "with the messageId of the message you want to read.",
          isError: true,
        };
      }

      const mailbox = access.resolveMailbox(input.mailbox);
      if (isToolError(mailbox)) return mailbox;

      log.info(`email-content: fetching message ${messageId} from ${mailbox.path}`);

      let message: MailMessageContent;
      try {
        message = await getMailMessage(graph, mailbox.path, messageId, log);
      } catch (err) {
        return classifyMailError(err, mailbox, log);
      }

      return { content: formatMessage(message, mailbox) };
    },
  };
}

/**
 * Renders one message for the model.
 *
 * The attribution header comes first and is never optional: the system prompt
 * requires who/to/when on every email claim, and a body handed over without
 * them is an invitation to state its contents as bare fact.
 */
function formatMessage(message: MailMessageContent, mailbox: ResolvedMailbox): string {
  const header = [
    `Subject: ${message.subject}`,
    `From: ${message.sender}`,
    `To: ${message.to.length > 0 ? message.to.join("; ") : "(no named recipients)"}`,
  ];

  if (message.cc.length > 0) header.push(`Cc: ${message.cc.join("; ")}`);

  header.push(`Date: ${message.timestamp} (UTC)`);
  header.push(`Mailbox: ${mailbox.label}`);

  if (message.categories.length > 0) {
    header.push(`Categories: ${message.categories.join(", ")}`);
  }
  if (message.webLink) {
    header.push(`Link: ${message.webLink}`);
  }

  header.push(
    message.attachmentNames.length > 0
      ? `Attachments (filenames only, contents NOT read): ${message.attachmentNames.join(", ")}`
      : message.hasAttachments
        ? "Attachments: this message has attachments, but their filenames could not be listed"
        : "Attachments: none"
  );

  const footer = [
    message.bodyTruncated
      ? "This body was truncated because the message is long. Do not claim to have read the whole " +
        "email, and do not infer what the cut-off part said."
      : "",
    "Attribute anything you take from this: name the sender, say when they sent it, and say it " +
      "came from an email. An email records that someone wrote something at a time -- not that " +
      "it is true, and not that it is still current.",
    "Apply the sensitivity rules: if this touches pay, HR, legal or contract matters, or reads " +
      "as personal rather than work correspondence, paraphrase it and cite it. Do not quote it.",
    "You have not read any attachment. Never describe what an attached file contains.",
  ]
    .filter(Boolean)
    .join(" ");

  return (
    "Full email message:\n\n" +
    header.join("\n") +
    "\n\n--- body ---\n" +
    (message.body || "(this message has no readable body text)") +
    "\n--- end of body ---\n\n" +
    footer
  );
}
