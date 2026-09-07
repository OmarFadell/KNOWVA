import { Client as GraphClient } from "@microsoft/teams.graph";
import type { ILogger } from "@microsoft/teams.common";
import { describeError } from "../errors";
import type { AgentTool, AgentToolResult } from "../agent/loop";

/**
 * The `get_document_metadata` tool: who last touched a document, and when.
 *
 *   GET https://graph.microsoft.com/v1.0/shares/{token}/driveItem
 *
 * Deliberately a separate tool rather than enrichment baked into
 * search_documents: metadata costs an extra Graph round trip per document, and
 * most questions never need it. Claude chains -- search, spot the relevant
 * document, then fetch metadata for that one on demand.
 *
 * The Retrieval API hands back a `webUrl`, not a drive item id, so the lookup
 * goes through the /shares endpoint, which accepts a URL-safe-base64 encoding
 * of any file URL the caller can already reach.
 *
 * PERMISSIONS CAVEAT: the shares-get reference lists Files.ReadWrite /
 * Files.ReadWrite.All / Sites.ReadWrite.All as the delegated permissions -- all
 * *write* scopes, because /shares can redeem a sharing link. We deliberately
 * call it with only the read scopes the SSO connection already carries
 * (Files.Read.All + Sites.Read.All), which is expected to work because the user
 * already has access to the item and nothing is being redeemed. If this starts
 * returning 403, that assumption is what broke -- see classifyError below.
 *
 * Docs:
 *   https://learn.microsoft.com/en-us/graph/api/shares-get?view=graph-rest-1.0
 *   https://learn.microsoft.com/en-us/graph/api/resources/driveitem
 */

/** Microsoft Graph `identity`: an actor, which may be a person, an app, or a device. */
interface GraphIdentity {
  id?: string;
  displayName?: string;
  email?: string;
}

/**
 * Graph `identitySet`. SharePoint drive items actually return a
 * `sharePointIdentitySet`, which adds siteUser/group on top -- hence the extra
 * optional members. Every one of these is optional: a file last touched by a
 * workflow or retention policy has an `application` and no `user` at all.
 */
interface GraphIdentitySet {
  user?: GraphIdentity;
  application?: GraphIdentity;
  device?: GraphIdentity;
  siteUser?: GraphIdentity;
  group?: GraphIdentity;
}

interface DriveItem {
  id?: string;
  name?: string;
  size?: number;
  webUrl?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  createdBy?: GraphIdentitySet;
  lastModifiedBy?: GraphIdentitySet;
  file?: { mimeType?: string };
  folder?: unknown;
}

export function createDocumentMetadataTool(graph: GraphClient, log: ILogger): AgentTool {
  return {
    definition: {
      name: "get_document_metadata",
      description:
        "Look up metadata for one SharePoint document: when it was last modified, who last " +
        "modified it, when it was created, who created it, its file name and size. Pass the " +
        "document's webUrl exactly as returned by search_documents. Use this whenever the user " +
        "asks how current, recent, or up to date a document is, or who wrote or last changed " +
        "it. All timestamps are returned in UTC (ISO 8601).",
      inputSchema: {
        type: "object",
        properties: {
          webUrl: {
            type: "string",
            description:
              "The full URL of the document, copied verbatim from a search_documents result. " +
              "Must be an absolute http(s) URL.",
          },
        },
        required: ["webUrl"],
        additionalProperties: false,
      },
    },

    async run(input: Record<string, unknown>): Promise<AgentToolResult> {
      const rawUrl = typeof input.webUrl === "string" ? input.webUrl : "";
      const webUrl = rawUrl.trim();

      const malformed = validateUrl(webUrl);
      if (malformed) {
        return { content: malformed, isError: true };
      }

      const token = encodeSharingUrl(webUrl);
      log.info(`document-metadata: resolving driveItem for ${webUrl}`);

      let item: DriveItem;
      try {
        const response = await graph.http.get<DriveItem>(`/shares/${token}/driveItem`);
        item = response.data;
      } catch (err) {
        return classifyError(err, log);
      }

      log.info(`document-metadata: resolved "${item.name ?? "(unnamed)"}"`);
      return { content: formatMetadata(item, webUrl) };
    },
  };
}

/**
 * Encodes a URL as a Graph sharing token.
 *
 * Verbatim from the shares-get reference: base64 the UTF-8 bytes, drop trailing
 * '=' padding, then '/' -> '_' and '+' -> '-', then prefix 'u!'. Getting any
 * step wrong produces a 400 that reads like a permissions problem, so this stays
 * exactly as documented -- do not "simplify" it to base64url without checking
 * that Node's base64url also strips padding (it does, but the explicit form is
 * what the docs specify and what is easy to verify against them).
 */
export function encodeSharingUrl(url: string): string {
  const base64 = Buffer.from(url, "utf8").toString("base64");
  const urlSafe = base64.replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  return `u!${urlSafe}`;
}

/** Returns an error message if the URL is unusable, else null. */
function validateUrl(webUrl: string): string | null {
  if (!webUrl) {
    return (
      "No document URL was provided. Call get_document_metadata again with the webUrl from a " +
      "search_documents result."
    );
  }

  let parsed: URL;
  try {
    parsed = new URL(webUrl);
  } catch {
    return (
      `"${webUrl}" is not a valid URL. Use the document link exactly as search_documents returned ` +
      "it, without shortening or paraphrasing it."
    );
  }

  if (parsed.protocol !== "https:" && parsed.protocol !== "http:") {
    return (
      `"${webUrl}" is not an http(s) URL, so it cannot identify a SharePoint document. ` +
      "Use the link from a search_documents result."
    );
  }

  return null;
}

/**
 * Names the actor behind a createdBy/lastModifiedBy.
 *
 * `user` is absent whenever an automated process touched the file -- a workflow,
 * a retention policy, a migration tool -- in which case Graph populates
 * `application` instead. Saying "an application" is the honest answer there;
 * silently falling through to "unknown" would invite the model to guess a person.
 */
function describeActor(identity: GraphIdentitySet | undefined): string {
  if (!identity) return "unknown";

  if (identity.user?.displayName) return identity.user.displayName;
  if (identity.siteUser?.displayName) return identity.siteUser.displayName;
  if (identity.application?.displayName) {
    return `${identity.application.displayName} (an application, not a person)`;
  }
  if (identity.device?.displayName) return `${identity.device.displayName} (a device)`;
  if (identity.group?.displayName) return `${identity.group.displayName} (a group)`;

  // Ids with no display name still beat "unknown" -- and still tell the model
  // which kind of actor it was.
  if (identity.user?.id) return `a user account (id ${identity.user.id}, no display name)`;
  if (identity.application?.id) {
    return `an application (id ${identity.application.id}, not a person)`;
  }

  return "unknown";
}

function formatMetadata(item: DriveItem, requestedUrl: string): string {
  const lines = [
    `Document metadata for "${item.name ?? "(name unavailable)"}":`,
    `- Last modified: ${item.lastModifiedDateTime ?? "unknown"} (UTC) by ${describeActor(item.lastModifiedBy)}`,
    `- Created: ${item.createdDateTime ?? "unknown"} (UTC) by ${describeActor(item.createdBy)}`,
    `- Size: ${formatSize(item.size)}`,
    `- Link: ${item.webUrl ?? requestedUrl}`,
  ];

  return (
    lines.join("\n") +
    "\n\nTimestamps are UTC. Present them readably and say they are UTC. If an application " +
    "rather than a person is named, say so rather than attributing the change to someone."
  );
}

function formatSize(size: number | undefined): string {
  if (typeof size !== "number" || !Number.isFinite(size)) return "unknown";
  if (size < 1024) return `${size} bytes`;
  const units = ["KB", "MB", "GB", "TB"];
  let value = size / 1024;
  let unit = 0;
  while (value >= 1024 && unit < units.length - 1) {
    value /= 1024;
    unit++;
  }
  return `${value.toFixed(1)} ${units[unit]} (${size} bytes)`;
}

/**
 * Turns a failed lookup into something the model can relay in plain language.
 * Same convention as sharepoint-search.ts: never let a stack trace reach a user.
 */
function classifyError(err: unknown, log: ILogger): AgentToolResult {
  log.error("document-metadata: /shares driveItem lookup failed", describeError(err), err);

  const e = err as { response?: { status?: number; data?: unknown } };
  const status = e?.response?.status;

  if (status === 400) {
    // /shares rejects a token it cannot decode with a 400. Since we build the
    // token ourselves, this is far more likely an unusual URL than a caller
    // mistake -- but it is emphatically NOT a permissions problem, and saying so
    // saves chasing consent that is already correct.
    return {
      content:
        "The document link could not be resolved -- the service rejected it as malformed. " +
        "This is not a permissions problem. Tell the user you couldn't look up that document's details.",
      isError: true,
    };
  }

  if (status === 401) {
    return {
      content:
        "The metadata lookup could not authenticate. Tell the user to sign out and sign back in, then retry.",
      isError: true,
    };
  }

  if (status === 403) {
    // The most likely cause, and the one worth naming loudly in the log: the
    // shares-get reference documents this endpoint at Files.ReadWrite /
    // Files.ReadWrite.All / Sites.ReadWrite.All, while our SSO connection only
    // carries the read equivalents. We expect read scopes to suffice for an item
    // the user can already reach, but if that expectation is wrong this is
    // exactly how it surfaces -- and it looks identical to a genuine ACL denial.
    log.warn(
      "document-metadata: 403 from /shares. If this is consistent rather than per-document, the " +
        "likely cause is that /shares requires Files.ReadWrite(.All) or Sites.ReadWrite.All -- the " +
        "OAuth connection currently grants only Files.Read.All + Sites.Read.All. " +
        "See https://learn.microsoft.com/en-us/graph/api/shares-get?view=graph-rest-1.0#permissions"
    );
    return {
      content:
        "Access to that document's metadata was denied. Tell the user you can see the document in " +
        "search results but aren't allowed to read its details.",
      isError: true,
    };
  }

  if (status === 404) {
    return {
      content:
        "No document was found at that link. It may have been moved, renamed, or deleted since it " +
        "was indexed. Tell the user that, and do not guess its details.",
      isError: true,
    };
  }

  if (status === 429) {
    return {
      content:
        "The metadata lookup is temporarily rate-limited. Tell the user to try again in a minute.",
      isError: true,
    };
  }

  return {
    content:
      `The metadata lookup failed${status ? ` (status ${status})` : ""}. ` +
      "Tell the user it's a problem on our side, not theirs, and do not guess the document's details.",
    isError: true,
  };
}
