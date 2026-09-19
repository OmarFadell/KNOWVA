import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import type { ILogger } from "@microsoft/teams.common";
import { describeError } from "../errors";

/**
 * Talks to GitHub's hosted MCP server.
 *
 *   https://api.githubcopilot.com/mcp/
 *
 * REMOTE, NOT SELF-HOSTED, AND NOT SHARED. There is a local/containerised
 * github-mcp-server that takes a static PAT at startup. It is rejected here for
 * the same reason a shared PAT is rejected everywhere else in this feature:
 * one credential baked into the process serves every user, so whoever asks
 * inherits whatever that token can reach. The remote server takes the bearer
 * token per request, which lets each query run as the person who asked and lets
 * GitHub apply their real repository permissions. That property is the entire
 * security model -- see src/auth/github.ts.
 *
 * Consequently a client is built PER CALL, around one user's token. These
 * objects are cheap and stateless from our side; a cached client would be a
 * cached credential, and the first thing anyone would do with it is share it
 * across users.
 *
 * WHY THE OFFICIAL SDK RATHER THAN HAND-ROLLED HTTP. Streamable HTTP is not
 * "POST some JSON": it is a JSON-RPC handshake (initialize, then an initialized
 * notification, then calls), responses arrive as either JSON or an SSE stream
 * depending on the server's mood, and there is a session id header to carry
 * between requests. All of that is knowable and all of it is the kind of detail
 * that breaks quietly six months later when the far end changes which response
 * mode it prefers. @modelcontextprotocol/sdk is maintained by the people who
 * define the protocol; this file stays small because of it.
 *
 * Docs:
 *   https://github.com/github/github-mcp-server
 *   https://modelcontextprotocol.io/docs/concepts/transports
 */

const GITHUB_MCP_URL = "https://api.githubcopilot.com/mcp/";

/** Whole-call budget. A Teams turn has a few seconds before the user gives up. */
const CALL_TIMEOUT_MS = 20_000;

/** Identifies Knowva to the MCP server, for its logs and for protocol negotiation. */
const CLIENT_INFO = { name: "knowva", version: "1.0.0" } as const;

/** Why an MCP call could not produce a result. Each maps to distinct user-facing advice. */
export type McpFailure =
  | "unauthorized"
  | "forbidden"
  | "rate-limited"
  | "tool-error"
  | "transport-error";

/**
 * An embedded resource block from a tool response.
 *
 * ===========================================================================
 * THIS IS WHERE get_file_contents ACTUALLY PUTS A FILE, AND WHY IT LOOKED EMPTY
 *
 * github/github-mcp-server#782 reports get_file_contents returning "empty"
 * responses despite correct parameters. Reproduced against a real repository,
 * the response is not empty at all -- it is two blocks:
 *
 *   { type: "text",     text: "successfully downloaded text file (SHA: ...)" }
 *   { type: "resource", resource: { uri, mimeType, text: "<the file>" } }
 *
 * A client that reads only `text` blocks sees the receipt and none of the file,
 * which is indistinguishable from an empty result. That is a client-side bug,
 * and this file had it: flattenContent() rendered every non-text block as "(a
 * resource result, which Knowva cannot read)".
 *
 * Binary files use the same shape with `blob` (base64) in place of `text`, and
 * a real mimeType -- image/png for a PNG. So "is this binary?" is answered by
 * which field is populated, not by guessing from the path extension.
 * ===========================================================================
 */
export interface McpResource {
  /** e.g. repo://owner/name/sha/<sha>/contents/path/to/file */
  uri?: string;
  /** e.g. "text/plain; charset=utf-8", or "image/png" for binary. */
  mimeType?: string;
  /** The decoded text, for text resources. Undefined when binary. */
  text?: string;
  /** True when the payload arrived as base64 `blob` rather than `text`. */
  isBinary: boolean;
  /** Decoded size in bytes, when it could be determined. */
  byteLength?: number;
}

/** One shape with optional fields -- see the note on ActingIdentity for why. */
export interface McpCallResult {
  ok: boolean;
  /** Set when ok: the server's TEXT blocks, flattened. */
  text?: string;
  /**
   * Set when ok: any embedded resource blocks, which is where file contents
   * arrive. Additive -- the search_* tools return text blocks only and are
   * unaffected by this field existing.
   */
  resources?: McpResource[];
  /** Set when not ok. */
  reason?: McpFailure;
  /** Set when not ok, if the far end explained itself. */
  detail?: string;
}

/**
 * Calls one tool on the GitHub MCP server as the given user.
 *
 * `accessToken` must be that user's own GitHub token. Passing anything else --
 * a shared token, another user's token -- silently produces answers scoped to
 * the wrong person, which GitHub cannot detect and neither can this function.
 * The only place that decides which token to pass is src/tools/github.ts, after
 * the identity check in src/auth/acting-identity.ts.
 */
export async function callGitHubMcpTool(
  accessToken: string,
  toolName: string,
  args: Record<string, unknown>,
  log: ILogger
): Promise<McpCallResult> {
  const transport = new StreamableHTTPClientTransport(new URL(GITHUB_MCP_URL), {
    requestInit: {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "X-MCP-Client": CLIENT_INFO.name,
      },
    },
  });

  const client = new Client(CLIENT_INFO, { capabilities: {} });

  try {
    await client.connect(transport);

    log.info(`github-mcp: calling ${toolName}`);
    const response = await client.callTool(
      { name: toolName, arguments: args },
      undefined,
      { timeout: CALL_TIMEOUT_MS }
    );

    const text = flattenContent(response.content);
    const resources = collectResources(response.content);

    if (response.isError) {
      // A tool-level error: the call reached GitHub and GitHub declined it.
      // `text` is the server's own explanation and is usually the most useful
      // thing available -- a bad search qualifier, an unknown repo, and so on.
      log.warn(`github-mcp: ${toolName} returned a tool error -- ${truncate(text, 400)}`);
      return { ok: false, reason: "tool-error", detail: text };
    }

    log.info(
      `github-mcp: ${toolName} returned ${text.length} text chars` +
        (resources.length > 0 ? ` and ${resources.length} resource block(s)` : "")
    );
    return { ok: true, text, resources };
  } catch (err) {
    return classifyTransportError(err, toolName, log);
  } finally {
    // Both closes are best-effort: the answer (or the failure) is already
    // decided by this point, and a noisy teardown must not become the thing the
    // user hears about.
    await client.close().catch(() => undefined);
    await transport.close().catch(() => undefined);
  }
}

/**
 * MCP content blocks to plain text.
 *
 * Resource blocks are deliberately NOT rendered here -- they are returned
 * separately as McpResource, because a caller reading a file needs the content
 * and its mimeType apart from the server's covering note, not glued together.
 * Anything that is neither text nor resource is still named rather than dropped
 * silently, so a future server change shows up as a visible gap instead of an
 * answer that quietly lost half its content.
 */
function flattenContent(content: unknown): string {
  if (!Array.isArray(content)) return "";

  return content
    .map((block) => {
      const b = block as { type?: string; text?: string };
      if (b?.type === "text" && typeof b.text === "string") return b.text;
      if (b?.type === "resource") return "";
      return b?.type ? `(a ${b.type} result, which Knowva cannot read)` : "";
    })
    .filter(Boolean)
    .join("\n")
    .trim();
}

/**
 * The embedded resource blocks of a response.
 *
 * A base64 `blob` is measured, never decoded into a string: the point of
 * detecting binary is to avoid pushing megabytes of it anywhere, and decoding
 * it just to find out how big it is would defeat that. base64 is 4 characters
 * per 3 bytes minus padding, which is exact rather than approximate.
 */
function collectResources(content: unknown): McpResource[] {
  if (!Array.isArray(content)) return [];

  const resources: McpResource[] = [];

  for (const block of content) {
    const b = block as {
      type?: string;
      resource?: { uri?: string; mimeType?: string; text?: string; blob?: string };
    };
    if (b?.type !== "resource" || !b.resource) continue;

    const { uri, mimeType, text, blob } = b.resource;

    if (typeof text === "string") {
      resources.push({ uri, mimeType, text, isBinary: false, byteLength: Buffer.byteLength(text, "utf8") });
    } else if (typeof blob === "string") {
      resources.push({ uri, mimeType, isBinary: true, byteLength: base64ByteLength(blob) });
    } else {
      // A resource with neither payload. Recorded rather than skipped so the
      // caller can say "there was something here I could not read".
      resources.push({ uri, mimeType, isBinary: false });
    }
  }

  return resources;
}

function base64ByteLength(value: string): number {
  const clean = value.replace(/[^A-Za-z0-9+/=]/g, "");
  const padding = clean.endsWith("==") ? 2 : clean.endsWith("=") ? 1 : 0;
  return Math.max(0, Math.floor((clean.length * 3) / 4) - padding);
}

/**
 * Turns a thrown transport/protocol failure into a classified reason.
 *
 * HTTP status is the useful signal and the SDK surfaces it inconsistently
 * across error types, so this looks in several places before falling back to
 * matching the message. Getting 401 apart from 403 matters: one means "sign in
 * again", the other means "your GitHub account cannot see that", and telling a
 * user the wrong one sends them somewhere useless.
 */
function classifyTransportError(err: unknown, toolName: string, log: ILogger): McpCallResult {
  log.error(`github-mcp: ${toolName} failed`, describeError(err), err);

  const status = statusOf(err);
  const message = (err as { message?: string })?.message ?? "";

  if (status === 401 || /\b401\b|unauthorized/i.test(message)) {
    return { ok: false, reason: "unauthorized" };
  }
  if (status === 403 || /\b403\b|forbidden/i.test(message)) {
    return { ok: false, reason: "forbidden" };
  }
  if (status === 429 || /\b429\b|rate limit/i.test(message)) {
    return { ok: false, reason: "rate-limited" };
  }

  return { ok: false, reason: "transport-error", detail: message };
}

function statusOf(err: unknown): number | undefined {
  const e = err as {
    code?: number;
    status?: number;
    statusCode?: number;
    response?: { status?: number };
  };

  // MCP's own JSON-RPC error codes are negative and are not HTTP statuses, so
  // `code` only counts when it looks like one.
  const candidates = [e?.status, e?.statusCode, e?.response?.status, e?.code];
  return candidates.find((n) => typeof n === "number" && n >= 100 && n < 600);
}

function truncate(value: string, max: number): string {
  return value.length > max ? `${value.slice(0, max)}...` : value;
}
