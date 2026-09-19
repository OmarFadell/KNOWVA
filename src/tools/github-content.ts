import type { ILogger } from "@microsoft/teams.common";
import type { AgentTool, AgentToolResult } from "../agent/loop";
import type { ActingIdentityResolver } from "../auth/acting-identity";
import { callGitHubMcpTool, type McpResource } from "../github/mcp-client";
import { mcpFailureResult, resolveGitHubToken } from "./github-access";

/**
 * The `get_repo_content` tool: browse a repository's tree, and read a file.
 *
 * Deliberately a separate tool from search_github, mirroring the split between
 * search_documents and get_document_metadata, and between search_emails and
 * get_email_content. The pattern is the same each time: search finds candidates
 * cheaply, a second tool pays for detail only once the model has decided which
 * thing matters. It buys something extra here -- keyword search cannot answer
 * "how is this repo laid out?", and guessing file paths is how a model ends up
 * confidently describing a file that does not exist.
 *
 * ONE MCP TOOL, TWO BEHAVIOURS. `get_file_contents` returns a directory listing
 * for a directory path (root included) and file content for a file path. Rather
 * than exposing that as two Knowva tools, this presents it as one and formats
 * whichever came back -- the model does not have to know in advance which a
 * path is, which is the entire point when it is exploring.
 *
 * No new permissions. Contents (read) on the existing GitHub App already covers
 * this, and it runs on the same per-user token as search_github, through the
 * same guard in src/tools/github-access.ts.
 *
 * ---------------------------------------------------------------------------
 * VERIFIED AGAINST THE LIVE SERVER, not assumed (github/github-mcp-server#782
 * reports "empty" responses here). What actually comes back:
 *
 *   directory -> one text block: a JSON array of
 *                { type: "file"|"dir", name, path, size, sha, ...urls }
 *   text file -> a text block ("successfully downloaded text file (SHA: ...)")
 *                PLUS a resource block holding the content
 *   binary    -> the same, but the resource carries base64 `blob` and a real
 *                mimeType (image/png) instead of `text`
 *
 * The "empty response" is a client that reads only text blocks. See the header
 * of src/github/mcp-client.ts, which is where that was fixed.
 *
 * Also verified: `path` may be omitted (defaults to root), `ref` accepts a bare
 * branch name as happily as refs/heads/<branch>, and `fields` trims directory
 * listings -- the root of a real repo dropped from ~13,000 characters to a
 * fraction of that.
 * ---------------------------------------------------------------------------
 */

const MCP_TOOL = "get_file_contents";

/**
 * Fields requested for directory entries.
 *
 * The full listing repeats four URLs per entry -- api, git, html and download
 * -- each carrying a 40-character SHA. On a real repository root that was most
 * of ~13KB of response, none of which the model can use: it navigates by path,
 * and a citation is built from owner/repo/path. The server ignores this for
 * single files, so it is safe to send on every call.
 */
const DIRECTORY_FIELDS = ["type", "name", "path", "size"];

/** Entries listed per directory. Deep trees get truncated rather than flooding the turn. */
const MAX_ENTRIES = 200;

/**
 * Cap on file content handed to the model. Roughly a long source file. Beyond
 * this the model is better served by search_github finding the relevant lines.
 */
const MAX_FILE_CHARS = 12_000;

interface DirectoryEntry {
  type?: string;
  name?: string;
  path?: string;
  size?: number;
}

export function createGitHubContentTool(
  identity: ActingIdentityResolver,
  log: ILogger
): AgentTool {
  return {
    definition: {
      name: "get_repo_content",
      description:
        "Browse a GitHub repository's files, or read one file. Give a directory path (or none " +
        "at all) to list what's in it; give a file path to read that file. Only repositories " +
        "the person you are talking to can already see are reachable. Use this to explore how a " +
        "repository is laid out and to read specific files -- start at the root and work down " +
        "rather than guessing paths. Use search_github instead when you are looking for a " +
        "keyword across many repositories.",
      inputSchema: {
        type: "object",
        properties: {
          owner: {
            type: "string",
            description: "Repository owner: a GitHub username or organization, e.g. 'contoso'.",
          },
          repo: {
            type: "string",
            description: "Repository name on its own, without the owner, e.g. 'billing-api'.",
          },
          path: {
            type: "string",
            description:
              "Path within the repository. Omit it (or pass an empty string) for the root " +
              "listing. A directory path lists its contents; a file path reads the file. No " +
              "leading slash, e.g. 'src/auth' or 'src/auth/token.ts'.",
          },
          branch: {
            type: "string",
            description:
              "Optional branch, tag, or commit SHA. Omit it for the repository's default " +
              "branch, which is almost always what you want.",
          },
        },
        required: ["owner", "repo"],
        additionalProperties: false,
      },
    },

    async run(input: Record<string, unknown>): Promise<AgentToolResult> {
      const owner = str(input.owner);
      const repo = str(input.repo);

      if (!owner || !repo) {
        return {
          content:
            "Both owner and repo are required. Call get_repo_content again with, for example, " +
            "owner 'contoso' and repo 'billing-api'. If you don't know the repository, use " +
            "search_github with scope 'repositories' to find it first.",
          isError: true,
        };
      }

      // Models reliably pass "owner/repo" as the repo when a path is on their
      // mind. Splitting it is kinder than a lecture, and unambiguous: a repo
      // name cannot contain a slash.
      const normalizedRepo = repo.includes("/") ? repo.split("/").pop() ?? repo : repo;

      // Leading slashes are the other habitual slip. The server treats "/src"
      // as a miss, which would surface as "path not found" for a path that
      // plainly exists.
      const path = str(input.path).replace(/^\/+/, "");
      const branch = str(input.branch);

      // Same guard as search_github: identity resolved from the token, then
      // that user's session. Nothing below this line runs without both.
      const auth = await resolveGitHubToken(identity, log);
      if (!auth.ok) return auth.refusal as AgentToolResult;

      const args: Record<string, unknown> = {
        owner,
        repo: normalizedRepo,
        path,
        fields: DIRECTORY_FIELDS,
      };
      // A bare branch name works as well as refs/heads/<name> (verified), so it
      // is passed through untouched -- which also lets a tag or commit SHA go
      // through the same argument without special-casing.
      if (branch) args.ref = branch;

      log.info(
        `github-content: ${owner}/${normalizedRepo} path="${path || "(root)"}"` +
          (branch ? ` ref="${branch}"` : "")
      );

      const result = await callGitHubMcpTool(auth.token as string, MCP_TOOL, args, log);

      if (!result.ok) {
        if (result.reason === "tool-error") {
          return toolErrorResult(result.detail, owner, normalizedRepo, path, branch);
        }
        return mcpFailureResult(result.reason, result.detail, "that repository content");
      }

      const where = describeLocation(owner, normalizedRepo, path, branch);

      // A resource block means a file came back. Directories never produce one.
      const resource = (result.resources ?? [])[0];
      if (resource) {
        return { content: formatFile(resource, where, path) };
      }

      const entries = parseDirectory(result.text);
      if (entries) {
        return { content: formatDirectory(entries, where, path) };
      }

      // Neither a resource nor a parseable listing. Rather than guess, hand the
      // server's own words over and say plainly that they were not understood.
      log.warn(
        `github-content: unrecognised response shape for ${where} -- ` +
          `${(result.text ?? "").slice(0, 200)}`
      );
      return {
        content:
          `GitHub returned something for ${where} that Knowva could not interpret as either a ` +
          `file or a directory listing. The raw response was:\n\n${result.text || "(empty)"}\n\n` +
          "Do not invent contents for this path.",
      };
    },
  };
}

function str(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}

function describeLocation(
  owner: string,
  repo: string,
  path: string,
  branch: string
): string {
  const location = `${owner}/${repo}/${path || "(repository root)"}`;
  return branch ? `${location} on ${branch}` : location;
}

/** The directory listing arrives as a JSON array in a text block. */
function parseDirectory(text: string | undefined): DirectoryEntry[] | null {
  if (!text) return null;

  const trimmed = text.trim();
  if (!trimmed.startsWith("[")) return null;

  try {
    const parsed = JSON.parse(trimmed);
    return Array.isArray(parsed) ? (parsed as DirectoryEntry[]) : null;
  } catch {
    return null;
  }
}

/**
 * Renders a directory listing.
 *
 * Directories first and each marked with a trailing slash, so the model can see
 * at a glance where it can go next -- which is the whole purpose of listing a
 * directory rather than searching it.
 */
function formatDirectory(
  entries: DirectoryEntry[],
  where: string,
  path: string
): string {
  if (entries.length === 0) {
    return (
      `${where} is a directory, and it is empty. Tell the user there's nothing in it, and do ` +
      "not guess at what it might contain."
    );
  }

  const shown = entries.slice(0, MAX_ENTRIES);
  const directories = shown.filter((e) => e.type === "dir");
  const files = shown.filter((e) => e.type !== "dir");

  const lines = [
    ...directories.map((e) => `  dir   ${e.name ?? "(unnamed)"}/    path: ${e.path ?? ""}`),
    ...files.map(
      (e) =>
        `  file  ${e.name ?? "(unnamed)"}    path: ${e.path ?? ""}` +
        (typeof e.size === "number" ? `  (${formatSize(e.size)})` : "")
    ),
  ];

  const truncated =
    entries.length > shown.length
      ? `\n\n${entries.length - shown.length} further entries were not listed.`
      : "";

  return (
    `Directory listing for ${where}:\n\n` +
    lines.join("\n") +
    truncated +
    "\n\nThese are the only entries at this level -- subdirectory contents are NOT included. " +
    "To go deeper, call get_repo_content again with one of the paths above, exactly as written. " +
    "Do not describe or infer the contents of any file listed here: you have seen its name and " +
    "size, nothing more." +
    (path ? "" : " This is the repository root.")
  );
}

/**
 * Renders one file.
 *
 * Binary is detected from which field the server populated, not from the file
 * extension -- see the header of src/github/mcp-client.ts. Saying so plainly
 * beats dumping base64 into the model's context, which would be several
 * thousand tokens of nothing.
 */
function formatFile(resource: McpResource, where: string, path: string): string {
  const kind = resource.mimeType ? ` (${resource.mimeType})` : "";

  if (resource.isBinary) {
    return (
      `${where} is a binary file${kind}, ${formatSize(resource.byteLength ?? 0)}. Knowva does ` +
      "not read binary files, so its contents are NOT available. Tell the user it's a binary " +
      "file you can't read, and do not speculate about what it contains beyond what its name " +
      "and type suggest."
    );
  }

  if (resource.text === undefined) {
    return (
      `GitHub returned ${where} but with no readable content. Tell the user you couldn't read ` +
      "that file, and do not invent its contents."
    );
  }

  const truncated = resource.text.length > MAX_FILE_CHARS;
  const body = truncated ? `${resource.text.slice(0, MAX_FILE_CHARS)}\n...(truncated)` : resource.text;

  const notes = [
    `Contents of ${where}${kind}.`,
    truncated
      ? "This file was TRUNCATED because it is long -- do not claim to have read all of it, and " +
        "do not infer what the cut-off part contains. Use search_github if you need a specific " +
        "part of it."
      : "",
    "Cite it by repository and path when you use it. Describe only what is actually below.",
  ].filter(Boolean);

  return `${notes.join(" ")}\n\n--- ${path || "file"} ---\n${body}\n--- end of file ---`;
}

/**
 * Classifies a tool-level error from the MCP server.
 *
 * These come back as prose, so the matching is on GitHub's own wording --
 * captured from live calls rather than guessed. The three cases have three
 * different fixes, and telling a user the wrong one sends them somewhere
 * useless: a wrong path is theirs to correct, a missing repository may be a
 * permissions problem, and a bad branch is neither.
 */
function toolErrorResult(
  detail: string | undefined,
  owner: string,
  repo: string,
  path: string,
  branch: string
): AgentToolResult {
  const message = detail ?? "";

  // "failed to resolve git reference: failed to get repository info: GET
  //  https://api.github.com/repos/<owner>/<repo>: 404 Not Found []"
  if (/repository info/i.test(message)) {
    return {
      content:
        `No repository ${owner}/${repo} could be reached. Either it doesn't exist, or this ` +
        "user's GitHub account can't see it -- from here those look identical and you should " +
        "say so rather than picking one. Suggest they check the owner and name, and that they " +
        "have access. Do not guess at its contents.",
      isError: true,
    };
  }

  // Two different wordings depending on what was passed, both observed live:
  //   ref "refs/heads/nope" -> "failed to get final reference for
  //                             \"refs/heads/nope\": GET .../git/ref/heads/nope: 404"
  //   ref "nope"            -> "could not resolve ref \"nope\" as a branch or a tag"
  // Matching only the first is how the second silently fell through to the
  // generic message -- which is exactly what happened before this was tested.
  if (/final reference|git\/ref|could not resolve ref/i.test(message)) {
    return {
      content:
        `The branch, tag or commit '${branch}' doesn't exist in ${owner}/${repo}. Tell the user, ` +
        "and offer to look on the default branch instead by calling get_repo_content again " +
        "without a branch.",
      isError: true,
    };
  }

  // "Failed to get file contents. The path does not point to a file or
  //  directory, or the file does not exist in the repository."
  if (/does not point to a file or directory|does not exist in the repository/i.test(message)) {
    return {
      content:
        `There is nothing at '${path || "(root)"}' in ${owner}/${repo}` +
        (branch ? ` on ${branch}` : "") +
        ". The path is wrong, or it was removed. Do NOT guess another path: list the parent " +
        "directory with get_repo_content and navigate from what's actually there. If you were " +
        "guessing at this path, say so to the user rather than presenting it as a finding.",
      isError: true,
    };
  }

  return {
    content:
      `GitHub couldn't return ${owner}/${repo}/${path || "(root)"}. GitHub said: ` +
      `${message || "(no detail given)"}. Relay that in plain language, and do not guess at ` +
      "the content.",
    isError: true,
  };
}

function formatSize(size: number): string {
  if (!Number.isFinite(size) || size < 0) return "unknown size";
  if (size < 1024) return `${size} bytes`;
  const units = ["KB", "MB", "GB"];
  let value = size / 1024;
  let unit = 0;
  while (value >= 1024 && unit < units.length - 1) {
    value /= 1024;
    unit++;
  }
  return `${value.toFixed(1)} ${units[unit]}`;
}
