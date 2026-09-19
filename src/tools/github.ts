import type { ILogger } from "@microsoft/teams.common";
import type { AgentTool, AgentToolResult } from "../agent/loop";
import type { ActingIdentityResolver } from "../auth/acting-identity";
import { callGitHubMcpTool } from "../github/mcp-client";
import { mcpFailureResult, resolveGitHubToken } from "./github-access";

/**
 * The `search_github` tool: code, repositories, issues and pull requests, as
 * seen by the person asking.
 *
 * ===========================================================================
 * THE SECURITY PROPERTY THIS FILE IS RESPONSIBLE FOR
 *
 * Knowva holds a GitHub token for every user who has signed in. This tool picks
 * which one to spend. That is the whole of the confused-deputy defence, and
 * unlike the email tools there is no second line:
 *
 *   - For email, `/me/messages` runs on the asking user's delegated Graph
 *     token. Pick the wrong person and Graph still refuses -- the token simply
 *     cannot open another mailbox.
 *   - For GitHub, the token IS the authority. Send Bob's token and GitHub
 *     answers as Bob, cheerfully and correctly, because from GitHub's side
 *     nothing is wrong. Nobody downstream can catch the mistake.
 *
 * So the identity resolution below is not a formality and must not be
 * "optimised" into reading `activity.from` directly. It resolves the acting
 * user from the Graph token itself and cross-checks the Teams-supplied id (see
 * src/auth/acting-identity.ts), and refuses on any doubt -- including the case
 * where the two disagree, which should be impossible and is therefore exactly
 * the case worth refusing on.
 *
 * In a group chat this is the difference that matters. Alice @mentions Knowva;
 * the turn runs on Alice's token; Alice's GitHub session is the only one that
 * may be used, regardless of who else is in the room or who signed in most
 * recently.
 * ===========================================================================
 *
 * NOT SIGNED IN IS NOT AN ERROR. When there is no usable session the tool
 * returns a `needs-auth` signal alongside its prose, and app.ts turns that into
 * an Adaptive Card carrying a one-time authorize URL. See AgentToolSignal in
 * src/agent/loop.ts for why this one case needs a channel that bypasses the
 * model.
 */

/** Maps the `scope` argument to the GitHub MCP server's tool names. */
const SCOPE_TO_MCP_TOOL: Record<string, string> = {
  code: "search_code",
  repositories: "search_repositories",
  issues: "search_issues",
  pull_requests: "search_pull_requests",
};

const SCOPES = Object.keys(SCOPE_TO_MCP_TOOL);

/** Results per search. GitHub search payloads are verbose; this keeps a turn affordable. */
const RESULTS_PER_SEARCH = 10;

/** Longest query accepted, before GitHub would reject it anyway. */
const MAX_QUERY_LENGTH = 400;

/**
 * Cap on the MCP payload handed to the model. GitHub's search responses embed
 * whole file fragments and full issue bodies; without a bound, one broad query
 * can crowd out the model's room to actually answer.
 */
const MAX_RESULT_CHARS = 12_000;

export function createGitHubSearchTool(
  identity: ActingIdentityResolver,
  log: ILogger
): AgentTool {
  return {
    definition: {
      name: "search_github",
      description:
        "Search GitHub for code, repositories, issues, or pull requests. Only repositories the " +
        "person you are talking to can already see are searched -- GitHub applies their own " +
        "permissions, so this can never reveal a private repository they lack access to. Use it " +
        "when the user asks about code, a repository, an issue, or a PR. Returns structured " +
        "results including repository names, file paths, and issue/PR numbers and URLs.",
      inputSchema: {
        type: "object",
        properties: {
          query: {
            type: "string",
            description:
              "A GitHub search query. Supports GitHub's own qualifiers, e.g. " +
              "'repo:owner/name retry logic', 'org:contoso language:typescript parseToken', " +
              "'is:open label:bug auth'. Keep it to the distinctive terms; at most " +
              `${MAX_QUERY_LENGTH} characters.`,
          },
          scope: {
            type: "string",
            enum: SCOPES,
            description:
              "What to search. 'code' for source files (the default, and the right choice for " +
              "'where is X implemented'), 'repositories' to find a repo by name or topic, " +
              "'issues' for bug reports and discussions, 'pull_requests' for changes and reviews.",
          },
        },
        required: ["query"],
        additionalProperties: false,
      },
    },

    async run(input: Record<string, unknown>): Promise<AgentToolResult> {
      const query = (typeof input.query === "string" ? input.query : "").trim();
      if (!query) {
        return {
          content:
            "No search terms were provided. Call search_github again with the keywords to look for.",
          isError: true,
        };
      }
      if (query.length > MAX_QUERY_LENGTH) {
        return {
          content:
            `That query is ${query.length} characters; GitHub search accepts at most ` +
            `${MAX_QUERY_LENGTH}. Shorten it to the distinctive terms and try again.`,
          isError: true,
        };
      }

      const scope = normalizeScope(input.scope);
      if (!scope) {
        return {
          content:
            `"${String(input.scope)}" is not a valid scope. Use one of: ${SCOPES.join(", ")}.`,
          isError: true,
        };
      }

      // Whose GitHub account is this query running as? Everything above is
      // argument validation; nothing below runs without an established
      // identity and that user's own session. See src/tools/github-access.ts.
      const auth = await resolveGitHubToken(identity, log);
      if (!auth.ok) return auth.refusal as AgentToolResult;

      // --- Run the search as that user -------------------------------------
      const mcpTool = SCOPE_TO_MCP_TOOL[scope];
      const result = await callGitHubMcpTool(
        auth.token as string,
        mcpTool,
        { query, perPage: RESULTS_PER_SEARCH },
        log
      );

      if (!result.ok) {
        if (result.reason === "tool-error") {
          // Almost always a malformed query -- a bad qualifier, an unknown
          // repo. GitHub's own message is the most useful thing available, so
          // it is passed through for the model to translate rather than
          // swallowed.
          return {
            content:
              `GitHub rejected that ${scope} search. GitHub said: ` +
              `${result.detail || "(no detail given)"}. If that looks like a query-syntax ` +
              "problem, fix the query and call search_github once more. Otherwise tell the user " +
              "what GitHub said, in plain language.",
            isError: true,
          };
        }
        return mcpFailureResult(result.reason, result.detail, `that ${scope} search`);
      }

      if (!result.text) {
        return {
          content:
            `GitHub returned no results for "${query}" (${scope}). Tell the user nothing matched, ` +
            "and remind them this only covers repositories their own GitHub account can see. Do " +
            "NOT answer from your own knowledge as though you had found something.",
        };
      }

      return { content: formatResults(result.text as string, query, scope, auth.login) };
    },
  };
}

function normalizeScope(raw: unknown): string | null {
  if (raw === undefined || raw === null || raw === "") return "code";
  if (typeof raw !== "string") return null;
  const scope = raw.trim().toLowerCase();
  return scope in SCOPE_TO_MCP_TOOL ? scope : null;
}

/**
 * Wraps the MCP server's payload with the citation rules.
 *
 * The payload is GitHub's own JSON and is passed through rather than reshaped:
 * it already carries repository names, file paths, issue numbers and html_url
 * values, which is exactly what a citation needs, and re-serialising it here
 * would be one more place for a URL to get mangled.
 */
function formatResults(
  payload: string,
  query: string,
  scope: string,
  login: string | undefined
): string {
  const truncated = payload.length > MAX_RESULT_CHARS;
  const body = truncated ? `${payload.slice(0, MAX_RESULT_CHARS)}\n...(truncated)` : payload;

  const caveats = [
    `GitHub ${scope} search results for "${query}"` +
      (login ? `, run as the GitHub account ${login}.` : "."),
    "These cover only repositories this user can already see.",
    truncated
      ? "The payload was truncated -- do not claim to have seen every result, and do not infer " +
        "what was cut off."
      : "",
  ].filter(Boolean);

  return (
    caveats.join(" ") +
    "\n\n" +
    "Cite everything you take from this. For code, name the repository and the file path. For " +
    "an issue or pull request, give its number and title and link its html_url in Markdown. " +
    "Never state that code or an issue exists without pointing at where you saw it, and never " +
    "describe a file's contents beyond what appears below.\n\n" +
    body
  );
}
