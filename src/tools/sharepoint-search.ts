import { Client as GraphClient } from "@microsoft/teams.graph";
import type { ILogger } from "@microsoft/teams.common";
import { describeError } from "../errors";
import type { AgentTool, AgentToolResult } from "../agent/loop";

/**
 * The `search_documents` tool: answers questions grounded in one SharePoint
 * site, via the Microsoft 365 Copilot Retrieval API.
 *
 *   POST https://graph.microsoft.com/v1.0/copilot/retrieval
 *
 * Retrieval happens inside the tenant -- nothing is downloaded or parsed here.
 * The API is delegated-only and security-trims to what the calling user may
 * see, so the GraphClient passed in must carry that user's token (the same one
 * app.ts already has from Teams SSO). It also requires the caller to hold a
 * Microsoft 365 Copilot license and the delegated Files.Read.All +
 * Sites.Read.All scopes on the `graph` OAuth connection.
 *
 * Docs (read before changing -- this API is new and moving):
 *   https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/retrieval/overview
 *   https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/retrieval/copilotroot-retrieval
 */

// API-imposed constraints. We enforce these ourselves and return a clear error
// rather than letting the request 400.
const MAX_QUERY_LENGTH = 1500;
const MIN_RESULTS = 1;
const MAX_RESULTS = 25;

// How many documents to pull back per search. The API's own guidance is not to
// cap this unless the LLM's token budget demands it -- which, in a Teams reply,
// it does. 10 keeps the grounding context bounded.
const RESULTS_PER_SEARCH = 10;

interface RetrievalExtract {
  text?: string;
  relevanceScore?: number;
}

interface RetrievalHit {
  webUrl?: string;
  extracts?: RetrievalExtract[];
  resourceType?: string;
  resourceMetadata?: Record<string, string>;
}

interface RetrievalResponse {
  retrievalHits?: RetrievalHit[];
}

export function createSharePointSearchTool(
  graph: GraphClient,
  siteUrl: string,
  log: ILogger
): AgentTool {
  return {
    definition: {
      name: "search_documents",
      description:
        "Search the organization's SharePoint documents for a single site and return short " +
        "text extracts with source links. Use this to answer questions about internal " +
        "documents, policies, or reference material. Pass one natural-language question as a " +
        "single sentence. Only this one site is searched.",
      inputSchema: {
        type: "object",
        properties: {
          query: {
            type: "string",
            description:
              "One natural-language question, phrased as a single sentence, " +
              `at most ${MAX_QUERY_LENGTH} characters. Include context-rich keywords; avoid typos.`,
          },
        },
        required: ["query"],
        additionalProperties: false,
      },
    },

    async run(input: Record<string, unknown>): Promise<AgentToolResult> {
      const rawQuery = typeof input.query === "string" ? input.query : "";
      const query = rawQuery.trim();

      const constraintError = validateQuery(query);
      if (constraintError) {
        return { content: constraintError, isError: true };
      }

      const body = {
        queryString: query,
        dataSource: "sharePoint",
        // Site scoping. The Retrieval API takes a KQL filterExpression; `path:`
        // restricts results to documents under this site URL.
        filterExpression: `path:"${siteUrl}"`,
        resourceMetadata: ["title", "author"],
        maximumNumberOfResults: clampResults(RESULTS_PER_SEARCH),
      };

      log.info(
        `sharepoint-search: querying Retrieval API (site="${siteUrl}" queryChars=${query.length})`
      );

      let hits: RetrievalHit[];
      try {
        const response = await graph.http.post<RetrievalResponse>(
          "/copilot/retrieval",
          body
        );
        hits = response.data.retrievalHits ?? [];
      } catch (err) {
        return classifyError(err, log);
      }

      if (hits.length === 0) {
        // Not an error -- the API found nothing relevant. The model must say so
        // plainly rather than inventing an answer.
        log.info("sharepoint-search: no results");
        return {
          content:
            `The document search returned no results for that query in ${siteUrl}. ` +
            `Tell the user the site's documents don't appear to cover this, and do not answer from your own knowledge as if they did.`,
        };
      }

      log.info(`sharepoint-search: ${hits.length} document(s) returned`);
      return { content: formatHits(hits) };
    },
  };
}

/** Returns an error message if the query breaks an API constraint, else null. */
function validateQuery(query: string): string | null {
  if (!query) {
    return "No query was provided. Call search_documents again with a single-sentence question.";
  }
  if (query.length > MAX_QUERY_LENGTH) {
    return (
      `That query is ${query.length} characters; the document search accepts at most ${MAX_QUERY_LENGTH}. ` +
      "Shorten it to one focused sentence and try again."
    );
  }
  // "Single sentence" is a Retrieval API best practice. A period/question
  // mark/exclamation followed by whitespace and more text is the clearest sign
  // of more than one sentence. (An abbreviation like "U.S. " can trip this; the
  // cost is one clarifying retry, not a bad answer.)
  const withoutTrailingPunctuation = query.replace(/[.!?]+\s*$/, "");
  if (/[.!?]\s+\S/.test(withoutTrailingPunctuation)) {
    return (
      "The document search takes one sentence at a time. " +
      "Combine this into a single question and call search_documents again."
    );
  }
  return null;
}

function clampResults(n: number): number {
  return Math.min(MAX_RESULTS, Math.max(MIN_RESULTS, n));
}

/**
 * Renders hits as plain text for the model. Each source is numbered and carries
 * its webUrl so the model can cite it with a link.
 */
function formatHits(hits: RetrievalHit[]): string {
  const blocks = hits.map((hit, i) => {
    const title = hit.resourceMetadata?.title?.trim();
    const author = hit.resourceMetadata?.author?.trim();
    const url = hit.webUrl ?? "(no link available)";
    const heading = title ? `${title} -- ${url}` : url;
    const byline = author ? ` (author: ${author})` : "";

    const extracts = (hit.extracts ?? [])
      .map((extract) => extract.text?.trim())
      .filter((text): text is string => Boolean(text))
      .map((text) => `  - ${text}`)
      .join("\n");

    return `[${i + 1}] ${heading}${byline}\n${extracts || "  (no extract text)"}`;
  });

  return (
    "Document search results. Ground your answer only in these extracts and cite each " +
    "claim with its source link in Markdown. If they do not answer the question, say so.\n\n" +
    blocks.join("\n\n")
  );
}

/**
 * Turns a failed Retrieval API call into a result the model can relay in plain
 * language. A user without a Copilot license must get a clear sentence, never a
 * stack trace.
 */
function classifyError(err: unknown, log: ILogger): AgentToolResult {
  log.error("sharepoint-search: Retrieval API call failed", describeError(err), err);

  const e = err as { response?: { status?: number; data?: unknown } };
  const status = e?.response?.status;
  const bodyText = safeStringify(e?.response?.data);

  if (status === 401) {
    return {
      content:
        "The document search could not authenticate. Tell the user to sign out and sign back in, then retry.",
      isError: true,
    };
  }

  if (status === 403) {
    if (/licen[sc]e/i.test(bodyText)) {
      return {
        content:
          "Document search is unavailable: this account does not have a Microsoft 365 Copilot license, " +
          "which this feature requires. Tell the user plainly and suggest they contact their administrator.",
        isError: true,
      };
    }
    return {
      content:
        "Document search was denied (permission error). An administrator likely needs to grant consent for " +
        "the Files.Read.All and Sites.Read.All permissions. Tell the user it isn't set up yet.",
      isError: true,
    };
  }

  if (status === 429) {
    return {
      content:
        "Document search is temporarily rate-limited. Tell the user to try again in a minute.",
      isError: true,
    };
  }

  return {
    content:
      `The document search service returned an error${status ? ` (status ${status})` : ""}. ` +
      "Tell the user it's a problem on our side, not theirs.",
    isError: true,
  };
}

function safeStringify(value: unknown): string {
  if (value == null) return "";
  if (typeof value === "string") return value;
  try {
    return JSON.stringify(value);
  } catch {
    return "";
  }
}
