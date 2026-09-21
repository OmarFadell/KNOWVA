import type { ILogger } from "@microsoft/teams.common";
import type { AgentTool, AgentToolResult } from "../agent/loop";
import type { ActingIdentityResolver } from "../auth/acting-identity";
import { getAtlassianIdentity, isAtlassianConfigured } from "../auth/atlassian";
import type { AtlassianIdentity } from "../auth/atlassian-sessions";
import {
  discoverXrayProjects,
  isXrayConfigured,
  searchTestsByJql,
  type XrayFailure,
  type XrayTestMatch,
} from "../xray/client";

/**
 * The `find_my_xray_items` tool: Xray Tests assigned to, or mentioning, the
 * person asking.
 *
 * ===========================================================================
 * TWO CREDENTIALS, TWO QUESTIONS, AND THE SPLIT IS THE DESIGN:
 *
 *     The user's Atlassian OAuth answers "WHO is asking."
 *     The shared Xray key answers   "WHAT can be searched."
 *
 * The OAuth half runs once per user, resolves their real Atlassian accountId
 * via /myself, and discards the token immediately -- see src/auth/atlassian.ts.
 * Nothing in this file holds or spends a user credential. Every search below
 * goes out on the shared Xray service credential.
 *
 * WHY IDENTITY STILL NEEDS OAUTH WHEN THE SEARCH DOES NOT USE IT: matching a
 * Teams user to a Jira user by comparing display names or email addresses
 * across two directories is the kind of almost-right that hands one person
 * another person's items. An accountId established by the user themselves is
 * the only truthful link, and it is what the `assignee =` clause below depends
 * on.
 *
 * WHAT THE RESULTS ARE, STATED PRECISELY, because it is easy to overclaim:
 * these are items naming this user, drawn from what the SHARED key can see. It
 * is not "your Jira" and it is not scoped to the asker's own Jira permissions.
 * For most people the two coincide; they diverge for anybody whose Jira access
 * is wider than the service account's. The text handed to the model says so.
 * ===========================================================================
 *
 * ---------------------------------------------------------------------------
 * TESTS ONLY. DEFECTS ARE NOT REACHABLE, AND NOT FOR WANT OF A QUERY SHAPE.
 *
 * Xray's GraphQL has no general Jira issue search: getTests(jql:) intersects
 * the JQL with Xray Tests and can only return Tests, getCoverableIssues returns
 * requirements-type issues for coverage mapping, and defects exist in Xray only
 * as links hung off Test Runs.
 *
 * The decisive part is the credential, not the API surface: an Xray API key
 * authenticates to xray.cloud.getxray.app and is NOT a Jira credential, so it
 * cannot call /rest/api/3/search on api.atlassian.com at all. Defects are
 * ordinary Jira issues, so finding them means holding a Jira token for
 * somebody -- exactly what the identity-only design gives up.
 *
 * So Defects are reported as not-yet-supported, with the reason, rather than
 * silently omitted. Supporting them later is a deliberate re-opening of the
 * token question, not an incremental addition.
 * ---------------------------------------------------------------------------
 *
 * WHY TWO SEARCHES RATHER THAN ONE. The tool has to say whether each item
 * matched because it is ASSIGNED to the user or because it MENTIONS them, and
 * one combined query cannot tell you which clause fired. So the assignee query
 * and the text query run separately and the results are merged, which makes
 * "both" an observable state rather than a guess.
 */

/** Results to request per search. Xray caps `limit` at 100. */
const RESULTS_PER_SEARCH = 50;

/** How each item came to be in the result set. */
type MatchReason = "assigned" | "mentioned" | "both";

interface FoundItem {
  key: string;
  summary?: string;
  reason: MatchReason;
  assigneeName?: string;
  url?: string;
}

export function createFindMyXrayItemsTool(
  identity: ActingIdentityResolver,
  log: ILogger
): AgentTool {
  return {
    definition: {
      name: "find_my_xray_items",
      description:
        "Find Xray test cases that are assigned to the person you are talking to, or that " +
        "mention them by name or email in the summary, description or comments. Use it when " +
        "they ask what tests are assigned to them, what test work involves them, or what names " +
        "them in Xray. Each result says whether it is assigned to them, mentions them, or both. " +
        "It covers Xray Tests only -- it cannot find Defects or other Jira issue types. It " +
        "returns issue keys, summaries and links, never test steps or test content.",
      inputSchema: {
        type: "object",
        properties: {},
        additionalProperties: false,
      },
    },

    async run(): Promise<AgentToolResult> {
      if (!isAtlassianConfigured()) return atlassianNotConfiguredResult();
      if (!isXrayConfigured()) return xrayNotConfiguredResult();

      // WHOSE items are we looking for? Resolved from the Graph token and
      // cross-checked against the id Teams put on the activity, exactly as the
      // GitHub tools do (see src/auth/acting-identity.ts). This matters as much
      // here as it does there: the whole tool is a lookup keyed on this value,
      // so a wrong key returns another person's items under this person's name.
      const acting = await identity.resolve();
      if (!acting.ok) {
        log.warn(`find-my-xray: refusing -- acting identity ${acting.reason}`);
        return identityRefusalResult(acting.reason);
      }

      const atlassian = getAtlassianIdentity(acting.userId as string);
      if (!atlassian) return needsAuthResult();

      // --- Which projects can be searched at all? --------------------------
      // Reuses the discovery walk that backs list_xray_projects, so both tools
      // agree on what "the projects Knowva can see" means. Bounding the JQL by
      // project also keeps it well under Xray's 100-issue ceiling.
      const discovery = await discoverXrayProjects(log);
      if (!discovery.ok) return xrayFailureResult(discovery.reason, discovery.detail);

      const projectKeys = (discovery.projects || [])
        .map((p) => p.key)
        .filter((key): key is string => Boolean(key));

      if (projectKeys.length === 0) return noProjectsResult();

      // --- Two searches, so "assigned" and "mentioned" stay distinguishable -
      const projectClause = `project in (${projectKeys.map(quoteJql).join(", ")})`;

      const assignedJql = `${projectClause} AND assignee = ${quoteJql(atlassian.accountId)}`;
      const assigned = await searchTestsByJql(assignedJql, RESULTS_PER_SEARCH, log);
      if (!assigned.ok) return xrayFailureResult(assigned.reason, assigned.detail, assignedJql);

      const mentionClause = buildMentionClause(atlassian);
      let mentioned: XrayTestMatch[] = [];
      let mentionTruncated = false;

      if (mentionClause) {
        const mentionJql = `${projectClause} AND (${mentionClause})`;
        const result = await searchTestsByJql(mentionJql, RESULTS_PER_SEARCH, log);
        if (!result.ok) return xrayFailureResult(result.reason, result.detail, mentionJql);
        mentioned = result.matches || [];
        mentionTruncated = (result.total || 0) > mentioned.length;
      } else {
        // No display name and no email means there is nothing to text-match on.
        // The assignee half still works, so this degrades rather than fails.
        log.warn(
          `find-my-xray: no display name or email for ${atlassian.accountId}; ` +
            "searching by assignee only"
        );
      }

      const items = mergeMatches(assigned.matches || [], mentioned);
      const truncated = ((assigned.total || 0) > (assigned.matches || []).length) || mentionTruncated;

      const found = items.map((item) => ({
        ...item,
        url: issueUrl(atlassian.siteUrl, item.key),
      }));

      printToConsole(found, atlassian, projectKeys, truncated);

      if (found.length === 0) return noMatchesResult(atlassian, projectKeys);

      return { content: formatForModel(found, atlassian, projectKeys, truncated) };
    },
  };
}

/**
 * The "mentions me" half of the match.
 *
 * `text ~ "..."` is Jira's master text field: it covers summary, description,
 * environment, comments and text custom fields in one clause, which is exactly
 * the scope this milestone asked for and is why the clause is not spelled out
 * field by field.
 *
 * THE EMAIL CLAUSE IS CONDITIONAL, and that is not defensive coding for its own
 * sake. Atlassian omits emailAddress from /myself unless profile visibility
 * allows it, and many tenants hide it by default -- so for a large share of
 * users this clause simply cannot be built, and the search has to be correct
 * without it. Returns null when there is nothing to match on at all.
 */
function buildMentionClause(identity: AtlassianIdentity): string | null {
  const terms: string[] = [];
  if (identity.displayName) terms.push(`text ~ ${quoteJql(identity.displayName)}`);
  if (identity.emailAddress) terms.push(`text ~ ${quoteJql(identity.emailAddress)}`);
  return terms.length > 0 ? terms.join(" OR ") : null;
}

/**
 * Quotes a value for JQL.
 *
 * Display names contain apostrophes and occasionally quotes, and a project key
 * should never contain either -- but this is string concatenation into a query
 * language, so every interpolated value goes through here regardless of how
 * safe it looks today. Backslashes are escaped first, or escaping the quotes
 * would then be undone by the backslash pass.
 */
function quoteJql(value: string): string {
  return `"${value.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
}

/**
 * Merges the two result sets, promoting anything in both to "both".
 *
 * The assignee check is re-derived from the row rather than assumed from which
 * query returned it: a row can come back from the mention search AND be
 * assigned to the user, and reporting that as merely "mentioned" would understate
 * it.
 */
function mergeMatches(assigned: XrayTestMatch[], mentioned: XrayTestMatch[]): FoundItem[] {
  const byKey = new Map<string, FoundItem>();

  for (const match of assigned) {
    byKey.set(match.key, {
      key: match.key,
      summary: match.summary,
      assigneeName: match.assigneeName,
      reason: "assigned",
    });
  }

  for (const match of mentioned) {
    const existing = byKey.get(match.key);
    if (existing) {
      existing.reason = "both";
      continue;
    }
    byKey.set(match.key, {
      key: match.key,
      summary: match.summary,
      assigneeName: match.assigneeName,
      reason: "mentioned",
    });
  }

  // Assigned first: "this is yours" is a stronger claim on the reader's
  // attention than "your name appears somewhere in this".
  const rank: Record<MatchReason, number> = { both: 0, assigned: 1, mentioned: 2 };
  return [...byKey.values()].sort(
    (a, b) => rank[a.reason] - rank[b.reason] || a.key.localeCompare(b.key)
  );
}

function issueUrl(siteUrl: string | undefined, key: string): string | undefined {
  if (!siteUrl) return undefined;
  return `${siteUrl.replace(/\/+$/, "")}/browse/${encodeURIComponent(key)}`;
}

/**
 * Prints the result to the terminal.
 *
 * console.log rather than the structured logger, matching list_xray_projects:
 * this milestone is verified by reading it off the terminal during development.
 */
function printToConsole(
  items: FoundItem[],
  identity: AtlassianIdentity,
  projectKeys: string[],
  truncated: boolean
): void {
  const who = identity.displayName || identity.accountId;
  const lines: string[] = [
    "",
    "=".repeat(72),
    `find_my_xray_items -- Xray Tests naming ${who}`,
    "=".repeat(72),
    `Atlassian accountId: ${identity.accountId}`,
    `Matched on:          assignee` +
      (identity.displayName ? `, name "${identity.displayName}"` : "") +
      (identity.emailAddress ? `, email ${identity.emailAddress}` : " (no email available)"),
    `Projects searched:   ${projectKeys.join(", ")}`,
    `Matches:             ${items.length}` + (truncated ? " (capped; there are more)" : ""),
  ];

  if (items.length > 0) {
    const keyWidth = Math.max(3, ...items.map((i) => i.key.length));
    lines.push(
      "",
      `  ${"KEY".padEnd(keyWidth)}  ${"WHY".padEnd(10)}  SUMMARY`,
      `  ${"-".repeat(keyWidth)}  ${"-".repeat(10)}  ${"-".repeat(34)}`
    );
    for (const item of items) {
      lines.push(
        `  ${item.key.padEnd(keyWidth)}  ${item.reason.padEnd(10)}  ${item.summary || ""}`
      );
    }
  } else {
    lines.push("", "  (nothing matched)");
  }

  lines.push("=".repeat(72), "");
  console.log(lines.join("\n"));
}

/**
 * The same result, phrased for the model.
 *
 * The citation rules are as strict as the ones for documents and email, and for
 * the same reason: an issue summary is a record that somebody wrote that text,
 * not a statement of fact about the world.
 */
function formatForModel(
  items: FoundItem[],
  identity: AtlassianIdentity,
  projectKeys: string[],
  truncated: boolean
): string {
  const assigned = items.filter((i) => i.reason === "assigned");
  const both = items.filter((i) => i.reason === "both");
  const mentioned = items.filter((i) => i.reason === "mentioned");

  const render = (item: FoundItem) =>
    `- ${item.key}: ${item.summary || "(no summary)"}` +
    (item.url ? ` — ${item.url}` : "") +
    (item.reason !== "assigned" && item.assigneeName
      ? ` (assigned to ${item.assigneeName})`
      : "");

  const lines: string[] = [
    `Xray Tests naming this user (Atlassian account ${identity.accountId}` +
      (identity.displayName ? `, ${identity.displayName}` : "") +
      `), across the projects Knowva can see: ${projectKeys.join(", ")}.`,
    "",
    "ASSIGNED TO THEM:",
    assigned.length > 0 ? assigned.map(render).join("\n") : "- none",
    "",
    "ASSIGNED TO THEM AND MENTIONING THEM:",
    both.length > 0 ? both.map(render).join("\n") : "- none",
    "",
    "MENTIONING THEM (their name or email appears in the summary, description or comments; " +
      "these are NOT assigned to them):",
    mentioned.length > 0 ? mentioned.map(render).join("\n") : "- none",
  ];

  if (truncated) {
    lines.push(
      "",
      "There were more matches than were returned. Say the list is partial if the user asks " +
        "whether it is everything; do not present it as complete."
    );
  }

  lines.push(
    "",
    "How to report this:",
    "- Keep 'assigned to them' and 'mentions them' apart. Being named in a comment is not the " +
      "same as owning the work, and merging the two tells someone they are responsible for " +
      "something they are not.",
    "- Cite every item by its issue key, and link it in Markdown using the URL given. Never " +
      "describe a test without naming which issue it is.",
    "- A summary is a record that somebody wrote that text, not a fact about the world. Do not " +
      "restate one as though it were true -- say what the issue says.",
    "- You have seen no test steps, no descriptions and no comments, only what is listed above. " +
      "Do not describe what a test does beyond its summary, and do not say why someone was " +
      "mentioned -- you were not shown the text that matched.",
    "- This covers Xray Tests only. If they ask about Defects or other issue types, say Knowva " +
      "cannot search those yet rather than implying none exist.",
    "- The search ran on Knowva's own Xray credential, not on this user's Jira account, so it " +
      "covers what Knowva can see rather than everything they could see in Jira. Mention that " +
      "only if it bears on the answer."
  );

  return lines.join("\n");
}

// --- Error results -------------------------------------------------------

function atlassianNotConfiguredResult(): AgentToolResult {
  return {
    content:
      "Jira sign-in isn't set up on this deployment -- Knowva has no Atlassian OAuth app " +
      "configured, so it cannot work out which Jira account belongs to this user. Tell them " +
      "it's a setup problem on our side, not something they did.",
    isError: true,
  };
}

function xrayNotConfiguredResult(): AgentToolResult {
  return {
    content:
      "Xray isn't set up on this deployment -- Knowva has no Xray service credential " +
      "configured. Tell the user it's a setup problem on our side, not something they did, and " +
      "not a problem with their own access.",
    isError: true,
  };
}

function identityRefusalResult(reason: string | undefined): AgentToolResult {
  return {
    content:
      reason === "mismatch"
        ? "Knowva could not safely establish whose items to look for, so it did nothing. This " +
          "is a safety stop. Tell the user something is wrong with their sign-in and to sign " +
          "out and back in."
        : "Knowva could not confirm who it would be searching for, so it did nothing. This is a " +
          "safety stop, not a permissions problem. Tell the user to sign out and sign back in, " +
          "then try again.",
    isError: true,
  };
}

/**
 * The "sign in first" result: prose for the model, plus the signal that makes
 * app.ts attach a sign-in card (see AgentToolSignal in src/agent/loop.ts).
 *
 * NOTE WHAT THIS DOES NOT SAY, compared with the GitHub equivalent. There is no
 * "expired" or "could not be renewed" variant, because there is no token with a
 * lifetime -- only a resolved identity that is either cached or not. A user
 * whose identity was dropped by a restart is in exactly the same position as
 * one who never signed in, and telling them their "connection expired" would
 * imply Knowva had access it never had.
 */
function needsAuthResult(): AgentToolResult {
  return {
    // No URL in here on purpose -- the card carries it.
    content:
      "Knowva doesn't yet know which Jira account belongs to this user, so nothing was " +
      "searched. It is showing them a sign-in button, so do NOT paste a link or invent one. " +
      "Tell them briefly that you need to know which Jira account is theirs and that there's a " +
      "button just below your message, then stop. Say nothing about what you might have found " +
      "-- you searched nothing. Make clear this is an Atlassian sign-in, separate from their " +
      "Microsoft Teams sign-in and from GitHub, and that it only reads their Jira profile -- " +
      "Knowva does not keep access to their Jira.",
    signal: { kind: "needs-auth", provider: "atlassian" },
  };
}

function noProjectsResult(): AgentToolResult {
  return {
    content:
      "Knowva's Xray credential can see no projects at all, so there was nothing to search. " +
      "This is not a problem with the person asking or their access. Tell them plainly that " +
      "Knowva cannot currently see any Xray projects and that an administrator would need to " +
      "check the Xray setup. Do not name any issue -- you were given none.",
    isError: true,
  };
}

function noMatchesResult(identity: AtlassianIdentity, projectKeys: string[]): AgentToolResult {
  return {
    // Deliberately NOT isError. Finding nothing is a real, correct answer, and
    // flagging it as an error invites the model to apologise for a fault that
    // did not occur.
    content:
      `No Xray Tests are assigned to this user or mention them, across the projects Knowva can ` +
      `see (${projectKeys.join(", ")}).` +
      (identity.emailAddress
        ? ""
        : " Note their Atlassian profile does not expose an email address, so the search matched " +
          "on their display name only -- a mention written as an email address would not have " +
          "been found.") +
      " This is a genuine 'nothing found', not a failure: say so plainly, without apologising " +
      "for a problem and without suggesting anything is broken. Mention that it covers Xray " +
      "Tests only and not Defects. Do NOT answer from your own knowledge or invent an example.",
  };
}

function xrayFailureResult(
  reason: XrayFailure | undefined,
  detail: string | undefined,
  jql?: string
): AgentToolResult {
  if (reason === "not-configured") return xrayNotConfiguredResult();

  // The one failure with a real-world cause the user might care about: this
  // user's name matches so much that Jira's 100-issue ceiling is hit. Common
  // for short or common names, and the advice differs completely from the
  // other cases.
  if (reason === "jql-too-broad") {
    return {
      content:
        "The search matched more than 100 Jira issues, which Xray refuses to process. This " +
        "usually means this user's name is common enough to appear in a great many issues. " +
        "Tell them the search was too broad to run and that narrowing it would need a feature " +
        "Knowva does not have yet -- do not guess at what it might have found, and do not " +
        "retry the same search.",
      isError: true,
    };
  }

  const explanation =
    reason === "auth-failed"
      ? "Knowva's shared Xray credential was rejected. This is a setup problem on our side -- " +
        "NOT anything to do with this user's sign-in or their own Xray access."
      : reason === "query-error"
        ? "Xray rejected Knowva's search query. That is a bug on our side, not the user's " +
          "problem -- say so plainly rather than suggesting anything they could do differently."
        : "Knowva couldn't reach Xray. Tell the user it's a problem on our side, not theirs, " +
          "and to try again shortly.";

  return {
    content:
      explanation +
      " Do not guess at which tests exist or which name this user -- you retrieved nothing." +
      (detail ? ` (Technical detail, not for the user: ${detail})` : "") +
      (jql ? ` (Query that failed, not for the user: ${jql})` : ""),
    isError: true,
  };
}
