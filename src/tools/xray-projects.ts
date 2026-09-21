import type { ILogger } from "@microsoft/teams.common";
import type { AgentTool, AgentToolResult } from "../agent/loop";
import {
  discoverXrayProjects,
  isXrayConfigured,
  type XrayFailure,
  type XrayProject,
} from "../xray/client";

/**
 * The `list_xray_projects` tool: which projects hold Xray test data.
 *
 * ===========================================================================
 * THERE IS NO IDENTITY CHECK IN THIS FILE, AND THAT IS DELIBERATE.
 *
 * Every other credential-spending tool in Knowva starts by resolving WHOSE
 * credentials this turn may spend -- src/tools/github-access.ts and
 * src/tools/email-access.ts both exist for that one job, and both refuse on any
 * doubt. This tool has no such gate, because there is nothing for it to decide:
 *
 *   - There is no per-user Xray credential to pick between. Xray Cloud
 *     authenticates with an app-level API key and offers no delegated or
 *     per-user access of any kind.
 *   - There is no per-user answer to produce. The result comes from one shared
 *     credential and is identical for every Teams user.
 *
 * So the confused-deputy question that makes identity resolution load-bearing
 * for GitHub -- "could we spend Alice's credential on Bob's question?" -- has
 * no analogue here. There is one credential and one answer.
 *
 * DO NOT "HARDEN" THIS BY ADDING AN ActingIdentityResolver. It would resolve an
 * identity, do nothing with it, and imply a per-user guarantee the underlying
 * API cannot provide. A check that cannot fail meaningfully is worse than no
 * check, because the next reader will believe it.
 * ===========================================================================
 *
 * ---------------------------------------------------------------------------
 * WHAT BOUNDS THE ANSWER: THE API KEY, AND NOTHING ELSE.
 *
 * An earlier version took an admin-set allowlist (XRAY_PROJECTS) and reported
 * only projects on it. That was removed deliberately, so this tool now reports
 * every project the service credential can see. Anybody who can talk to Knowva
 * can therefore learn which projects have Xray tests and roughly how many.
 *
 * The way to narrow that is to issue the API key for an Xray user with
 * narrower Jira permissions. There is no application-level filter any more, and
 * adding one back means reintroducing the allowlist.
 * ---------------------------------------------------------------------------
 *
 * DIAGNOSTIC BY DESIGN. This is a milestone checkpoint, not the Xray feature.
 * It reports which projects hold tests; it does not search Xray, read a test,
 * or return any test content, and its description to the model is deliberately
 * narrow so the model does not offer capabilities that do not exist yet.
 *
 * IT ALSO PRINTS TO THE CONSOLE. console.log, not log.info, and on purpose: the
 * point of this milestone is to be able to read the raw result straight off the
 * terminal while developing, without it being interleaved into the structured
 * bot logs. The same content goes back to the model.
 */

export function createXrayProjectsTool(log: ILogger): AgentTool {
  return {
    definition: {
      name: "list_xray_projects",
      description:
        "Diagnostic. Lists the Jira projects that contain Xray test data, with a test count for " +
        "each. The result comes from a single shared Xray service credential, so it is the same " +
        "for every user and is NOT personalised to whoever is asking. It does NOT search Xray " +
        "or return any test, test-run, or issue content -- only project keys and test counts. " +
        "Use it when the user asks which Xray projects exist or are available, or asks you to " +
        "check that the Xray connection is working.",
      inputSchema: {
        type: "object",
        properties: {},
        additionalProperties: false,
      },
    },

    async run(): Promise<AgentToolResult> {
      if (!isXrayConfigured()) return notConfiguredResult();

      const discovery = await discoverXrayProjects(log);

      if (!discovery.ok) return failureResult(discovery.reason, discovery.detail);

      const projects = discovery.projects || [];

      if (projects.length === 0) {
        printEmptyToConsole();
        return { content: emptyResultForModel() };
      }

      printToConsole(projects, discovery.totalTests, discovery.scannedAll !== false);

      return {
        content: formatForModel(projects, discovery.totalTests, discovery.scannedAll !== false),
      };
    },
  };
}

/**
 * Prints the result to the terminal.
 *
 * Explicitly console.log rather than the structured logger, because this
 * milestone is verified by reading it. Deliberately plain: aligned columns, no
 * colour codes, nothing that depends on a particular terminal.
 */
function printToConsole(
  projects: XrayProject[],
  totalTests: number | undefined,
  scannedAll: boolean
): void {
  const keyWidth = Math.max(3, ...projects.map((p) => label(p).length));

  const lines: string[] = [
    "",
    "=".repeat(72),
    "list_xray_projects -- everything the shared Xray credential can see",
    "=".repeat(72),
    `Projects with Xray test data: ${projects.length}` +
      (typeof totalTests === "number" ? `   (${totalTests} test(s) in total)` : ""),
  ];

  if (!scannedAll) {
    lines.push(
      "NOTE: the discovery scan hit its page cap, so a project whose tests all",
      "      sat beyond it would have been missed. Treat this as incomplete."
    );
  }

  lines.push(
    "",
    `  ${"PROJECT".padEnd(keyWidth)}  ${"TESTS".padEnd(8)}  NAME`,
    `  ${"-".repeat(keyWidth)}  ${"-".repeat(8)}  ${"-".repeat(28)}`
  );

  for (const project of projects) {
    lines.push(
      `  ${label(project).padEnd(keyWidth)}  ` +
        `${countCell(project).padEnd(8)}  ${project.name || ""}`
    );
  }

  lines.push("=".repeat(72), "");

  console.log(lines.join("\n"));
}

function printEmptyToConsole(): void {
  console.log(
    [
      "",
      "=".repeat(72),
      "list_xray_projects -- everything the shared Xray credential can see",
      "=".repeat(72),
      "Projects with Xray test data: 0",
      "",
      "The credential authenticated, but Xray returned no tests at all. Either",
      "this tenant has no Xray tests yet, or the Jira user behind the API key",
      "cannot browse any project that has them.",
      "=".repeat(72),
      "",
    ].join("\n")
  );
}

/** Project key when Xray resolved one, otherwise the numeric id so the row is still identifiable. */
function label(project: XrayProject): string {
  return project.key || `id ${project.projectId}`;
}

function countCell(project: XrayProject): string {
  return typeof project.testCount === "number" ? String(project.testCount) : "?";
}

/**
 * The same result, phrased for the model.
 *
 * The "not personalised" caveat is load-bearing and is repeated here rather
 * than left to the system prompt alone. Without it the model reliably describes
 * this as "the projects you have access to", which is a per-user claim the tool
 * has not made and cannot make.
 */
function formatForModel(
  projects: XrayProject[],
  totalTests: number | undefined,
  scannedAll: boolean
): string {
  const lines: string[] = [
    `Projects containing Xray test data: ${projects.length}` +
      (typeof totalTests === "number" ? ` (${totalTests} Xray tests in total).` : ".") +
      " This came from a single shared Xray service credential, so it is the same for every " +
      "user and is NOT based on who is asking. Never call it 'the projects you have access to', " +
      "never say it was checked against their account, and never imply another user would see " +
      "a different list.",
    "",
    projects
      .map((p) => {
        const name = p.name ? ` (${p.name})` : "";
        const count =
          typeof p.testCount === "number"
            ? `${p.testCount} Xray test(s)`
            : "test count unavailable";
        return `- ${label(p)}${name} -- ${count}`;
      })
      .join("\n"),
  ];

  if (!scannedAll) {
    lines.push(
      "",
      "IMPORTANT: this scan hit its limit, so the list may be incomplete -- a project could " +
        "exist that is not shown here. Say so if the user asks whether this is everything. Do " +
        "not present it as a complete list of all projects."
    );
  }

  lines.push(
    "",
    "How to report this: list the projects plainly. Do not invent test names, counts you were " +
      "not given, or anything about what the tests contain -- you have seen no test content. If " +
      "the user asks you to search these projects or read a test, say you cannot do that yet; " +
      "listing which projects have tests is all this tool does."
  );

  return lines.join("\n");
}

function emptyResultForModel(): string {
  return (
    "Xray authenticated successfully but reported no tests at all, so there are no projects to " +
    "list. This is an empty result, NOT an error and NOT a permissions problem with the person " +
    "asking. It means either this Xray tenant has no tests yet, or the account behind Knowva's " +
    "Xray credential cannot see any project that has them. Tell them plainly and suggest an " +
    "administrator check the Xray setup. Do not name any project -- you were given none."
  );
}

function failureResult(
  reason: XrayFailure | undefined,
  detail: string | undefined
): AgentToolResult {
  if (reason === "not-configured") return notConfiguredResult();

  const explanation =
    reason === "auth-failed"
      ? "Knowva's shared Xray credential was rejected. This is a setup problem on our side -- " +
        "NOT anything to do with the person asking, their sign-in, or their own Xray access. " +
        "Tell them it's broken on our end and that an administrator needs to look at it."
      : reason === "unknown-project"
        ? "Xray could not resolve a project it had just reported. That is a problem on our side; " +
          "tell the user so and suggest trying again shortly."
        : reason === "query-error"
          ? "Xray rejected Knowva's query. That is a bug on our side, not the user's problem -- " +
            "say so plainly rather than suggesting anything they could do differently."
          : "Knowva couldn't reach Xray. Tell the user it's a problem on our side, not theirs, " +
            "and to try again shortly.";

  return {
    content:
      explanation +
      " Do not guess at which projects exist or how many tests they have -- you retrieved " +
      "nothing." +
      (detail ? ` (Technical detail, not for the user: ${detail})` : ""),
    isError: true,
  };
}

function notConfiguredResult(): AgentToolResult {
  return {
    content:
      "Xray isn't set up on this deployment -- Knowva has no Xray service credential configured. " +
      "Tell the user it's a setup problem on our side, not something they did, and not a " +
      "problem with their own access.",
    isError: true,
  };
}
