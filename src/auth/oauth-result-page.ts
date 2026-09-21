/**
 * The little "you can close this tab now" page an OAuth callback renders.
 *
 * Extracted when Atlassian became the second provider, rather than copied a
 * second time. src/auth/github.ts still carries its own private copy of this:
 * it predates this file and the milestone that added Atlassian was scoped not
 * to touch the working GitHub flow. Folding it in is a one-line change and
 * should happen the next time that file is opened for another reason.
 *
 * Every interpolated value is escaped at the call site AND the heading is
 * escaped here. The only dynamic values that ever reach this are an account
 * name and a site name, but a page that renders in a user's browser is not the
 * place to rely on a value happening to be safe today.
 */
export function oauthResultPage(heading: string, body: string): string {
  return (
    "<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\">" +
    "<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">" +
    `<title>${escapeHtml(heading)} - Knowva</title>` +
    "<style>body{font-family:system-ui,-apple-system,Segoe UI,sans-serif;margin:0;" +
    "display:grid;place-items:center;min-height:100vh;background:#faf9f8;color:#201f1e}" +
    "main{max-width:32rem;padding:2rem;text-align:center}" +
    "h1{font-size:1.35rem;margin:0 0 .75rem}p{margin:0;line-height:1.5;color:#484644}</style>" +
    `</head><body><main><h1>${escapeHtml(heading)}</h1><p>${body}</p></main></body></html>`
  );
}

export function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
