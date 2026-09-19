import { Client as GraphClient, GraphError } from "@microsoft/teams.graph";

export interface GraphMe {
  /**
   * The signed-in user's Entra object id. This is the identity the email
   * opt-out in src/user/email-preferences.ts is keyed on -- it is the same
   * value Teams puts on `activity.from.aadObjectId`, which is what lets
   * /disable-email and the pre-flight check in src/tools/email-access.ts agree
   * on who they are talking about.
   */
  id?: string;
  displayName: string;
  userPrincipalName?: string;
  mail?: string;
}

/**
 * Calls Microsoft Graph's /me endpoint.
 * `graph` must already carry the signed-in user's token
 * (e.g. `context.userGraph` from the Teams SSO connection).
 */
export async function getMe(graph: GraphClient): Promise<GraphMe> {
  const response = await graph.http.get<GraphMe>("/me");
  return response.data;
}

export { GraphError };
