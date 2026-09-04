import { Client as GraphClient, GraphError } from "@microsoft/teams.graph";

export interface GraphMe {
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
