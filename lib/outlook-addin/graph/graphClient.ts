import { Client } from "@microsoft/microsoft-graph-client";
import { acquireGraphToken } from "../auth/msal";

let graphClient: Client | null = null;

/**
 * Get or create the Microsoft Graph client
 */
export async function getGraphClient(): Promise<Client> {
  if (graphClient) {
    return graphClient;
  }

  graphClient = Client.init({
    authProvider: async (done) => {
      try {
        const token = await acquireGraphToken();
        done(null, token);
      } catch (error) {
        done(error as Error, null);
      }
    },
  });

  return graphClient;
}

/**
 * Reset the graph client (useful after sign out)
 */
export function resetGraphClient(): void {
  graphClient = null;
}
