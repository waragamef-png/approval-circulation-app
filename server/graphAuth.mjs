import { readFile } from "node:fs/promises";
import { ConfidentialClientApplication } from "@azure/msal-node";

function required(name) {
  const value = process.env[name]?.trim();
  if (!value) throw new Error(`${name} is not configured`);
  return value;
}

export function isGraphConfigured() {
  const hasBase = Boolean(process.env.GRAPH_TENANT_ID?.trim() && process.env.GRAPH_CLIENT_ID?.trim());
  const hasCertificate = Boolean(process.env.GRAPH_CLIENT_CERTIFICATE_PATH?.trim() && process.env.GRAPH_CLIENT_CERTIFICATE_THUMBPRINT?.trim());
  const hasSecret = Boolean(process.env.GRAPH_CLIENT_SECRET?.trim());
  return hasBase && (hasCertificate || hasSecret);
}

export async function createAccessTokenProvider() {
  const tenantId = required("GRAPH_TENANT_ID");
  const clientId = required("GRAPH_CLIENT_ID");
  const certificatePath = process.env.GRAPH_CLIENT_CERTIFICATE_PATH?.trim();
  const certificateThumbprint = process.env.GRAPH_CLIENT_CERTIFICATE_THUMBPRINT?.trim();
  const clientSecret = process.env.GRAPH_CLIENT_SECRET?.trim();

  let credential;
  if (certificatePath && certificateThumbprint) {
    credential = {
      clientCertificate: {
        thumbprintSha256: certificateThumbprint,
        privateKey: await readFile(certificatePath, "utf8"),
      },
    };
  } else if (clientSecret) {
    credential = { clientSecret };
  } else {
    throw new Error("Graph certificate or client secret is not configured");
  }

  const client = new ConfidentialClientApplication({
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
      ...credential,
    },
  });

  return async () => {
    const response = await client.acquireTokenByClientCredential({
      scopes: ["https://graph.microsoft.com/.default"],
    });
    if (!response?.accessToken) throw new Error("Microsoft Graph access token was not returned");
    return response.accessToken;
  };
}
