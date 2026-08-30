const DEFAULT_GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

function normalizeUser(user) {
  const email = String(user.mail || user.userPrincipalName || "").trim();
  return {
    id: String(user.id || ""),
    name: String(user.displayName || email || "氏名未設定").trim(),
    email,
    department: String(user.department || "").trim(),
  };
}

function includesQuery(user, query) {
  const normalized = query.toLocaleLowerCase("ja-JP");
  return [user.name, user.email, user.department]
    .some((value) => value.toLocaleLowerCase("ja-JP").includes(normalized));
}

export function createGraphDirectory({
  acquireAccessToken,
  fetchImpl = fetch,
  graphBaseUrl = DEFAULT_GRAPH_BASE_URL,
  cacheTtlMs = 300_000,
  maxUsers = 5_000,
  now = Date.now,
} = {}) {
  let cachedUsers = [];
  let expiresAt = 0;
  let inFlight;

  async function loadUsers() {
    if (cachedUsers.length > 0 && now() < expiresAt) return cachedUsers;
    if (inFlight) return inFlight;

    inFlight = (async () => {
      const token = await acquireAccessToken();
      const users = [];
      const select = "id,displayName,mail,userPrincipalName,department,accountEnabled";
      let nextUrl = `${graphBaseUrl}/users?$select=${encodeURIComponent(select)}&$top=999`;

      while (nextUrl && users.length < maxUsers) {
        const response = await fetchImpl(nextUrl, {
          headers: { Authorization: `Bearer ${token}` },
        });
        if (!response.ok) {
          const requestId = response.headers?.get?.("request-id") || response.headers?.get?.("client-request-id");
          throw new Error(`Microsoft Graph user request failed (${response.status})${requestId ? ` request-id=${requestId}` : ""}`);
        }
        const payload = await response.json();
        for (const raw of payload.value || []) {
          if (raw.accountEnabled === false) continue;
          const user = normalizeUser(raw);
          if (user.id && user.email) users.push(user);
          if (users.length >= maxUsers) break;
        }
        nextUrl = typeof payload["@odata.nextLink"] === "string" ? payload["@odata.nextLink"] : "";
      }

      cachedUsers = users.sort((left, right) => left.name.localeCompare(right.name, "ja"));
      expiresAt = now() + cacheTtlMs;
      return cachedUsers;
    })().finally(() => {
      inFlight = undefined;
    });

    return inFlight;
  }

  return {
    async search(query, limit = 10) {
      const normalizedQuery = String(query || "").trim();
      if (!normalizedQuery) return [];
      const users = await loadUsers();
      return users.filter((user) => includesQuery(user, normalizedQuery)).slice(0, limit);
    },
    clearCache() {
      cachedUsers = [];
      expiresAt = 0;
    },
  };
}

export { normalizeUser };
