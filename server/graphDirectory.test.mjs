import assert from "node:assert/strict";
import test from "node:test";
import { createGraphDirectory } from "./graphDirectory.mjs";

test("氏名・メール・部門を途中一致で検索し、無効ユーザーを除外する", async () => {
  let requestCount = 0;
  const responses = [
    {
      value: [
        { id: "1", displayName: "山田 太郎", mail: "taro.yamada@example.com", department: "技術部", accountEnabled: true },
        { id: "2", displayName: "無効 利用者", mail: "disabled@example.com", department: "技術部", accountEnabled: false },
      ],
      "@odata.nextLink": "https://graph.example/users?page=2",
    },
    {
      value: [
        { id: "3", displayName: "鈴木 一郎", mail: null, userPrincipalName: "ichiro.suzuki@example.com", department: null, accountEnabled: true },
      ],
    },
  ];
  const directory = createGraphDirectory({
    acquireAccessToken: async () => "test-token",
    fetchImpl: async () => {
      const payload = responses[requestCount++];
      return { ok: true, json: async () => payload, headers: { get: () => null } };
    },
    graphBaseUrl: "https://graph.example",
  });

  assert.deepEqual((await directory.search("yamada")).map((user) => user.id), ["1"]);
  assert.deepEqual((await directory.search("技")).map((user) => user.id), ["1"]);
  assert.equal((await directory.search("suzuki"))[0].email, "ichiro.suzuki@example.com");
  assert.equal((await directory.search("suzuki"))[0].department, "");
  assert.equal(requestCount, 2, "2回目以降の検索はキャッシュを使用する");
});
