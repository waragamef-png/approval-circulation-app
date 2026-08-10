import type { AppData, MockFile, RouteMember } from "./types";

const now = "2026-08-10T09:00:00+09:00";

export const mockFiles: MockFile[] = [
  { id: "sp-001", provider: "sharepoint", name: "2026年度_製品仕様書.pdf", type: "PDF", location: "技術部 / 製品開発", updatedAt: "2026-08-09T16:40:00+09:00", updatedBy: "佐藤 花子" },
  { id: "sp-002", provider: "sharepoint", name: "品質検査チェックリスト.xlsx", type: "Excel", location: "品質保証部 / 品質文書", updatedAt: "2026-08-08T14:10:00+09:00", updatedBy: "田中 一郎" },
  { id: "od-001", provider: "onedrive", name: "設備投資申請書.xlsx", type: "Excel", location: "マイファイル / 申請書", updatedAt: "2026-08-10T08:20:00+09:00", updatedBy: "藤原 美咲" },
  { id: "od-002", provider: "onedrive", name: "新製品企画書.pptx", type: "PowerPoint", location: "マイファイル / 企画", updatedAt: "2026-08-07T17:35:00+09:00", updatedBy: "藤原 美咲" },
  { id: "sf-001", provider: "shared-folder", name: "月次実績報告.xlsx", type: "Excel", location: "\\\\fileserver\\管理部\\月次", updatedAt: "2026-08-10T07:50:00+09:00", updatedBy: "佐藤 花子" },
];

const approvers = [
  { id: "a-yamada", name: "山田 太郎", email: "taro.yamada@example.com", department: "技術部", entraUserId: "", active: true, createdAt: now, updatedAt: now },
  { id: "a-tanaka", name: "田中 一郎", email: "ichiro.tanaka@example.com", department: "品質保証部", entraUserId: "", active: true, createdAt: now, updatedAt: now },
  { id: "a-suzuki", name: "鈴木 次郎", email: "jiro.suzuki@example.com", department: "開発部", entraUserId: "", active: true, createdAt: now, updatedAt: now },
  { id: "a-sato", name: "佐藤 花子", email: "hanako.sato@example.com", department: "管理部", entraUserId: "", active: true, createdAt: now, updatedAt: now },
  { id: "a-fujiwara", name: "藤原 美咲", email: "misaki.fujiwara@example.com", department: "経営企画部", entraUserId: "", active: true, createdAt: now, updatedAt: now },
  { id: "a-takahashi", name: "高橋 健太", email: "kenta.takahashi@example.com", department: "情報システム部", entraUserId: "", active: false, createdAt: now, updatedAt: now },
];

const member = (approverId: string, sequence: number, requiresStamp: boolean, status: RouteMember["status"], completedAt?: string): RouteMember => {
  const person = approvers.find((item) => item.id === approverId)!;
  return { id: `member-${approverId}-${sequence}`, approverId, name: person.name, email: person.email, department: person.department, sequence, requiresStamp, status, completedAt };
};

export const seedData: AppData = {
  approvers,
  currentUserId: "a-fujiwara",
  templates: [
    {
      id: "template-dev",
      name: "開発仕様書承認ルート",
      description: "技術確認から品質保証、部門承認までの標準ルート",
      createdBy: "藤原 美咲",
      createdAt: now,
      updatedAt: now,
      members: [
        { ...member("a-yamada", 1, true, "pending"), status: undefined as never },
        { ...member("a-tanaka", 2, false, "pending"), status: undefined as never },
        { ...member("a-suzuki", 3, false, "pending"), status: undefined as never },
      ].map(({ status: _status, completedAt: _completedAt, ...item }) => item),
    },
    {
      id: "template-quality",
      name: "品質保証ルート",
      description: "品質保証部を中心とした確認ルート",
      createdBy: "藤原 美咲",
      createdAt: now,
      updatedAt: now,
      members: [
        { ...member("a-tanaka", 1, true, "pending"), status: undefined as never },
        { ...member("a-sato", 2, false, "pending"), status: undefined as never },
      ].map(({ status: _status, completedAt: _completedAt, ...item }) => item),
    },
  ],
  cases: [
    {
      id: "CASE-2026-0810-001",
      accessKey: "a8f3d1c9e7424b6aa0518d732c901e45",
      provider: "sharepoint",
      fileId: "sp-001",
      fileName: "2026年度_製品仕様書.pdf",
      fileUrl: "#demo-document",
      documents: [
        { id: "doc-sp-001", fileId: "sp-001", name: "2026年度_製品仕様書.pdf", type: "PDF", location: "技術部 / 製品開発", fileUrl: "#demo-document", requiresStamp: true },
        { id: "doc-sp-002", fileId: "sp-002", name: "品質検査チェックリスト.xlsx", type: "Excel", location: "品質保証部 / 品質文書", fileUrl: "#demo-document", requiresStamp: false },
      ],
      initiatorId: "a-sato",
      initiatorName: "佐藤 花子",
      startedAt: "2026-08-10T09:00:00+09:00",
      updatedAt: "2026-08-10T11:10:00+09:00",
      state: "circulating",
      currentMemberId: "member-a-fujiwara-3",
      members: [
        member("a-yamada", 1, true, "stamped", "2026-08-10T10:15:00+09:00"),
        member("a-tanaka", 2, false, "approved", "2026-08-10T11:10:00+09:00"),
        member("a-fujiwara", 3, true, "current"),
        member("a-suzuki", 4, false, "pending"),
      ],
      history: [
        { id: "h-1", caseId: "CASE-2026-0810-001", actionUserId: "a-sato", actionUserName: "佐藤 花子", action: "回覧開始", previousState: "下書き", newState: "回覧中（山田 太郎）", createdAt: "2026-08-10T09:00:00+09:00" },
        { id: "h-2", caseId: "CASE-2026-0810-001", actionUserId: "a-yamada", actionUserName: "山田 太郎", action: "捺印済み・回付", previousState: "山田 太郎が処理中", newState: "田中 一郎が処理中", createdAt: "2026-08-10T10:15:00+09:00" },
        { id: "h-3", caseId: "CASE-2026-0810-001", actionUserId: "a-tanaka", actionUserName: "田中 一郎", action: "確認OK・回付", previousState: "田中 一郎が処理中", newState: "藤原 美咲が処理中", createdAt: "2026-08-10T11:10:00+09:00" },
      ],
    },
    {
      id: "CASE-2026-0808-002",
      accessKey: "c4b8720fe31549d1b608e27a4f913c60",
      provider: "onedrive",
      fileId: "od-001",
      fileName: "設備投資申請書.xlsx",
      fileUrl: "#demo-document",
      documents: [
        { id: "doc-od-001", fileId: "od-001", name: "設備投資申請書.xlsx", type: "Excel", location: "マイファイル / 申請書", fileUrl: "#demo-document", requiresStamp: true },
      ],
      initiatorId: "a-fujiwara",
      initiatorName: "藤原 美咲",
      startedAt: "2026-08-08T13:00:00+09:00",
      updatedAt: "2026-08-09T15:30:00+09:00",
      state: "completed",
      members: [
        member("a-yamada", 1, false, "approved", "2026-08-08T15:10:00+09:00"),
        member("a-sato", 2, true, "stamped", "2026-08-09T15:30:00+09:00"),
      ],
      history: [
        { id: "h-4", caseId: "CASE-2026-0808-002", actionUserId: "a-fujiwara", actionUserName: "藤原 美咲", action: "回覧開始", previousState: "下書き", newState: "回覧中（山田 太郎）", createdAt: "2026-08-08T13:00:00+09:00" },
        { id: "h-5", caseId: "CASE-2026-0808-002", actionUserId: "a-yamada", actionUserName: "山田 太郎", action: "確認OK・回付", previousState: "山田 太郎が処理中", newState: "佐藤 花子が処理中", createdAt: "2026-08-08T15:10:00+09:00" },
        { id: "h-6", caseId: "CASE-2026-0808-002", actionUserId: "a-sato", actionUserName: "佐藤 花子", action: "捺印済み・回付", previousState: "佐藤 花子が処理中", newState: "回覧完了", createdAt: "2026-08-09T15:30:00+09:00" },
      ],
    },
  ],
};
