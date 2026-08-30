import { useEffect, useId, useMemo, useState } from "react";
import { mockFiles } from "./mockData";
import { loadData, resetData, saveData, STORAGE_KEY } from "./storage";
import { directorySearchEnabled, searchDirectoryUsers, type DirectoryUser } from "./services/directory";
import { listSharePointFiles, sharePointFileListingEnabled } from "./services/sharepoint";
import type {
  AppData,
  Approver,
  CaseDocument,
  CirculationCase,
  HistoryEntry,
  MockFile,
  ProviderType,
  RouteMember,
  RouteTemplate,
} from "./types";

type Page = "home" | "approvers" | "new" | "case";
type CaseTab = "waiting" | "all" | "returned" | "completed";
type EmailDraft = { to: string; cc: string[]; recipientName: string; subject: string; body: string; purpose: string };
type PendingStart = { item: CirculationCase; draft: EmailDraft };
type PendingCaseTransition = { changed: CirculationCase; draft: EmailDraft; successMessage: string };

function readLocation(): { page: Page; caseId?: string } {
  const basePath = import.meta.env.BASE_URL.replace(/\/$/, "");
  const pathname = basePath && window.location.pathname.startsWith(basePath)
    ? window.location.pathname.slice(basePath.length)
    : window.location.pathname;
  const path = (window.location.hash.startsWith("#/") ? window.location.hash.slice(1) : pathname).replace(/\/$/, "") || "/";
  if (path.startsWith("/c/")) return { page: "case", caseId: decodeURIComponent(path.slice("/c/".length)) };
  if (path.startsWith("/cases/")) return { page: "case", caseId: decodeURIComponent(path.slice("/cases/".length)) };
  if (path === "/cases") return { page: "home" };
  if (path === "/circulations/new") return { page: "new" };
  if (path === "/approvers") return { page: "approvers" };
  return { page: "home" };
}

function pathFor(page: Page, caseId?: string) {
  const path = page === "case" && caseId
    ? `/c/${encodeURIComponent(caseId)}`
    : ({ home: "/", approvers: "/approvers", new: "/circulations/new", case: "/" } as Record<Page, string>)[page];
  return import.meta.env.PROD ? `#${path}` : path;
}

const providerLabels: Record<ProviderType, string> = {
  sharepoint: "SharePoint",
  onedrive: "OneDrive",
  "shared-folder": "共有フォルダ",
};

const uid = (prefix: string) => `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
const nowIso = () => new Date().toISOString();
const formatDate = (value: string) => new Intl.DateTimeFormat("ja-JP", {
  year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit",
}).format(new Date(value));

const stateLabel = (state: CirculationCase["state"]) => ({ circulating: "回覧中", returned: "差し戻し", completed: "回覧完了" }[state]);
const caseTitle = (item: CirculationCase) => item.title?.trim() || (item.documents.length > 1 ? `${item.documents[0].name} ほか${item.documents.length - 1}件` : item.documents[0]?.name ?? item.fileName);
const fileLabel = (type: string) => type === "PDF" ? "PDF" : type === "Excel" ? "XLS" : type === "PowerPoint" ? "PPT" : "DOC";
const fileClass = (type: string) => type === "PDF" ? "pdf" : "excel";

function uniqueEmailAddresses(addresses: string[], excludedAddress = "") {
  const excluded = excludedAddress.trim().toLowerCase();
  const seen = new Set<string>();
  return addresses.map((address) => address.trim()).filter((address) => {
    const normalized = address.toLowerCase();
    if (!normalized || normalized === excluded || seen.has(normalized)) return false;
    seen.add(normalized);
    return true;
  });
}

function createPastApproverCc(item: CirculationCase, throughMemberId: string | undefined, to: string) {
  const throughIndex = throughMemberId ? item.members.findIndex((member) => member.id === throughMemberId) : -1;
  const participantHistory = item.history.filter((entry) => !["回覧開始", "再回覧", "回付ルート変更"].includes(entry.action));
  const historyMemberIds = new Set(participantHistory.flatMap((entry) => entry.actionMemberId ? [entry.actionMemberId] : []));
  const historyUserIds = new Set(participantHistory.map((entry) => entry.actionUserId));
  const addresses = item.members.filter((member, index) => !member.isInitiator && (
    (throughIndex >= 0 && index <= throughIndex) || historyMemberIds.has(member.id) || historyUserIds.has(member.approverId)
  )).map((member) => member.email);
  return uniqueEmailAddresses(addresses, to);
}

type RouteTimelineSegment = {
  id: string;
  members: RouteMember[];
  returnEntry?: HistoryEntry;
  current: boolean;
  startSequence: number;
};

function buildRouteTimelineSegments(item: CirculationCase): RouteTimelineSegment[] {
  const returnEntries = item.history.filter((entry) => entry.action.includes("NG") && entry.routeSnapshot?.length);
  if (returnEntries.length === 0) return [{ id: "current-route", members: item.members, current: true, startSequence: 1 }];

  const segments: RouteTimelineSegment[] = [];
  let nextStartMemberId: string | undefined;
  let nextStartUserId: string | undefined;
  let startSequence = 1;
  returnEntries.forEach((entry, index) => {
    const snapshot = entry.routeSnapshot ?? [];
    const snapshotStart = index === 0 ? 0 : Math.max(snapshot.findIndex((member) => nextStartMemberId ? member.id === nextStartMemberId : member.approverId === nextStartUserId), 0);
    const returnerIndex = snapshot.findIndex((member) => entry.actionMemberId ? member.id === entry.actionMemberId : member.approverId === entry.actionUserId);
    const endIndex = returnerIndex >= snapshotStart ? returnerIndex + 1 : snapshot.length;
    const members = snapshot.slice(snapshotStart, endIndex);
    segments.push({ id: entry.id, members, returnEntry: entry, current: false, startSequence });
    startSequence += members.length;
    nextStartMemberId = entry.returnToMemberId;
    nextStartUserId = entry.returnToUserId;
  });

  const currentStart = nextStartMemberId || nextStartUserId ? item.members.findIndex((member) => nextStartMemberId ? member.id === nextStartMemberId : member.approverId === nextStartUserId) : 0;
  const currentMembers = item.members.slice(Math.max(currentStart, 0));
  segments.push({ id: "current-route", members: currentMembers, current: true, startSequence });
  return segments;
}

function createStartEmailDraft(item: CirculationCase): EmailDraft | undefined {
  const recipient = item.members.find((member) => member.id === item.currentMemberId);
  if (!recipient) return undefined;
  const title = caseTitle(item);
  const stampCount = item.documents.filter((document) => document.requiresStamp).length;
  const needsStamp = recipient.requiresStamp && stampCount > 0;
  const caseUrl = new URL(pathFor("case", item.accessKey), `${window.location.origin}${window.location.pathname}`).toString();
  return {
    to: recipient.email,
    cc: [],
    recipientName: recipient.name,
    purpose: "最初の確認者への回付通知",
    subject: `【承認依頼】${title}`,
    body: `${recipient.name}さん\n\n${item.initiatorName}さんから承認回覧が開始され、あなたの承認待ちになりました。\n\n案件名：${title}\n文書数：${item.documents.length}件\n案件ID：${item.id}\n現在の確認者：${recipient.name}\n対応内容：${needsStamp ? `捺印対象${stampCount}件への捺印と全体確認` : "全文書の内容確認"}\n\n案件URL：\n${caseUrl}\n\nよろしくお願いいたします。`,
  };
}

function Modal({ title, children, onClose, wide = false }: { title: string; children: React.ReactNode; onClose: () => void; wide?: boolean }) {
  return (
    <div className="modal-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && onClose()}>
      <section className={`modal ${wide ? "modal-wide" : ""}`} role="dialog" aria-modal="true" aria-label={title}>
        <header className="modal-header"><h2>{title}</h2><button className="modal-close-button" onClick={onClose}>閉じる</button></header>
        {children}
      </section>
    </div>
  );
}

function HelpButton({ title, children, label = "使い方" }: { title: string; children: React.ReactNode; label?: string }) {
  const [open, setOpen] = useState(false);
  return <><button className="help-button" onClick={() => setOpen(true)} aria-label={`${title}を表示`}>{label}</button>{open && <Modal title={title} onClose={() => setOpen(false)}><div className="modal-body help-content">{children}</div><div className="modal-footer"><button className="primary-button" onClick={() => setOpen(false)}>閉じる</button></div></Modal>}</>;
}

function EmailComposerModal({ draft, onClose, onSend }: { draft: EmailDraft; onClose: () => void; onSend: (message: EmailDraft) => void }) {
  const [to, setTo] = useState(draft.to);
  const [cc, setCc] = useState(draft.cc.join("; "));
  const [subject, setSubject] = useState(draft.subject);
  const [body, setBody] = useState(draft.body);
  const parsedCc = uniqueEmailAddresses(cc.split(/[;,\n]/), to);
  const valid = to.includes("@") && parsedCc.every((address) => address.includes("@")) && subject.trim() && body.trim();
  return <Modal title="メールを作成" onClose={onClose} wide><div className="modal-body email-composer"><div className="demo-banner">確認用のため、実際には送信されません。「送信を確認」を押すと回付内容を反映します。</div><div className="mail-purpose"><span>送信目的</span><strong>{draft.purpose}</strong></div><div className="mail-fields"><label><span>差出人</span><div className="sender-field">社内承認回覧 &lt;approval-circulation@example.com&gt;<em>本運用時に設定</em></div></label><label><span>宛先</span><input type="email" value={to} onChange={(event) => setTo(event.target.value)} /></label><label><span>CC</span><input value={cc} onChange={(event) => setCc(event.target.value)} placeholder="過去の承認者を自動入力" /></label><label><span>件名</span><input value={subject} onChange={(event) => setSubject(event.target.value)} /></label><label><span>本文</span><textarea value={body} onChange={(event) => setBody(event.target.value)} rows={14} /></label></div></div><div className="modal-footer"><button className="quiet-button" onClick={onClose}>キャンセル</button><button className="primary-button send-button" disabled={!valid} onClick={() => onSend({ ...draft, to: to.trim(), cc: parsedCc, subject: subject.trim(), body: body.trim() })}>送信を確認</button></div></Modal>;
}

function RouteTypeLabel({ requiresStamp }: { requiresStamp: boolean }) {
  const label = requiresStamp ? "捺印あり" : "確認のみ";
  return <span className={`route-type-label ${requiresStamp ? "stamp" : "review"}`}>{label}</span>;
}

function StatusBadge({ state }: { state: CirculationCase["state"] }) {
  return <span className={`status-badge status-${state}`}>{stateLabel(state)}</span>;
}

function EmptyState({ title, text }: { title: string; text: string }) {
  return <div className="empty-state"><h3>{title}</h3><p>{text}</p></div>;
}

function ApproverAutocomplete({ approvers, onSelect, label = "承認者を検索", actionLabel = "追加", autoFocus = false }: {
  approvers: Approver[]; onSelect: (person: Approver) => void; label?: string; actionLabel?: string; autoFocus?: boolean;
}) {
  const inputId = useId();
  const [query, setQuery] = useState("");
  const [debounced, setDebounced] = useState("");
  const [directoryUsers, setDirectoryUsers] = useState<DirectoryUser[]>([]);
  const [directoryLoading, setDirectoryLoading] = useState(false);
  const [directoryError, setDirectoryError] = useState("");
  useEffect(() => {
    const timer = window.setTimeout(() => setDebounced(query.trim().toLowerCase()), 300);
    return () => window.clearTimeout(timer);
  }, [query]);
  useEffect(() => {
    if (!directorySearchEnabled || debounced.length < 1) {
      setDirectoryUsers([]);
      setDirectoryError("");
      setDirectoryLoading(false);
      return;
    }
    const controller = new AbortController();
    setDirectoryLoading(true);
    setDirectoryError("");
    searchDirectoryUsers(debounced, controller.signal)
      .then(setDirectoryUsers)
      .catch((error) => {
        if (error instanceof DOMException && error.name === "AbortError") return;
        setDirectoryUsers([]);
        setDirectoryError(error instanceof Error ? error.message : "社内名簿を検索できませんでした。");
      })
      .finally(() => {
        if (!controller.signal.aborted) setDirectoryLoading(false);
      });
    return () => controller.abort();
  }, [debounced]);
  const localResults = debounced.length < 1 ? [] : approvers.filter((person) =>
    person.active &&
    [person.name, person.email, person.department].some((value) => value.toLowerCase().includes(debounced)),
  );
  const localEmails = new Set(approvers.map((person) => person.email.toLowerCase()));
  const localEntraUserIds = new Set(approvers.flatMap((person) => person.entraUserId ? [person.entraUserId] : []));
  const remoteResults: Approver[] = directoryUsers
    .filter((person) => !localEntraUserIds.has(person.id) && !localEmails.has(person.email.toLowerCase()))
    .map((person) => ({
      id: `entra:${person.id}`,
      entraUserId: person.id,
      name: person.name,
      email: person.email,
      department: person.department,
      active: true,
      createdAt: "",
      updatedAt: "",
    }));
  const results = [...localResults, ...remoteResults].slice(0, 8);

  return (
    <div className="autocomplete">
      <label className="field-label" htmlFor={inputId}>{label}</label>
      <div className="search-input"><input id={inputId} value={query} onChange={(event) => setQuery(event.target.value)} placeholder="氏名・メール・部門を1文字以上入力" autoComplete="off" autoFocus={autoFocus} /></div>
      {debounced && <div className="suggestions">
        {results.length > 0 ? results.map((person) => (
          <button key={person.id} className="suggestion" onClick={() => { onSelect(person); setQuery(""); setDebounced(""); }}>
            <span><strong>{person.name}</strong><small>{person.department || "部門未設定"} ・ {person.email}</small></span>
            {person.id.startsWith("entra:") && <em className="directory-source">社内名簿</em>}
            <span className="suggestion-action">{actionLabel}</span>
          </button>
        )) : !directoryLoading && <div className="suggestion-empty">該当する有効な承認者はいません</div>}
        {directoryLoading && <div className="suggestion-note">社内名簿を検索中…</div>}
        {directoryError && <div className="suggestion-error">{directoryError}<small>承認者マスタの候補は引き続き利用できます。</small></div>}
      </div>}
    </div>
  );
}

function InitiatorPickerModal({ approvers, onSelect, onClose }: { approvers: Approver[]; onSelect: (person: Approver) => void; onClose: () => void }) {
  return <Modal title="開始者を変更" onClose={onClose}><div className="modal-body initiator-picker"><ApproverAutocomplete approvers={approvers} onSelect={onSelect} label="開始者を検索" actionLabel="選択" autoFocus /></div></Modal>;
}

export default function App() {
  const [data, setData] = useState<AppData>(() => loadData());
  const initialLocation = useMemo(() => readLocation(), []);
  const [page, setPage] = useState<Page>(initialLocation.page);
  const [selectedCaseId, setSelectedCaseId] = useState<string | undefined>(initialLocation.caseId);
  const [toast, setToast] = useState<string>();
  const [initiatorPickerOpen, setInitiatorPickerOpen] = useState(false);
  const [pendingStart, setPendingStart] = useState<PendingStart>();

  useEffect(() => saveData(data), [data]);
  useEffect(() => {
    const refreshFromAnotherWindow = (event: StorageEvent) => {
      if (event.key === STORAGE_KEY) setData(loadData());
    };
    window.addEventListener("storage", refreshFromAnotherWindow);
    return () => window.removeEventListener("storage", refreshFromAnotherWindow);
  }, []);
  useEffect(() => {
    const handlePopState = () => {
      const next = readLocation();
      setPage(next.page);
      setSelectedCaseId(next.caseId);
    };
    window.addEventListener("popstate", handlePopState);
    return () => window.removeEventListener("popstate", handlePopState);
  }, []);
  useEffect(() => {
    if (!toast) return;
    const timer = window.setTimeout(() => setToast(undefined), 3200);
    return () => window.clearTimeout(timer);
  }, [toast]);
  const currentUser = data.approvers.find((person) => person.id === data.currentUserId) ?? data.approvers[0];

  const navigate = (next: Page, caseId?: string) => {
    const addressKey = next === "case" ? (data.cases.find((item) => item.id === caseId)?.accessKey ?? caseId) : caseId;
    window.history.pushState({}, "", pathFor(next, addressKey));
    setPage(next);
    setSelectedCaseId(next === "case" ? addressKey : undefined);
    window.scrollTo({ top: 0, behavior: "smooth" });
  };
  const notify = (message: string) => setToast(message);
  const selectCurrentUser = (person: Approver) => {
    setData((current) => {
      const existing = current.approvers.find((candidate) =>
        candidate.id === person.id ||
        Boolean(person.entraUserId && candidate.entraUserId === person.entraUserId) ||
        candidate.email.toLowerCase() === person.email.toLowerCase(),
      );
      if (existing) return { ...current, currentUserId: existing.id };
      const timestamp = nowIso();
      const added = { ...person, createdAt: person.createdAt || timestamp, updatedAt: person.updatedAt || timestamp };
      return { ...current, approvers: [...current.approvers, added], currentUserId: added.id };
    });
    setInitiatorPickerOpen(false);
    notify(`${person.name}さんを開始者に設定しました`);
  };
  const updateApprovers = (approvers: Approver[]) => setData((current) => ({ ...current, approvers }));
  const updateTemplates = (templates: RouteTemplate[]) => setData((current) => ({ ...current, templates }));
  const updateCase = (changed: CirculationCase) => setData((current) => ({ ...current, cases: current.cases.map((item) => item.id === changed.id ? changed : item) }));

  const selectedCase = data.cases.find((item) => item.accessKey === selectedCaseId || item.id === selectedCaseId);
  useEffect(() => {
    if (page !== "case" || !selectedCase || selectedCaseId === selectedCase.accessKey) return;
    window.history.replaceState({}, "", pathFor("case", selectedCase.accessKey));
    setSelectedCaseId(selectedCase.accessKey);
  }, [page, selectedCase, selectedCaseId]);

  return (
    <div className="app-shell">
      <div className="main-column">
        <header className="topbar">
          <div className="topbar-start"><button className="app-title-button" onClick={() => navigate("home")}>社内承認回覧</button><span className="demo-pill">画面確認用</span></div>
          <div className="topbar-controls">{page === "new" && <div className="user-switcher"><span className="user-caption">開始者</span><strong className="current-user-name">{currentUser.name}</strong><button className="user-change-button" onClick={() => setInitiatorPickerOpen(true)}>変更</button></div>}</div>
        </header>

        <main>
          {page === "home" && <CasesPage cases={data.cases} onOpen={(id) => navigate("case", id)} newWindowUrl={pathFor("new")} approversWindowUrl={pathFor("approvers")} onReset={() => { setData(resetData()); notify("デモデータを初期状態に戻しました"); }} />}
          {page === "approvers" && <ApproversPage approvers={data.approvers} onChange={updateApprovers} notify={notify} />}
          {page === "new" && <NewCirculationPage data={data} currentUser={currentUser} setTemplates={updateTemplates} onStart={(created) => { const draft = createStartEmailDraft(created); if (!draft) { notify("最初の回付者のメール下書きを作成できませんでした"); return; } setPendingStart({ item: created, draft }); }} notify={notify} />}
          {page === "case" && selectedCase && <CaseDetail item={selectedCase} approvers={data.approvers} onBack={() => navigate("home")} onChange={updateCase} notify={notify} />}
          {page === "case" && !selectedCase && <div className="page"><section className="panel"><EmptyState title="案件が見つかりません" text="URLが正しいか、データがこのブラウザに保存されているか確認してください。" /><div className="not-found-action"><button className="primary-button" onClick={() => navigate("home")}>案件一覧へ</button></div></section></div>}
        </main>
      </div>
      {initiatorPickerOpen && <InitiatorPickerModal approvers={data.approvers} onSelect={selectCurrentUser} onClose={() => setInitiatorPickerOpen(false)} />}
      {pendingStart && <EmailComposerModal draft={pendingStart.draft} onClose={() => { setPendingStart(undefined); notify("メール送信をキャンセルしました。回覧は開始されていません"); }} onSend={(message) => { const created = pendingStart.item; setData((current) => ({ ...current, cases: [created, ...current.cases] })); setPendingStart(undefined); navigate("case", created.accessKey); notify(`${message.recipientName}さんへのメール送信を確認し、回覧を開始しました（デモ）`); }} />}
      {toast && <div className="toast">{toast}</div>}
    </div>
  );
}

function PageHeading({ title, meta, help, actions }: { title: string; meta?: string; help?: React.ReactNode; actions?: React.ReactNode }) {
  return <header className="page-heading"><div className="page-heading-copy"><div className="page-title-row"><h1>{title}</h1>{help && <HelpButton title={`${title}の使い方`}>{help}</HelpButton>}</div>{meta && <p className="page-meta">{meta}</p>}</div>{actions && <div className="page-actions">{actions}</div>}</header>;
}

function ApproversPage({ approvers, onChange, notify }: { approvers: Approver[]; onChange: (items: Approver[]) => void; notify: (message: string) => void }) {
  const [query, setQuery] = useState("");
  const [editing, setEditing] = useState<Approver | "new">();
  const [directoryOpen, setDirectoryOpen] = useState(false);
  const visible = approvers.filter((person) => [person.name, person.email, person.department].some((value) => value.toLowerCase().includes(query.toLowerCase())));

  const save = (values: { name: string; email: string; department: string; entraUserId: string }) => {
    if (editing === "new") {
      const timestamp = nowIso();
      onChange([...approvers, { id: uid("approver"), ...values, active: true, createdAt: timestamp, updatedAt: timestamp }]);
      notify(`${values.name}さんを登録しました`);
    } else if (editing) {
      onChange(approvers.map((person) => person.id === editing.id ? { ...person, ...values, updatedAt: nowIso() } : person));
      notify("承認者情報を更新しました");
    }
    setEditing(undefined);
  };

  const registerDirectoryUser = (person: DirectoryUser) => {
    const duplicate = approvers.find((item) => item.entraUserId === person.id || item.email.toLowerCase() === person.email.toLowerCase());
    if (duplicate) {
      notify(`${duplicate.name}さんは既に承認者マスタへ登録されています`);
      return;
    }
    const timestamp = nowIso();
    onChange([...approvers, {
      id: uid("approver"),
      entraUserId: person.id,
      name: person.name,
      email: person.email,
      department: person.department,
      active: true,
      createdAt: timestamp,
      updatedAt: timestamp,
    }]);
    setDirectoryOpen(false);
    notify(`${person.name}さんを社内名簿から登録しました`);
  };

  return <div className="page">
    <PageHeading title="承認者マスタ" help={<p>回付ルートで使う承認者を登録・編集します。Microsoft 365設定時は社内名簿から追加できます。</p>} actions={<div className="heading-actions">{directorySearchEnabled && <button className="quiet-button" onClick={() => setDirectoryOpen(true)}>社内名簿から追加</button>}<button className="primary-button" onClick={() => setEditing("new")}>承認者を追加</button></div>} />
    <section className="panel">
      <div className="toolbar"><div className="search-input large"><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="氏名・メールアドレス・部門で検索" /></div><span className="result-count">{visible.length}名</span></div>
      <div className="table-wrap"><table><thead><tr><th>氏名</th><th>メールアドレス</th><th>部門</th><th>状態</th><th>更新日時</th><th aria-label="操作" /></tr></thead><tbody>
        {visible.map((person) => <tr key={person.id} className={!person.active ? "muted-row" : ""}><td><div className="person-cell"><strong>{person.name}</strong></div></td><td><a href={`mailto:${person.email}`}>{person.email}</a></td><td>{person.department || "未設定"}</td><td><span className={`active-badge ${person.active ? "is-active" : ""}`}>{person.active ? "有効" : "無効"}</span></td><td>{formatDate(person.updatedAt)}</td><td><div className="row-actions"><button onClick={() => setEditing(person)}>編集</button><button onClick={() => { onChange(approvers.map((item) => item.id === person.id ? { ...item, active: !item.active, updatedAt: nowIso() } : item)); notify(`${person.name}さんを${person.active ? "無効化" : "有効化"}しました`); }}>{person.active ? "無効化" : "有効化"}</button><button className="danger-text" onClick={() => { if (window.confirm(`${person.name}さんを削除しますか？`)) { onChange(approvers.filter((item) => item.id !== person.id)); notify("承認者を削除しました"); } }}>削除</button></div></td></tr>)}
      </tbody></table></div>
      {visible.length === 0 && <EmptyState title="承認者が見つかりません" text="検索条件を変えてお試しください。" />}
    </section>
    {editing && <ApproverForm person={editing === "new" ? undefined : editing} onClose={() => setEditing(undefined)} onSave={save} />}
    {directoryOpen && <DirectoryUserModal onClose={() => setDirectoryOpen(false)} onSelect={registerDirectoryUser} />}
  </div>;
}

function DirectoryUserModal({ onClose, onSelect }: { onClose: () => void; onSelect: (person: DirectoryUser) => void }) {
  const [query, setQuery] = useState("");
  const [debounced, setDebounced] = useState("");
  const [results, setResults] = useState<DirectoryUser[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  useEffect(() => {
    const timer = window.setTimeout(() => setDebounced(query.trim()), 300);
    return () => window.clearTimeout(timer);
  }, [query]);
  useEffect(() => {
    if (!debounced) {
      setResults([]);
      setError("");
      setLoading(false);
      return;
    }
    const controller = new AbortController();
    setLoading(true);
    setError("");
    searchDirectoryUsers(debounced, controller.signal)
      .then(setResults)
      .catch((caught) => {
        if (caught instanceof DOMException && caught.name === "AbortError") return;
        setResults([]);
        setError(caught instanceof Error ? caught.message : "社内名簿を検索できませんでした。");
      })
      .finally(() => {
        if (!controller.signal.aborted) setLoading(false);
      });
    return () => controller.abort();
  }, [debounced]);

  return <Modal title="社内名簿から承認者を追加" onClose={onClose}><div className="modal-body directory-modal"><label className="field-label" htmlFor="directory-user-search">氏名・メール・部門で検索</label><div className="search-input"><input id="directory-user-search" value={query} onChange={(event) => setQuery(event.target.value)} placeholder="1文字以上入力" autoComplete="off" autoFocus /></div><div className="directory-results">{loading && <div className="suggestion-note">社内名簿を検索中…</div>}{error && <div className="suggestion-error">{error}</div>}{!loading && !error && debounced && results.length === 0 && <div className="suggestion-empty">該当する利用者はいません</div>}{results.map((person) => <button key={person.id} className="suggestion" onClick={() => onSelect(person)}><span><strong>{person.name}</strong><small>{person.department || "部門未設定"} ・ {person.email}</small></span><span className="suggestion-action">追加</span></button>)}</div></div><div className="modal-footer"><button className="quiet-button" onClick={onClose}>閉じる</button></div></Modal>;
}

function ApproverForm({ person, onClose, onSave }: { person?: Approver; onClose: () => void; onSave: (values: { name: string; email: string; department: string; entraUserId: string }) => void }) {
  const [name, setName] = useState(person?.name ?? "");
  const [email, setEmail] = useState(person?.email ?? "");
  const [department, setDepartment] = useState(person?.department ?? "");
  const valid = name.trim() && email.includes("@");
  return <Modal title={person ? "承認者を編集" : "承認者を新規登録"} onClose={onClose}><form onSubmit={(event) => { event.preventDefault(); if (valid) onSave({ name: name.trim(), email: email.trim(), department: department.trim(), entraUserId: person?.entraUserId ?? "" }); }}><div className="modal-body form-grid"><label><span>氏名 <b>必須</b></span><input value={name} onChange={(event) => setName(event.target.value)} placeholder="例：山田 太郎" autoFocus /></label><label><span>メールアドレス <b>必須</b></span><input type="email" value={email} onChange={(event) => setEmail(event.target.value)} placeholder="taro.yamada@example.com" /></label><label><span>部門 <em>任意</em></span><input value={department} onChange={(event) => setDepartment(event.target.value)} placeholder="例：技術部" /></label></div><div className="modal-footer"><button type="button" className="quiet-button" onClick={onClose}>キャンセル</button><button type="submit" className="primary-button" disabled={!valid}>{person ? "変更を保存" : "登録する"}</button></div></form></Modal>;
}

function NewCirculationPage({ data, currentUser, setTemplates, onStart, notify }: { data: AppData; currentUser: Approver; setTemplates: (items: RouteTemplate[]) => void; onStart: (item: CirculationCase) => void; notify: (message: string) => void }) {
  const createRouteMember = (person: Approver, isInitiator = false): RouteMember => ({ id: uid("route"), approverId: person.id, name: person.name, email: person.email, department: person.department, sequence: 1, requiresStamp: false, isInitiator, status: "pending" });
  const normalize = (items: RouteMember[]) => items.map((item, index) => ({ ...item, sequence: index + 1 }));
  const [circulationName, setCirculationName] = useState("");
  const [provider, setProvider] = useState<ProviderType>("sharepoint");
  const [selectedDocuments, setSelectedDocuments] = useState<CaseDocument[]>([]);
  const [filePicker, setFilePicker] = useState(false);
  const [selectedTemplateId, setSelectedTemplateId] = useState("");
  const [route, setRoute] = useState<RouteMember[]>(() => [createRouteMember(currentUser, true)]);
  const [saveModal, setSaveModal] = useState(false);
  const [draggedRouteId, setDraggedRouteId] = useState<string>();
  const [dragOverRouteId, setDragOverRouteId] = useState<string>();
  const hasStampDocuments = selectedDocuments.some((document) => document.requiresStamp);

  useEffect(() => {
    setRoute((current) => {
      const existing = current.find((member) => member.isInitiator);
      const initiator = { ...(existing ?? createRouteMember(currentUser, true)), approverId: currentUser.id, name: currentUser.name, email: currentUser.email, department: currentUser.department, requiresStamp: false, isInitiator: true };
      return normalize([initiator, ...current.filter((member) => !member.isInitiator)]);
    });
  }, [currentUser.id, currentUser.name, currentUser.email, currentUser.department]);
  useEffect(() => {
    if (!draggedRouteId) return;
    const stopDragging = () => {
      setDraggedRouteId(undefined);
      setDragOverRouteId(undefined);
    };
    window.addEventListener("mouseup", stopDragging);
    return () => window.removeEventListener("mouseup", stopDragging);
  }, [draggedRouteId]);

  const addApprover = (person: Approver) => {
    setRoute(normalize([...route, createRouteMember(person)]));
  };
  const move = (from: number, to: number) => {
    if (from <= 0 || to <= 0 || to >= route.length) return;
    const changed = [...route]; const [picked] = changed.splice(from, 1); changed.splice(to, 0, picked); setRoute(normalize(changed));
  };
  const reorderById = (sourceId: string, targetId: string) => {
    setRoute((current) => {
      const from = current.findIndex((member) => member.id === sourceId);
      const to = current.findIndex((member) => member.id === targetId);
      if (from <= 0 || to <= 0 || from === to) return current;
      const changed = [...current];
      const [picked] = changed.splice(from, 1);
      changed.splice(to, 0, picked);
      return normalize(changed);
    });
  };
  const loadTemplate = () => {
    const template = data.templates.find((item) => item.id === selectedTemplateId);
    if (!template) { notify("読み込むテンプレートを選択してください"); return; }
    const loaded = template.members.map((member, index) => {
      const latest = data.approvers.find((person) => person.id === member.approverId);
      return { ...member, id: uid("route"), sequence: index + 1, status: "pending" as const, name: latest?.name ?? member.name, email: latest?.email ?? member.email, department: latest?.department ?? member.department };
    });
    const initiator = route.find((member) => member.isInitiator) ?? createRouteMember(currentUser, true);
    setRoute(normalize([initiator, ...loaded]));
    notify(`「${template.name}」を読み込みました`);
  };
  const overwrite = () => {
    const target = data.templates.find((item) => item.id === selectedTemplateId);
    if (!target) return setSaveModal(true);
    setTemplates(data.templates.map((item) => item.id === target.id ? { ...item, updatedAt: nowIso(), members: route.filter((member) => !member.isInitiator).map(({ status: _status, completedAt: _completedAt, ...member }) => member) } : item));
    notify(`「${target.name}」を上書き保存しました`);
  };
  const duplicateTemplate = () => {
    const target = data.templates.find((item) => item.id === selectedTemplateId); if (!target) return;
    const copy = { ...target, id: uid("template"), name: `${target.name}（コピー）`, createdAt: nowIso(), updatedAt: nowIso(), members: target.members.map((item) => ({ ...item, id: uid("template-member") })) };
    setTemplates([...data.templates, copy]); setSelectedTemplateId(copy.id); notify("テンプレートを複製しました");
  };
  const recipients = route.filter((member) => !member.isInitiator);
  const canStart = selectedDocuments.length > 0 && recipients.length > 0;
  const start = () => {
    if (!canStart) return;
    const timestamp = nowIso();
    const normalizedRoute = normalize(route);
    const firstApprover = normalizedRoute.find((member) => !member.isInitiator)!;
    const members = normalizedRoute.map((item) => ({ ...item, id: uid("case-member"), requiresStamp: !item.isInitiator && hasStampDocuments && item.requiresStamp, status: item.isInitiator ? "approved" as const : item.id === firstApprover.id ? "current" as const : "pending" as const, completedAt: item.isInitiator ? timestamp : undefined }));
    const initiatorMember = members.find((member) => member.isInitiator)!;
    const currentMember = members.find((member) => member.approverId === firstApprover.approverId && !member.isInitiator)!;
    const id = `CASE-${new Date().getFullYear()}-${String(Date.now()).slice(-8)}`;
    const primary = selectedDocuments[0];
    onStart({ id, accessKey: crypto.randomUUID().replaceAll("-", ""), title: circulationName.trim() || undefined, provider, fileId: primary.fileId, fileName: primary.name, fileUrl: primary.fileUrl, documents: selectedDocuments, initiatorId: currentUser.id, initiatorName: currentUser.name, startedAt: timestamp, updatedAt: timestamp, state: "circulating", currentMemberId: currentMember.id, members, history: [{ id: uid("history"), caseId: id, actionUserId: currentUser.id, actionMemberId: initiatorMember.id, actionUserName: currentUser.name, action: "回覧開始", previousState: "下書き", newState: `回覧中（${currentMember.name}）`, createdAt: timestamp }] });
  };

  return <div className="page">
    <PageHeading title="新しい回覧" help={<ol><li>必要なら回覧名を入力し、文書ごとの捺印要否を設定します。</li><li>開始者はルートの先頭へ自動で入ります。同じ人の再確認が必要な場合は、その人を複数回追加できます。</li><li>承認者を追加し、ドラッグまたは「上へ／下へ」で並べ替えます。</li><li>順番と処理内容を確認して回覧を開始します。</li></ol>} />
    <div className="builder-main">
        <section className="panel form-section"><label className="circulation-name-field"><span className="field-label">回覧名（任意）</span><input value={circulationName} onChange={(event) => setCirculationName(event.target.value)} placeholder="未入力の場合は文書名を表示" maxLength={100} /><small>{circulationName.length} / 100</small></label><div className="number-title"><b>1</b><h2>対象文書</h2></div><div className="document-source"><label><span className="field-label">保存場所</span><select value={provider} onChange={(event) => { setProvider(event.target.value as ProviderType); setSelectedDocuments([]); }}><option value="sharepoint">SharePoint</option><option value="onedrive">OneDrive</option><option value="shared-folder">共有フォルダ</option></select></label><button className="secondary-button" onClick={() => setFilePicker(true)}>文書を選択</button></div>
          {selectedDocuments.length === 0 ? <div className="document-empty">文書が選択されていません</div> : <div className="selected-documents">{selectedDocuments.map((document) => <div className="selected-document" key={document.id}><span className={`file-type-label ${fileClass(document.type)}`}>{fileLabel(document.type)}</span><div className="selected-document-main"><strong>{document.name}</strong></div><div className="document-action-choice" aria-label={`${document.name}の処理`}><button className={document.requiresStamp ? "active stamp" : ""} onClick={() => setSelectedDocuments(selectedDocuments.map((item) => item.id === document.id ? { ...item, requiresStamp: true } : item))}>捺印対象</button><button className={!document.requiresStamp ? "active" : ""} onClick={() => setSelectedDocuments(selectedDocuments.map((item) => item.id === document.id ? { ...item, requiresStamp: false } : item))}>確認のみ</button></div><button className="document-remove" onClick={() => setSelectedDocuments(selectedDocuments.filter((item) => item.id !== document.id))}>削除</button></div>)}</div>}
        </section>
        <section className="panel form-section"><div className="number-title"><b>2</b><h2>回付ルート</h2></div><div className="template-bar"><select value={selectedTemplateId} onChange={(event) => setSelectedTemplateId(event.target.value)}><option value="">ルートテンプレートを選択</option>{data.templates.map((item) => <option key={item.id} value={item.id}>{item.name}</option>)}</select><button className="secondary-button" onClick={loadTemplate}>読み込み</button><span className="bar-divider" /><button className="text-button" disabled={!selectedTemplateId} onClick={overwrite}>上書き保存</button><button className="text-button" disabled={!selectedTemplateId} onClick={duplicateTemplate}>複製</button><button className="text-button danger-text" disabled={!selectedTemplateId} onClick={() => { const target = data.templates.find((item) => item.id === selectedTemplateId); if (target && window.confirm(`「${target.name}」を削除しますか？`)) { setTemplates(data.templates.filter((item) => item.id !== target.id)); setSelectedTemplateId(""); notify("テンプレートを削除しました"); } }}>削除</button></div>
          <ApproverAutocomplete approvers={data.approvers} onSelect={addApprover} />
          <div className="route-list">
            {route.map((member, index) => { const effectiveStamp = hasStampDocuments && member.requiresStamp; return <div key={member.id} className={`route-card ${member.isInitiator ? "initiator-route" : "draggable-route"} ${draggedRouteId === member.id ? "dragging" : ""} ${dragOverRouteId === member.id ? "drag-over" : ""}`} onMouseDown={(event) => { const target = event.target as HTMLElement; if (member.isInitiator || target.closest("button, input, label, a")) return; event.preventDefault(); setDraggedRouteId(member.id); }} onMouseEnter={() => { if (!member.isInitiator && draggedRouteId && draggedRouteId !== member.id) setDragOverRouteId(member.id); }} onMouseUp={() => { if (draggedRouteId && !member.isInitiator) reorderById(draggedRouteId, member.id); setDraggedRouteId(undefined); setDragOverRouteId(undefined); }}><span className="sequence">{index + 1}</span>{member.isInitiator ? <span className="route-type-label initiator">開始者</span> : <label className="route-type-control" title={hasStampDocuments ? (member.requiresStamp ? "確認のみに変更" : "捺印ありに変更") : "捺印対象文書がないため確認のみ"}><input type="checkbox" checked={effectiveStamp} disabled={!hasStampDocuments} onChange={() => setRoute(route.map((item) => item.id === member.id ? { ...item, requiresStamp: !item.requiresStamp } : item))} aria-label={`${member.name}の処理を${effectiveStamp ? "確認のみ" : "捺印あり"}へ変更`} /><RouteTypeLabel requiresStamp={effectiveStamp} /></label>}<div className="route-person"><strong>{member.name}</strong><small>{member.department || "部門未設定"} ・ {member.email}</small></div>{!member.isInitiator && <div className="move-actions"><button disabled={index <= 1} onClick={() => move(index, index - 1)}>上へ</button><button disabled={index === route.length - 1} onClick={() => move(index, index + 1)}>下へ</button><button className="remove" onClick={() => setRoute(normalize(route.filter((item) => item.id !== member.id)))}>削除</button></div>}</div>; })}
          </div>
          <div className="route-footer"><span>{recipients.length === 0 ? "開始者以外の回付者を追加してください" : `開始者を含む${route.length}名・開始者以外はドラッグで並べ替えできます`}</span><button className="secondary-button" disabled={recipients.length === 0} onClick={() => setSaveModal(true)}>ルートを名前を付けて保存</button></div>
        </section>
      <section className="panel start-section"><div><h2>回覧を開始</h2><p>{circulationName.trim() || "文書名を回覧名として使用"} ・ 文書 {selectedDocuments.length}件 ・ 回付者 {route.length}名</p>{!canStart && <small>{selectedDocuments.length === 0 ? "文書を選択してください" : "開始者以外の回付者を追加してください"}</small>}</div><button className="primary-button large-button" disabled={!canStart} onClick={start}>回覧を開始する</button></section>
    </div>
    {filePicker && <FilePicker provider={provider} selectedIds={selectedDocuments.map((item) => item.fileId)} onClose={() => setFilePicker(false)} onConfirm={(files) => { const existing = new Map(selectedDocuments.map((item) => [item.fileId, item])); setSelectedDocuments(files.map((file) => { const previous = existing.get(file.id); return { id: previous?.id ?? uid("document"), fileId: file.id, name: file.name, type: file.type, location: file.location, fileUrl: file.fileUrl, requiresStamp: previous?.requiresStamp ?? false }; })); setFilePicker(false); notify(`${files.length}件の文書を選択しました`); }} />}
    {saveModal && <TemplateSaveModal currentUser={currentUser} route={recipients} onClose={() => setSaveModal(false)} onSave={(template) => { setTemplates([...data.templates, template]); setSelectedTemplateId(template.id); setSaveModal(false); notify(`「${template.name}」を保存しました`); }} />}
  </div>;
}

function FilePicker({ provider, selectedIds, onClose, onConfirm }: { provider: ProviderType; selectedIds: string[]; onClose: () => void; onConfirm: (files: MockFile[]) => void }) {
  const useSharePoint = provider === "sharepoint" && sharePointFileListingEnabled;
  const [files, setFiles] = useState<MockFile[]>(() => useSharePoint ? [] : mockFiles.filter((item) => item.provider === provider));
  const [draftIds, setDraftIds] = useState(selectedIds);
  const [loading, setLoading] = useState(useSharePoint);
  const [error, setError] = useState("");
  const [sourceLabel, setSourceLabel] = useState("");
  const [reloadCount, setReloadCount] = useState(0);

  useEffect(() => {
    if (!useSharePoint) {
      setFiles(mockFiles.filter((item) => item.provider === provider));
      setLoading(false);
      setError("");
      setSourceLabel("");
      return;
    }
    const controller = new AbortController();
    setLoading(true);
    setError("");
    setSourceLabel("");
    listSharePointFiles(controller.signal)
      .then((result) => {
        setFiles(result.files);
        setSourceLabel([result.siteName, result.libraryName, result.folderPath].filter(Boolean).join(" / "));
      })
      .catch((caught) => {
        if (caught instanceof DOMException && caught.name === "AbortError") return;
        setFiles([]);
        setError(caught instanceof Error ? caught.message : "SharePointの文書一覧を取得できませんでした。");
      })
      .finally(() => {
        if (!controller.signal.aborted) setLoading(false);
      });
    return () => controller.abort();
  }, [provider, reloadCount, useSharePoint]);

  const toggle = (id: string) => setDraftIds((current) => current.includes(id) ? current.filter((item) => item !== id) : [...current, id]);
  return <Modal title="文書を選択" onClose={onClose} wide><div className="modal-body">{useSharePoint ? sourceLabel && <div className="source-banner">SharePoint：{sourceLabel}</div> : <div className="demo-banner">現在は確認用のダミーファイルです。複数選択できます。</div>}{loading && <div className="file-picker-state">SharePointから文書を取得しています…</div>}{error && <div className="file-picker-state error"><span>{error}</span><button className="secondary-button" onClick={() => setReloadCount((value) => value + 1)}>再読み込み</button></div>}{!loading && !error && files.length === 0 && <div className="file-picker-state">表示できる文書がありません</div>}{!loading && !error && <div className="file-grid">{files.map((file) => { const selected = draftIds.includes(file.id); return <button key={file.id} className={`file-choice ${selected ? "selected" : ""}`} onClick={() => toggle(file.id)}><span className={`file-type-label large ${fileClass(file.type)}`}>{fileLabel(file.type)}</span><span><strong>{file.name}</strong><small>更新：{file.updatedAt ? formatDate(file.updatedAt) : "日時不明"} ・ {file.updatedBy}</small></span><em>{selected ? "選択済み" : "選択"}</em></button>; })}</div>}</div><div className="modal-footer"><span className="selection-count">{draftIds.filter((id) => files.some((file) => file.id === id)).length}件選択</span><button className="quiet-button" onClick={onClose}>キャンセル</button><button className="primary-button" disabled={loading || Boolean(error) || !draftIds.some((id) => files.some((file) => file.id === id))} onClick={() => onConfirm(files.filter((file) => draftIds.includes(file.id)))}>選択を確定</button></div></Modal>;
}

function TemplateSaveModal({ currentUser, route, onClose, onSave }: { currentUser: Approver; route: RouteMember[]; onClose: () => void; onSave: (template: RouteTemplate) => void }) {
  const [name, setName] = useState(""); const [description, setDescription] = useState("");
  return <Modal title="ルートを名前を付けて保存" onClose={onClose}><form onSubmit={(event) => { event.preventDefault(); if (!name.trim()) return; const timestamp = nowIso(); onSave({ id: uid("template"), name: name.trim(), description: description.trim(), createdBy: currentUser.name, createdAt: timestamp, updatedAt: timestamp, members: route.map(({ status: _status, completedAt: _completedAt, ...member }) => ({ ...member, id: uid("template-member") })) }); }}><div className="modal-body form-grid"><label><span>テンプレート名 <b>必須</b></span><input value={name} onChange={(event) => setName(event.target.value)} placeholder="例：部長承認ルート" autoFocus /></label><label><span>説明</span><textarea value={description} onChange={(event) => setDescription(event.target.value)} placeholder="このルートを使う場面を入力" rows={3} /></label><div className="save-preview"><strong>{route.length}名を保存します</strong><span>{route.map((item) => item.name).join(" → ")}</span></div></div><div className="modal-footer"><button type="button" className="quiet-button" onClick={onClose}>キャンセル</button><button type="submit" className="primary-button" disabled={!name.trim()}>保存する</button></div></form></Modal>;
}

function CasesPage({ cases, onOpen, newWindowUrl, approversWindowUrl, onReset }: {
  cases: CirculationCase[];
  onOpen: (accessKey: string) => void;
  newWindowUrl: string;
  approversWindowUrl: string;
  onReset: () => void;
}) {
  const [tab, setTab] = useState<CaseTab>("waiting");
  const [query, setQuery] = useState("");
  const groups = useMemo(() => ({
    waiting: cases.filter((item) => item.state !== "completed" && (item.currentMemberId || item.returnedToStarter)),
    all: cases,
    returned: cases.filter((item) => item.state === "returned"),
    completed: cases.filter((item) => item.state === "completed"),
  }), [cases]);
  const visible = groups[tab].filter((item) => caseTitle(item).toLowerCase().includes(query.toLowerCase()) || item.documents.some((document) => document.name.toLowerCase().includes(query.toLowerCase())) || item.id.toLowerCase().includes(query.toLowerCase()));
  const showRowStatus = tab === "waiting" || tab === "all";
  const tabLabels: Record<CaseTab, string> = { waiting: "承認待ち", all: "すべて", returned: "差し戻し", completed: "完了" };
  return <div className="page home-page"><PageHeading title="ホーム" help={<><p>案件は状態タブまたは検索で絞り込めます。案件を選ぶと詳細を表示します。</p><p>「新しい回覧」と「承認者マスタ」は別ウィンドウで開きます。</p><div className="help-actions"><button className="quiet-button" onClick={onReset}>デモデータを初期化</button></div></>} actions={<div className="home-actions"><a className="primary-button home-action-link" href={newWindowUrl} target="_blank" rel="noreferrer">新しい回覧</a><a className="secondary-button home-action-link" href={approversWindowUrl} target="_blank" rel="noreferrer">承認者マスタ</a></div>} />
    <div className="cases-heading"><h2>案件一覧</h2></div>
    <div className="tabs" aria-label="案件の状態">{(["waiting", "all", "returned", "completed"] as CaseTab[]).map((key) => <button type="button" key={key} className={tab === key ? "active" : ""} aria-pressed={tab === key} onMouseDown={(event) => event.preventDefault()} onClick={() => setTab(key)}><span>{tabLabels[key]}</span><span className="tab-count">{groups[key].length}</span></button>)}</div>
    <section className="panel"><div className="toolbar"><div className="search-input large"><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="ファイル名・案件IDで検索" /></div></div>
      {visible.length ? <div className="case-cards">{visible.map((item) => { const current = item.members.find((member) => member.id === item.currentMemberId); return <button key={item.id} className={`case-card ${showRowStatus ? "with-status" : ""}`} onClick={() => onOpen(item.id)}><div className="case-card-main"><h3>{caseTitle(item)}</h3><small>{item.id} ・ 文書{item.documents.length}件</small></div>{showRowStatus && <StatusBadge state={item.state} />}<div className="current-owner"><small>{item.state === "completed" ? "完了日時" : "現在の担当"}</small><strong>{item.state === "completed" ? formatDate(item.updatedAt) : current?.name ?? "回覧開始者"}</strong></div></button>; })}</div> : <EmptyState title="該当する案件はありません" text="別のタブまたは検索条件をお試しください。" />}
    </section>
  </div>;
}

function CaseDetail({ item, approvers, onBack, onChange, notify }: { item: CirculationCase; approvers: Approver[]; onBack: () => void; onChange: (item: CirculationCase) => void; notify: (message: string) => void }) {
  const [rejecting, setRejecting] = useState(false);
  const [editingRoute, setEditingRoute] = useState(false);
  const [pendingTransition, setPendingTransition] = useState<PendingCaseTransition>();
  const current = item.members.find((member) => member.id === item.currentMemberId)!;
  const canProcess = item.state !== "completed" && Boolean(current);
  const canRestart = item.state === "returned" && Boolean(item.returnedToStarter);
  const canEdit = item.state !== "completed";
  const title = caseTitle(item);
  const stampDocuments = item.documents.filter((document) => document.requiresStamp);
  const needsStampNow = Boolean(current?.requiresStamp && stampDocuments.length > 0);

  const approve = () => {
    if (!current) return;
    const index = item.members.findIndex((member) => member.id === current.id);
    const timestamp = nowIso(); const next = item.members[index + 1];
    const members = item.members.map((member, memberIndex) => memberIndex === index ? { ...member, status: needsStampNow ? "stamped" as const : "approved" as const, completedAt: timestamp } : memberIndex === index + 1 ? { ...member, status: "current" as const } : member);
    const action = needsStampNow ? "捺印・確認完了" : "確認完了";
    const changed: CirculationCase = { ...item, members, state: next ? "circulating" : "completed", currentMemberId: next?.id, returnedToStarter: false, updatedAt: timestamp, history: [...item.history, { id: uid("history"), caseId: item.id, actionUserId: current.approverId, actionMemberId: current.id, actionUserName: current.name, action, previousState: `${current.name}が処理中`, newState: next ? `${next.name}が処理中` : "回覧完了", createdAt: timestamp }] };
    if (next) {
      const nextNeedsStamp = next.requiresStamp && stampDocuments.length > 0;
      setPendingTransition({ changed, successMessage: `${next.name}さんへ回付しました`, draft: { to: next.email, cc: createPastApproverCc(item, current.id, next.email), recipientName: next.name, purpose: "次の確認者への回付通知", subject: `【承認依頼】${title}`, body: `${next.name}さん\n\n${current.name}さんの処理が完了し、あなたの承認待ちになりました。\n\n案件名：${title}\n文書数：${item.documents.length}件\n案件ID：${item.id}\n現在の確認者：${next.name}\n対応内容：${nextNeedsStamp ? `捺印対象${stampDocuments.length}件への捺印と全体確認` : "全文書の内容確認"}\n\n案件URL：\n${window.location.href}\n\nよろしくお願いいたします。` } });
    } else {
      const initiator = approvers.find((person) => person.id === item.initiatorId) ?? item.members.find((member) => member.approverId === item.initiatorId || member.isInitiator);
      if (!initiator?.email) {
        notify("回付開始者のメールアドレスを確認できないため、完了通知を作成できませんでした");
        return;
      }
      setPendingTransition({ changed, successMessage: "回覧を完了しました", draft: { to: initiator.email, cc: createPastApproverCc(item, current.id, initiator.email), recipientName: initiator.name, purpose: "回覧完了通知", subject: `【回覧完了】${title}`, body: `${initiator.name}さん\n\n以下の案件の承認回覧が完了しました。\n\n案件名：${title}\n文書数：${item.documents.length}件\n案件ID：${item.id}\n最終確認者：${current.name}\n完了日時：${formatDate(timestamp)}\n\n案件URL：\n${window.location.href}\n\n回覧履歴は案件ページから確認できます。` } });
    }
  };
  const restart = () => {
    const first = item.members.find((member) => !member.isInitiator) ?? item.members[0]; const timestamp = nowIso();
    if (!first?.email) {
      notify("先頭の承認者のメールアドレスを確認できないため、再回覧通知を作成できませんでした");
      return;
    }
    const changed: CirculationCase = { ...item, state: "circulating", returnedToStarter: false, currentMemberId: first.id, updatedAt: timestamp, members: item.members.map((member) => member.isInitiator ? { ...member, status: "approved" as const, completedAt: member.completedAt ?? timestamp } : { ...member, status: member.id === first.id ? "current" as const : "pending" as const, completedAt: undefined }), history: [...item.history, { id: uid("history"), caseId: item.id, actionUserId: item.initiatorId, actionMemberId: item.members.find((member) => member.isInitiator)?.id, actionUserName: item.initiatorName, action: "再回覧", previousState: "回覧開始者へ差し戻し", newState: `${first.name}が処理中`, createdAt: timestamp }] };
    const firstNeedsStamp = first.requiresStamp && stampDocuments.length > 0;
    setPendingTransition({ changed, successMessage: `${first.name}さんから再回覧を開始しました`, draft: { to: first.email, cc: createPastApproverCc(item, undefined, first.email), recipientName: first.name, purpose: "再回覧通知", subject: `【承認依頼・再回覧】${title}`, body: `${first.name}さん\n\n${item.initiatorName}さんから案件が再回覧され、あなたの承認待ちになりました。\n\n案件名：${title}\n文書数：${item.documents.length}件\n案件ID：${item.id}\n現在の確認者：${first.name}\n対応内容：${firstNeedsStamp ? `捺印対象${stampDocuments.length}件への捺印と全体確認` : "全文書の内容確認"}\n\n案件URL：\n${window.location.href}\n\n内容を確認し、案件ページから対応してください。` } });
  };
  const progressMembers = item.members.filter((member) => !member.isInitiator);
  const doneCount = progressMembers.filter((member) => member.status === "approved" || member.status === "stamped").length;
  const documentOwnerId = current?.id ?? (canRestart ? item.members.find((member) => member.isInitiator)?.id : item.members.at(-1)?.id);
  const routeSegments = buildRouteTimelineSegments(item);
  return <div className="page case-detail-page">
    <button className="back-button" onClick={onBack}>案件一覧へ戻る</button>
    <PageHeading title={title} meta={`${item.id} ・ 文書${item.documents.length}件 ・ ${providerLabels[item.provider]}`} help={<p>水色で強調された現在の承認待ち者の欄で、対象文書の確認と処理を行います。差し戻しコメントは任意です。メールの「送信を確認」を押したときに処理結果と回付ルートを反映します。</p>} />
    <section className="panel progress-panel route-workflow-panel"><div className="section-title"><h2>回付ルート・履歴</h2>{canEdit && <button className="secondary-button" onClick={() => setEditingRoute(true)}>未処理ルートを編集</button>}</div>
      <div className="progress-summary"><span><i style={{ width: `${item.state === "completed" ? 100 : Math.round(doneCount / Math.max(progressMembers.length, 1) * 100)}%` }} /></span><p>{doneCount}/{progressMembers.length}名 完了</p></div>
      <div className="timeline">
        {routeSegments.map((segment, segmentIndex) => <div className="route-timeline-segment" key={segment.id}>
          {segmentIndex > 0 && <div className="route-repeat-divider"><strong>差し戻し後の回付ルート</strong><span>{segment.members[0]?.name ?? "戻り先"}から再開</span></div>}
          {segment.members.map((member, memberIndex) => {
            const memberHistory = item.history.filter((entry) => (entry.actionMemberId ? entry.actionMemberId === member.id : entry.actionUserId === member.approverId) && !["回覧開始", "再回覧", "回付ルート変更"].includes(entry.action));
            const isReturner = Boolean(segment.returnEntry && (segment.returnEntry.actionMemberId ? segment.returnEntry.actionMemberId === member.id : segment.returnEntry.actionUserId === member.approverId) && memberIndex === segment.members.length - 1);
            const isRepeatedInitiator = segment.current && segmentIndex > 0 && member.isInitiator && item.history.some((entry) => entry.action === "再回覧");
            const restartEntry = isRepeatedInitiator ? item.history.filter((entry) => entry.action === "再回覧").at(-1) : undefined;
            const processedAt = isReturner ? segment.returnEntry?.createdAt : restartEntry?.createdAt ?? member.completedAt ?? (routeSegments.length === 1 ? memberHistory.filter((entry) => !entry.action.includes("NG")).at(-1)?.createdAt : undefined);
            const showsDocuments = segment.current && member.id === documentOwnerId;
            const isWaiting = showsDocuments && item.state !== "completed";
            const workHeading = canRestart && isWaiting ? "差し戻し内容と対象文書を確認してください" : canProcess && isWaiting ? (needsStampNow ? "捺印対象文書へ捺印し、確認してください" : "対象文書を確認してください") : "対象文書";
            const stateText = isReturner ? "差し戻し" : canRestart && isWaiting ? "差し戻し対応" : isRepeatedInitiator ? "再回覧" : member.isInitiator ? "回覧開始" : member.status === "stamped" ? "捺印・確認済み" : member.status === "approved" ? "確認済み" : member.status === "current" ? "承認待ち" : "未処理";
            return <div key={`${segment.id}-${member.id}`} className={`timeline-item ${member.status} ${isWaiting ? "current-waiting" : ""} ${isReturner ? "returned-step" : ""}`}><div className="timeline-marker"><span>{segment.startSequence + memberIndex}</span>{memberIndex < segment.members.length - 1 && <i />}</div>{member.isInitiator ? <span className="route-type-label initiator">開始者</span> : <RouteTypeLabel requiresStamp={member.requiresStamp && stampDocuments.length > 0} />}<div className="timeline-person"><div className="timeline-person-heading"><div><strong>{member.name}</strong><small>{member.department || "部門未設定"}</small></div><div className="timeline-state"><strong>{stateText}</strong>{!isWaiting && processedAt && <small>{formatDate(processedAt)}</small>}</div></div>{showsDocuments && <div className="member-work"><div className="member-work-heading"><h3>{workHeading}</h3><span>{item.documents.length}件</span>{isWaiting && <HelpButton title={canRestart ? "再回覧の方法" : "処理方法"} label="操作方法"><p>{canRestart ? "差し戻し内容と文書を確認してから、先頭の承認者へ再回覧します。" : needsStampNow ? `「捺印対象」の${stampDocuments.length}件へ捺印し、すべての文書を確認してください。` : `対象文書${item.documents.length}件を確認してください。`}</p><p>通知メールの「送信を確認」を押したときに処理結果と回付ルートを反映します。</p></HelpButton>}</div><div className="member-document-list">{item.documents.map((document) => { const stampRequired = Boolean(member.requiresStamp && document.requiresStamp); return <div className="member-document-row" key={document.id}><span className={`file-type-label ${fileClass(document.type)}`}>{fileLabel(document.type)}</span><strong>{document.name}</strong>{stampRequired && <span className="document-kind stamp-kind">捺印対象</span>}<button className="document-button" onClick={() => notify(`確認用：${document.name} を開きました`)}>文書を開く</button></div>; })}</div>{canProcess && isWaiting && <div className="member-work-actions"><button className="danger-button" onClick={() => setRejecting(true)}>差し戻し</button><button className="primary-button" onClick={approve}>{needsStampNow ? "捺印・確認完了して回付" : "確認完了して回付"}</button></div>}{canRestart && isWaiting && <div className="member-work-actions"><button className="primary-button" onClick={restart}>先頭から再回覧</button></div>}</div>}</div></div>;
          })}
        </div>)}
      </div>
    </section>
    {rejecting && current && <RejectModal item={item} current={current} user={{ id: current.approverId, name: current.name }} onClose={() => setRejecting(false)} onSubmit={(changed) => { setRejecting(false); const destination = changed.returnedToStarter ? approvers.find((person) => person.id === changed.initiatorId) ?? changed.members.find((member) => member.isInitiator) : changed.members.find((member) => member.id === changed.currentMemberId); const latest = changed.history.at(-1); if (!destination?.email) { notify("差し戻し先のメールアドレスを確認できないため、通知を作成できませんでした"); return; } setPendingTransition({ changed, successMessage: `${destination.name}さんへ差し戻しました`, draft: { to: destination.email, cc: createPastApproverCc(item, current.id, destination.email), recipientName: destination.name, purpose: "差し戻し通知", subject: `【差し戻し】${caseTitle(changed)}`, body: `${destination.name}さん\n\n${current.name}さんから案件が差し戻されました。\n\n案件名：${caseTitle(changed)}\n文書数：${changed.documents.length}件\n案件ID：${changed.id}\n戻り先：${destination.name}\nコメント：${latest?.comment || "なし"}\n\n案件URL：\n${window.location.href}\n\n内容を確認し、案件ページから対応してください。` } }); }} />}
    {editingRoute && <RouteEditModal item={item} approvers={approvers} onClose={() => setEditingRoute(false)} onSave={(changed) => { onChange(changed); setEditingRoute(false); notify("未処理の回付ルートを更新しました"); }} notify={notify} />}
    {pendingTransition && <EmailComposerModal draft={pendingTransition.draft} onClose={() => { setPendingTransition(undefined); notify("メール送信をキャンセルしました。処理結果と回付ルートは変更されていません"); }} onSend={() => { const pending = pendingTransition; onChange(pending.changed); setPendingTransition(undefined); notify(`メール送信を確認しました。${pending.successMessage}（デモ）`); }} />}
  </div>;
}

function RejectModal({ item, current, user, onClose, onSubmit }: { item: CirculationCase; current: RouteMember; user: { id: string; name: string }; onClose: () => void; onSubmit: (changed: CirculationCase) => void }) {
  const currentIndex = item.members.findIndex((member) => member.id === current.id);
  const prior = item.members.slice(0, currentIndex).filter((member) => !member.isInitiator && (member.status === "approved" || member.status === "stamped"));
  const [target, setTarget] = useState("starter"); const [comment, setComment] = useState("");
  const submit = () => {
    const timestamp = nowIso(); const isStarter = target === "starter"; const returnMember = item.members.find((member) => member.id === target);
    const members = isStarter ? item.members.map((member) => ({ ...member, status: member.isInitiator ? "current" as const : "pending" as const, completedAt: member.isInitiator ? member.completedAt : undefined })) : item.members.map((member, index) => {
      const returnIndex = item.members.findIndex((candidate) => candidate.id === target);
      if (index < returnIndex) return member;
      return { ...member, status: index === returnIndex ? "current" as const : "pending" as const, completedAt: undefined };
    });
    onSubmit({ ...item, state: "returned", returnedToStarter: isStarter, currentMemberId: isStarter ? undefined : returnMember?.id, members, updatedAt: timestamp, history: [...item.history, { id: uid("history"), caseId: item.id, actionUserId: user.id, actionMemberId: current.id, actionUserName: user.name, action: "NG・差し戻し", comment: comment.trim() || undefined, returnToUserId: isStarter ? item.initiatorId : returnMember?.approverId, returnToMemberId: isStarter ? item.members.find((member) => member.isInitiator)?.id : returnMember?.id, returnToUserName: isStarter ? item.initiatorName : returnMember?.name, previousState: `${current.name}が処理中`, newState: `${isStarter ? item.initiatorName : returnMember?.name}へ差し戻し`, createdAt: timestamp, routeSnapshot: structuredClone(item.members) }] });
  };
  return <Modal title="NG・差し戻し" onClose={onClose}><div className="modal-body"><div className="warning-box"><p>差し戻し先を選んでください。コメントは任意です。処理済みの後続メンバーは未処理に戻ります。</p></div><fieldset className="return-options"><legend>戻り先</legend><label><input type="radio" name="target" value="starter" checked={target === "starter"} onChange={(event) => setTarget(event.target.value)} /><span><strong>{item.initiatorName}</strong><small>回覧開始者</small></span></label>{prior.map((member) => <label key={member.id}><input type="radio" name="target" value={member.id} checked={target === member.id} onChange={(event) => setTarget(event.target.value)} /><span><strong>{member.name}</strong><small>回付順 {member.sequence} ・ {member.department || "部門未設定"} ・ 処理済み</small></span></label>)}</fieldset><label className="comment-field"><span>コメント（任意）</span><textarea value={comment} onChange={(event) => setComment(event.target.value)} placeholder="必要に応じてコメントを入力してください" rows={4} autoFocus /><small>{comment.length} / 500</small></label></div><div className="modal-footer"><button className="quiet-button" onClick={onClose}>キャンセル</button><button className="danger-button solid" onClick={submit}>差し戻す</button></div></Modal>;
}

function RouteEditModal({ item, approvers, onClose, onSave, notify }: { item: CirculationCase; approvers: Approver[]; onClose: () => void; onSave: (changed: CirculationCase) => void; notify: (message: string) => void }) {
  const [members, setMembers] = useState<RouteMember[]>(structuredClone(item.members));
  const currentIndex = members.findIndex((member) => member.id === item.currentMemberId);
  const hasStampDocuments = item.documents.some((document) => document.requiresStamp);
  const normalize = (items: RouteMember[]) => items.map((member, index) => ({ ...member, sequence: index + 1 }));
  const add = (person: Approver) => setMembers(normalize([...members, { id: uid("case-member"), approverId: person.id, name: person.name, email: person.email, department: person.department, sequence: members.length + 1, requiresStamp: false, status: "pending" }]));
  const move = (index: number, direction: -1 | 1) => {
    const target = index + direction;
    if (target <= currentIndex || target >= members.length || members[target].status !== "pending") return;
    const changed = [...members]; [changed[index], changed[target]] = [changed[target], changed[index]]; setMembers(normalize(changed));
  };
  const changedSummary = JSON.stringify(members.map(({ id: _id, ...member }) => member)) !== JSON.stringify(item.members.map(({ id: _id, ...member }) => member));
  return <Modal title="未処理の回付ルートを編集" onClose={onClose} wide><div className="modal-body"><div className="demo-banner neutral"><span>編集範囲</span>処理済み承認者は履歴として固定されます。現在処理中の承認者は削除できません。</div><ApproverAutocomplete approvers={approvers} onSelect={add} label="未処理の末尾へ承認者を追加" /><div className="route-list edit-route-list">{members.map((member, index) => { const editable = member.status === "pending"; const effectiveStamp = hasStampDocuments && member.requiresStamp; return <div key={member.id} className={`route-card ${!editable ? "locked" : ""}`}><span className="sequence">{index + 1}</span>{member.isInitiator ? <span className="route-type-label initiator">開始者</span> : hasStampDocuments ? <label className="route-type-control" title={editable ? (member.requiresStamp ? "確認のみに変更" : "捺印ありに変更") : "処理済み・処理中の設定は変更できません"}><input type="checkbox" checked={effectiveStamp} disabled={!editable} onChange={() => setMembers(members.map((candidate) => candidate.id === member.id ? { ...candidate, requiresStamp: !candidate.requiresStamp } : candidate))} aria-label={`${member.name}の処理を${effectiveStamp ? "確認のみ" : "捺印あり"}へ変更`} /><RouteTypeLabel requiresStamp={effectiveStamp} /></label> : <RouteTypeLabel requiresStamp={false} />}<div className="route-person"><strong>{member.name}</strong><small>{member.department} ・ {member.isInitiator ? "回覧開始者（固定）" : member.status === "pending" ? "未処理" : member.status === "current" ? "処理中（固定）" : "処理済み（固定）"}</small></div><div className="move-actions"><button disabled={!editable || index - 1 <= currentIndex || members[index - 1]?.status !== "pending"} onClick={() => move(index, -1)}>上へ</button><button disabled={!editable || index === members.length - 1 || members[index + 1]?.status !== "pending"} onClick={() => move(index, 1)}>下へ</button><button className="remove" disabled={!editable} title={!editable ? "処理済み・処理中の承認者は削除できません" : "削除"} onClick={() => { if (!editable) { notify("現在処理中または処理済みの承認者は削除できません"); return; } setMembers(normalize(members.filter((candidate) => candidate.id !== member.id))); }}>削除</button></div></div>; })}</div></div><div className="modal-footer"><button className="quiet-button" onClick={onClose}>キャンセル</button><button className="primary-button" disabled={!changedSummary} onClick={() => { const timestamp = nowIso(); onSave({ ...item, members, updatedAt: timestamp, history: [...item.history, { id: uid("history"), caseId: item.id, actionUserId: item.initiatorId, actionUserName: item.initiatorName, action: "回付ルート変更", comment: "未処理部分の追加・削除・順番・捺印要否を更新", previousState: `${item.members.length}名`, newState: `${members.length}名`, createdAt: timestamp }] }); }}>変更を保存</button></div></Modal>;
}
