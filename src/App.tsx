import { useEffect, useMemo, useState } from "react";
import { mockFiles } from "./mockData";
import { loadData, resetData, saveData } from "./storage";
import type {
  AppData,
  Approver,
  CirculationCase,
  MockFile,
  ProviderType,
  RouteMember,
  RouteTemplate,
} from "./types";

type Page = "home" | "approvers" | "new" | "cases" | "case";
type CaseTab = "waiting" | "all" | "returned" | "completed";
type EmailDraft = { to: string; recipientName: string; subject: string; body: string; purpose: string };

function readLocation(): { page: Page; caseId?: string } {
  const path = window.location.pathname.replace(/\/$/, "") || "/";
  if (path.startsWith("/c/")) return { page: "case", caseId: decodeURIComponent(path.slice("/c/".length)) };
  if (path.startsWith("/cases/")) return { page: "case", caseId: decodeURIComponent(path.slice("/cases/".length)) };
  if (path === "/cases") return { page: "cases" };
  if (path === "/circulations/new") return { page: "new" };
  if (path === "/approvers") return { page: "approvers" };
  return { page: "home" };
}

function pathFor(page: Page, caseId?: string) {
  if (page === "case" && caseId) return `/c/${encodeURIComponent(caseId)}`;
  return ({ home: "/", approvers: "/approvers", new: "/circulations/new", cases: "/cases", case: "/cases" } as Record<Page, string>)[page];
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

function Modal({ title, children, onClose, wide = false }: { title: string; children: React.ReactNode; onClose: () => void; wide?: boolean }) {
  return (
    <div className="modal-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && onClose()}>
      <section className={`modal ${wide ? "modal-wide" : ""}`} role="dialog" aria-modal="true" aria-label={title}>
        <header className="modal-header"><h2>{title}</h2><button className="icon-button" onClick={onClose} aria-label="閉じる">×</button></header>
        {children}
      </section>
    </div>
  );
}

function EmailComposerModal({ draft, onClose, onSend }: { draft: EmailDraft; onClose: () => void; onSend: (message: EmailDraft) => void }) {
  const [to, setTo] = useState(draft.to);
  const [subject, setSubject] = useState(draft.subject);
  const [body, setBody] = useState(draft.body);
  const valid = to.includes("@") && subject.trim() && body.trim();
  return <Modal title="メールを作成" onClose={onClose} wide><div className="modal-body email-composer"><div className="demo-banner"><span>STEP 1</span>現在は送信UIのデモです。本接続後、この送信ボタンをMicrosoft Graphへ接続します。</div><div className="mail-purpose"><span>送信目的</span><strong>{draft.purpose}</strong></div><div className="mail-fields"><label><span>差出人</span><div className="sender-field">社内承認回覧 &lt;approval-circulation@example.com&gt;<em>本接続後に設定</em></div></label><label><span>宛先</span><input type="email" value={to} onChange={(event) => setTo(event.target.value)} /></label><label><span>件名</span><input value={subject} onChange={(event) => setSubject(event.target.value)} /></label><label><span>本文</span><textarea value={body} onChange={(event) => setBody(event.target.value)} rows={14} /></label></div><div className="mail-recipient-note"><Avatar name={draft.recipientName} small /><span><strong>{draft.recipientName}さん宛て</strong><small>送信前に宛先・件名・本文を編集できます</small></span></div></div><div className="modal-footer"><button className="quiet-button" onClick={onClose}>キャンセル</button><button className="primary-button send-button" disabled={!valid} onClick={() => onSend({ ...draft, to: to.trim(), subject: subject.trim(), body: body.trim() })}>送信（デモ）</button></div></Modal>;
}

function Avatar({ name, small = false }: { name: string; small?: boolean }) {
  return <span className={`avatar ${small ? "avatar-small" : ""}`} aria-hidden="true">{name.replace(/\s/g, "").slice(0, 1)}</span>;
}

function StatusBadge({ state }: { state: CirculationCase["state"] }) {
  return <span className={`status-badge status-${state}`}><span className="status-dot" />{stateLabel(state)}</span>;
}

function EmptyState({ icon, title, text }: { icon: string; title: string; text: string }) {
  return <div className="empty-state"><span className="empty-icon">{icon}</span><h3>{title}</h3><p>{text}</p></div>;
}

function ApproverAutocomplete({ approvers, excludedIds, onSelect, label = "承認者を検索" }: {
  approvers: Approver[]; excludedIds: string[]; onSelect: (person: Approver) => void; label?: string;
}) {
  const [query, setQuery] = useState("");
  const [debounced, setDebounced] = useState("");
  useEffect(() => {
    const timer = window.setTimeout(() => setDebounced(query.trim().toLowerCase()), 300);
    return () => window.clearTimeout(timer);
  }, [query]);
  const results = debounced.length < 1 ? [] : approvers.filter((person) =>
    person.active &&
    [person.name, person.email, person.department].some((value) => value.toLowerCase().includes(debounced)),
  ).slice(0, 6);

  return (
    <div className="autocomplete">
      <label className="field-label" htmlFor="approver-autocomplete">{label}</label>
      <div className="search-input"><span aria-hidden="true">⌕</span><input id="approver-autocomplete" value={query} onChange={(event) => setQuery(event.target.value)} placeholder="氏名・メール・部門を1文字以上入力" autoComplete="off" /></div>
      {debounced && <div className="suggestions">
        {results.length > 0 ? results.map((person) => (
          <button key={person.id} className={`suggestion ${excludedIds.includes(person.id) ? "already-added" : ""}`} onClick={() => { if (excludedIds.includes(person.id)) { window.alert("同じ承認者は重複して追加できません。"); return; } onSelect(person); setQuery(""); setDebounced(""); }}>
            <Avatar name={person.name} small />
            <span><strong>{person.name}</strong><small>{person.department} ・ {person.email}</small></span>
            <span className="add-mark">{excludedIds.includes(person.id) ? "追加済み" : "＋"}</span>
          </button>
        )) : <div className="suggestion-empty">該当する有効な承認者はいません</div>}
      </div>}
    </div>
  );
}

export default function App() {
  const [data, setData] = useState<AppData>(() => loadData());
  const initialLocation = useMemo(() => readLocation(), []);
  const [page, setPage] = useState<Page>(initialLocation.page);
  const [selectedCaseId, setSelectedCaseId] = useState<string | undefined>(initialLocation.caseId);
  const [toast, setToast] = useState<string>();

  useEffect(() => saveData(data), [data]);
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
  const incomingCount = data.cases.filter((item) => item.state !== "completed" && (item.currentMemberId || item.returnedToStarter)).length;

  const navigate = (next: Page, caseId?: string) => {
    const addressKey = next === "case" ? (data.cases.find((item) => item.id === caseId)?.accessKey ?? caseId) : caseId;
    window.history.pushState({}, "", pathFor(next, addressKey));
    setPage(next);
    setSelectedCaseId(next === "case" ? addressKey : undefined);
    window.scrollTo({ top: 0, behavior: "smooth" });
  };
  const notify = (message: string) => setToast(message);
  const setCurrentUser = (id: string) => setData((current) => ({ ...current, currentUserId: id }));
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
      <aside className="sidebar">
        <div className="brand"><span className="brand-mark">承</span><div><strong>社内承認回覧</strong></div></div>
        <nav aria-label="メインメニュー">
          <button className={page === "home" ? "active" : ""} onClick={() => navigate("home")}><span>⌂</span>ホーム</button>
          <button className={page === "new" ? "active" : ""} onClick={() => navigate("new")}><span>＋</span>新しい回覧</button>
          <button className={page === "cases" || page === "case" ? "active" : ""} onClick={() => navigate("cases")}><span>▤</span>案件一覧{incomingCount > 0 && <b>{incomingCount}</b>}</button>
          <button className={page === "approvers" ? "active" : ""} onClick={() => navigate("approvers")}><span>♙</span>承認者マスタ</button>
        </nav>
        <div className="sidebar-note"><span>DEMO</span><strong>STEP 1</strong><p>データはこのブラウザ内だけに保存されます。</p></div>
      </aside>

      <div className="main-column">
        <header className="topbar">
          <div><span className="demo-pill">ローカルUIモック</span></div>
          <div className="topbar-controls">{page === "new" ? <div className="user-switcher"><span className="user-caption">回覧開始者（デモ）</span><Avatar name={currentUser.name} small /><select value={currentUser.id} onChange={(event) => setCurrentUser(event.target.value)} aria-label="回覧開始者"><option value="" disabled>回覧開始者</option>{data.approvers.filter((person) => person.active).map((person) => <option key={person.id} value={person.id}>{person.name}（{person.department}）</option>)}</select></div> : <div className="address-model"><span>↗</span><div><strong>案件URL方式</strong><small>専用URLをメールで共有</small></div></div>}</div>
        </header>

        <main>
          {page === "home" && <Home data={data} navigate={navigate} onReset={() => { setData(resetData()); notify("デモデータを初期状態に戻しました"); }} />}
          {page === "approvers" && <ApproversPage approvers={data.approvers} onChange={updateApprovers} notify={notify} />}
          {page === "new" && <NewCirculationPage data={data} currentUser={currentUser} setTemplates={updateTemplates} onStart={(created) => { setData((current) => ({ ...current, cases: [created, ...current.cases] })); notify("回覧を開始しました。最初の承認者が処理待ちです"); navigate("case", created.accessKey); }} notify={notify} />}
          {page === "cases" && <CasesPage cases={data.cases} onOpen={(id) => navigate("case", id)} />}
          {page === "case" && selectedCase && <CaseDetail item={selectedCase} approvers={data.approvers} onBack={() => navigate("cases")} onChange={updateCase} notify={notify} />}
          {page === "case" && !selectedCase && <div className="page"><section className="panel"><EmptyState icon="?" title="案件が見つかりません" text="URLが正しいか、データがこのブラウザに保存されているか確認してください。" /><div className="not-found-action"><button className="primary-button" onClick={() => navigate("cases")}>案件一覧へ</button></div></section></div>}
        </main>
      </div>
      {toast && <div className="toast"><span>✓</span>{toast}</div>}
    </div>
  );
}

function PageHeading({ eyebrow, title, description, actions }: { eyebrow: string; title: string; description: string; actions?: React.ReactNode }) {
  return <header className="page-heading"><div><span className="eyebrow">{eyebrow}</span><h1>{title}</h1><p>{description}</p></div>{actions && <div className="page-actions">{actions}</div>}</header>;
}

function Home({ data, navigate, onReset }: { data: AppData; navigate: (page: Page, id?: string) => void; onReset: () => void }) {
  const incoming = data.cases.filter((item) => item.state !== "completed" && (item.currentMemberId || item.returnedToStarter));
  const started = data.cases.filter((item) => item.state !== "completed");
  const completed = data.cases.filter((item) => item.state === "completed");
  const recent = [...data.cases].sort((a, b) => b.updatedAt.localeCompare(a.updatedAt)).slice(0, 4);
  return <div className="page home-page">
    <section className="hero-panel">
      <div><span className="eyebrow light">WORKFLOW OVERVIEW</span><h1>承認状況を、<br />ひと目で。</h1><p>案件URLを開くだけで、現在の確認者と回覧状況がわかります。</p></div>
      <div className="hero-cta"><span>今日の処理待ち</span><strong>{incoming.length}</strong><small>件</small><button onClick={() => navigate("cases")}>処理する →</button></div>
    </section>
    <div className="stats-grid">
      <button className="stat-card accent" onClick={() => navigate("cases")}><span className="stat-icon">◷</span><div><small>承認待ち案件</small><strong>{incoming.length}<em>件</em></strong><p>現在確認者の処理を待っています</p></div><span className="chevron">›</span></button>
      <button className="stat-card" onClick={() => navigate("cases")}><span className="stat-icon blue">↗</span><div><small>回覧中の案件</small><strong>{started.length}<em>件</em></strong><p>社内で回覧中の全案件</p></div><span className="chevron">›</span></button>
      <button className="stat-card" onClick={() => navigate("cases")}><span className="stat-icon green">✓</span><div><small>完了案件</small><strong>{completed.length}<em>件</em></strong><p>回覧が完了した案件</p></div><span className="chevron">›</span></button>
    </div>
    <div className="home-grid">
      <section className="panel">
        <div className="section-title"><div><span className="eyebrow">RECENT ACTIVITY</span><h2>最近の案件</h2></div><button className="text-button" onClick={() => navigate("cases")}>すべて見る →</button></div>
        <div className="case-list compact">
          {recent.map((item) => <button key={item.id} className="case-row" onClick={() => navigate("case", item.id)}><span className={`file-icon ${item.fileName.endsWith(".pdf") ? "pdf" : "excel"}`}>{item.fileName.endsWith(".pdf") ? "PDF" : "XLS"}</span><span className="case-main"><strong>{item.fileName}</strong><small>{item.id} ・ {providerLabels[item.provider]}</small></span><StatusBadge state={item.state} /><span className="case-date">{formatDate(item.updatedAt)}</span><span>›</span></button>)}
        </div>
      </section>
      <aside className="panel quick-panel"><span className="eyebrow">QUICK START</span><h2>新しい回覧を始める</h2><p>対象ファイルと承認ルートを選んで、すぐに回覧を開始できます。</p><div className="route-preview"><span>1</span><i /><span>2</span><i /><span>3</span></div><button className="primary-button full" onClick={() => navigate("new")}>＋ 新しい回覧を作成</button><button className="quiet-button full" onClick={onReset}>デモデータを初期化</button></aside>
    </div>
  </div>;
}

function ApproversPage({ approvers, onChange, notify }: { approvers: Approver[]; onChange: (items: Approver[]) => void; notify: (message: string) => void }) {
  const [query, setQuery] = useState("");
  const [editing, setEditing] = useState<Approver | "new">();
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

  return <div className="page">
    <PageHeading eyebrow="APPROVER DIRECTORY" title="承認者マスタ" description="回付ルートで利用する承認者を管理します。" actions={<button className="primary-button" onClick={() => setEditing("new")}>＋ 承認者を追加</button>} />
    <section className="panel">
      <div className="toolbar"><div className="search-input large"><span>⌕</span><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="氏名・メールアドレス・部門で検索" /></div><span className="result-count">{visible.length}名</span></div>
      <div className="table-wrap"><table><thead><tr><th>氏名</th><th>メールアドレス</th><th>部門</th><th>状態</th><th>更新日時</th><th aria-label="操作" /></tr></thead><tbody>
        {visible.map((person) => <tr key={person.id} className={!person.active ? "muted-row" : ""}><td><div className="person-cell"><Avatar name={person.name} /><strong>{person.name}</strong></div></td><td><a href={`mailto:${person.email}`}>{person.email}</a></td><td>{person.department || "未設定"}</td><td><span className={`active-badge ${person.active ? "is-active" : ""}`}>{person.active ? "有効" : "無効"}</span></td><td>{formatDate(person.updatedAt)}</td><td><div className="row-actions"><button onClick={() => setEditing(person)}>編集</button><button onClick={() => { onChange(approvers.map((item) => item.id === person.id ? { ...item, active: !item.active, updatedAt: nowIso() } : item)); notify(`${person.name}さんを${person.active ? "無効化" : "有効化"}しました`); }}>{person.active ? "無効化" : "有効化"}</button><button className="danger-text" onClick={() => { if (window.confirm(`${person.name}さんを削除しますか？`)) { onChange(approvers.filter((item) => item.id !== person.id)); notify("承認者を削除しました"); } }}>削除</button></div></td></tr>)}
      </tbody></table></div>
      {visible.length === 0 && <EmptyState icon="⌕" title="承認者が見つかりません" text="検索条件を変えてお試しください。" />}
    </section>
    {editing && <ApproverForm person={editing === "new" ? undefined : editing} onClose={() => setEditing(undefined)} onSave={save} />}
  </div>;
}

function ApproverForm({ person, onClose, onSave }: { person?: Approver; onClose: () => void; onSave: (values: { name: string; email: string; department: string; entraUserId: string }) => void }) {
  const [name, setName] = useState(person?.name ?? "");
  const [email, setEmail] = useState(person?.email ?? "");
  const [department, setDepartment] = useState(person?.department ?? "");
  const [entraUserId, setEntraUserId] = useState(person?.entraUserId ?? "");
  const valid = name.trim() && email.includes("@") && department.trim();
  return <Modal title={person ? "承認者を編集" : "承認者を新規登録"} onClose={onClose}><form onSubmit={(event) => { event.preventDefault(); if (valid) onSave({ name: name.trim(), email: email.trim(), department: department.trim(), entraUserId: entraUserId.trim() }); }}><div className="modal-body form-grid"><label><span>氏名 <b>必須</b></span><input value={name} onChange={(event) => setName(event.target.value)} placeholder="例：山田 太郎" autoFocus /></label><label><span>メールアドレス <b>必須</b></span><input type="email" value={email} onChange={(event) => setEmail(event.target.value)} placeholder="taro.yamada@example.com" /></label><label><span>部門 <b>必須</b></span><input value={department} onChange={(event) => setDepartment(event.target.value)} placeholder="例：技術部" /></label><label><span>Entra ID userId <i>任意・STEP 4で同期</i></span><input value={entraUserId} onChange={(event) => setEntraUserId(event.target.value)} placeholder="現在は未接続です" disabled /></label></div><div className="modal-footer"><button type="button" className="quiet-button" onClick={onClose}>キャンセル</button><button type="submit" className="primary-button" disabled={!valid}>{person ? "変更を保存" : "登録する"}</button></div></form></Modal>;
}

function NewCirculationPage({ data, currentUser, setTemplates, onStart, notify }: { data: AppData; currentUser: Approver; setTemplates: (items: RouteTemplate[]) => void; onStart: (item: CirculationCase) => void; notify: (message: string) => void }) {
  const [provider, setProvider] = useState<ProviderType>("sharepoint");
  const [selectedFile, setSelectedFile] = useState<MockFile>();
  const [filePicker, setFilePicker] = useState(false);
  const [selectedTemplateId, setSelectedTemplateId] = useState("");
  const [route, setRoute] = useState<RouteMember[]>([]);
  const [saveModal, setSaveModal] = useState(false);
  const [dragIndex, setDragIndex] = useState<number>();

  const normalize = (items: RouteMember[]) => items.map((item, index) => ({ ...item, sequence: index + 1 }));
  const addApprover = (person: Approver) => {
    if (route.some((item) => item.approverId === person.id)) { notify("同じ承認者は重複して追加できません"); return; }
    setRoute(normalize([...route, { id: uid("route"), approverId: person.id, name: person.name, email: person.email, department: person.department, sequence: route.length + 1, requiresStamp: false, status: "pending" }]));
  };
  const move = (from: number, to: number) => {
    if (to < 0 || to >= route.length) return;
    const changed = [...route]; const [picked] = changed.splice(from, 1); changed.splice(to, 0, picked); setRoute(normalize(changed));
  };
  const loadTemplate = () => {
    const template = data.templates.find((item) => item.id === selectedTemplateId);
    if (!template) { notify("読み込むテンプレートを選択してください"); return; }
    setRoute(template.members.map((member, index) => {
      const latest = data.approvers.find((person) => person.id === member.approverId);
      return { ...member, id: uid("route"), sequence: index + 1, status: "pending", name: latest?.name ?? member.name, email: latest?.email ?? member.email, department: latest?.department ?? member.department };
    }));
    notify(`「${template.name}」を読み込みました`);
  };
  const overwrite = () => {
    const target = data.templates.find((item) => item.id === selectedTemplateId);
    if (!target) return setSaveModal(true);
    setTemplates(data.templates.map((item) => item.id === target.id ? { ...item, updatedAt: nowIso(), members: route.map(({ status: _status, completedAt: _completedAt, ...member }) => member) } : item));
    notify(`「${target.name}」を上書き保存しました`);
  };
  const duplicateTemplate = () => {
    const target = data.templates.find((item) => item.id === selectedTemplateId); if (!target) return;
    const copy = { ...target, id: uid("template"), name: `${target.name}（コピー）`, createdAt: nowIso(), updatedAt: nowIso(), members: target.members.map((item) => ({ ...item, id: uid("template-member") })) };
    setTemplates([...data.templates, copy]); setSelectedTemplateId(copy.id); notify("テンプレートを複製しました");
  };
  const start = () => {
    if (!selectedFile || route.length === 0) return;
    const timestamp = nowIso();
    const members = normalize(route).map((item, index) => ({ ...item, id: uid("case-member"), status: index === 0 ? "current" as const : "pending" as const }));
    const id = `CASE-${new Date().getFullYear()}-${String(Date.now()).slice(-8)}`;
    onStart({ id, accessKey: crypto.randomUUID().replaceAll("-", ""), provider, fileId: selectedFile.id, fileName: selectedFile.name, fileUrl: "#demo-document", initiatorId: currentUser.id, initiatorName: currentUser.name, startedAt: timestamp, updatedAt: timestamp, state: "circulating", currentMemberId: members[0].id, members, history: [{ id: uid("history"), caseId: id, actionUserId: currentUser.id, actionUserName: currentUser.name, action: "回覧開始", previousState: "下書き", newState: `回覧中（${members[0].name}）`, createdAt: timestamp }] });
  };

  return <div className="page">
    <PageHeading eyebrow="NEW CIRCULATION" title="新しい回覧" description="対象ファイルと回付ルートを設定して回覧を開始します。" />
    <div className="stepper"><span className="done"><b>1</b>対象ファイル</span><i /><span className={selectedFile ? "done" : ""}><b>2</b>回付ルート</span><i /><span className={route.length ? "done" : ""}><b>3</b>確認・開始</span></div>
    <div className="builder-grid">
      <div className="builder-main">
        <section className="panel form-section"><div className="number-title"><b>1</b><div><h2>対象ファイル</h2><p>STEP 1では候補ファイルから選択するデモです。</p></div></div><div className="two-fields"><label><span className="field-label">保存場所</span><select value={provider} onChange={(event) => { setProvider(event.target.value as ProviderType); setSelectedFile(undefined); }}><option value="sharepoint">SharePoint</option><option value="onedrive">OneDrive</option><option value="shared-folder">共有フォルダ</option></select></label><div><span className="field-label">対象ファイル</span><button className={`file-select ${selectedFile ? "selected" : ""}`} onClick={() => setFilePicker(true)}>{selectedFile ? <><span className={`file-icon ${selectedFile.type === "PDF" ? "pdf" : "excel"}`}>{selectedFile.type === "PDF" ? "PDF" : "XLS"}</span><span><strong>{selectedFile.name}</strong><small>{selectedFile.location}</small></span><em>変更</em></> : <><span className="upload-mark">＋</span><span><strong>ファイルを選択</strong><small>{providerLabels[provider]} のダミーファイル</small></span></>}</button></div></div></section>
        <section className="panel form-section"><div className="number-title"><b>2</b><div><h2>回付ルート</h2><p>テンプレートを読み込むか、承認者を検索して追加します。</p></div></div><div className="template-bar"><select value={selectedTemplateId} onChange={(event) => setSelectedTemplateId(event.target.value)}><option value="">ルートテンプレートを選択</option>{data.templates.map((item) => <option key={item.id} value={item.id}>{item.name}</option>)}</select><button className="secondary-button" onClick={loadTemplate}>読み込み</button><span className="bar-divider" /><button className="text-button" disabled={!selectedTemplateId} onClick={overwrite}>上書き保存</button><button className="text-button" disabled={!selectedTemplateId} onClick={duplicateTemplate}>複製</button><button className="text-button danger-text" disabled={!selectedTemplateId} onClick={() => { const target = data.templates.find((item) => item.id === selectedTemplateId); if (target && window.confirm(`「${target.name}」を削除しますか？`)) { setTemplates(data.templates.filter((item) => item.id !== target.id)); setSelectedTemplateId(""); notify("テンプレートを削除しました"); } }}>削除</button></div>
          <ApproverAutocomplete approvers={data.approvers} excludedIds={route.map((item) => item.approverId)} onSelect={addApprover} />
          <div className="route-list">
            {route.length === 0 ? <EmptyState icon="♙" title="回付者がまだいません" text="上の検索欄から承認者を追加してください。" /> : route.map((member, index) => <div key={member.id} className="route-card" draggable onDragStart={() => setDragIndex(index)} onDragOver={(event) => event.preventDefault()} onDrop={() => { if (dragIndex !== undefined) move(dragIndex, index); setDragIndex(undefined); }}><span className="drag-handle" title="ドラッグして並べ替え">⠿</span><span className="sequence">{index + 1}</span><Avatar name={member.name} /><div className="route-person"><strong>{member.name}</strong><small>{member.department} ・ {member.email}</small></div><label className="stamp-toggle"><input type="checkbox" checked={member.requiresStamp} onChange={() => setRoute(route.map((item) => item.id === member.id ? { ...item, requiresStamp: !item.requiresStamp } : item))} /><span>{member.requiresStamp ? "捺印必要" : "確認のみ"}</span></label><div className="move-actions"><button disabled={index === 0} onClick={() => move(index, index - 1)} aria-label={`${member.name}を上へ`}>↑</button><button disabled={index === route.length - 1} onClick={() => move(index, index + 1)} aria-label={`${member.name}を下へ`}>↓</button><button className="remove" onClick={() => setRoute(normalize(route.filter((item) => item.id !== member.id)))} aria-label={`${member.name}を削除`}>×</button></div></div>)}
          </div>
          {route.length > 0 && <div className="route-footer"><span>{route.length}名の回付ルート</span><button className="secondary-button" onClick={() => setSaveModal(true)}>ルートを名前を付けて保存</button></div>}
        </section>
      </div>
      <aside className="panel summary-card"><span className="eyebrow">SUMMARY</span><h2>回覧内容の確認</h2><dl><div><dt>対象ファイル</dt><dd>{selectedFile?.name ?? "未選択"}</dd></div><div><dt>保存場所</dt><dd>{providerLabels[provider]}</dd></div><div><dt>回覧開始者</dt><dd>{currentUser.name}</dd></div><div><dt>承認者</dt><dd>{route.length}名</dd></div></dl><div className="summary-route">{route.map((member, index) => <div key={member.id}><span>{index + 1}</span><p><strong>{member.name}</strong><small>{member.requiresStamp ? "捺印必要" : "確認のみ"}</small></p></div>)}</div><button className="primary-button full large-button" disabled={!selectedFile || route.length === 0} onClick={start}>回覧を開始する →</button>{(!selectedFile || route.length === 0) && <small className="helper">ファイルと1名以上の承認者を設定してください</small>}</aside>
    </div>
    {filePicker && <FilePicker provider={provider} selected={selectedFile?.id} onClose={() => setFilePicker(false)} onSelect={(file) => { setSelectedFile(file); setFilePicker(false); notify(`${file.name} を選択しました`); }} />}
    {saveModal && <TemplateSaveModal currentUser={currentUser} route={route} onClose={() => setSaveModal(false)} onSave={(template) => { setTemplates([...data.templates, template]); setSelectedTemplateId(template.id); setSaveModal(false); notify(`「${template.name}」を保存しました`); }} />}
  </div>;
}

function FilePicker({ provider, selected, onClose, onSelect }: { provider: ProviderType; selected?: string; onClose: () => void; onSelect: (file: MockFile) => void }) {
  const files = mockFiles.filter((item) => item.provider === provider);
  return <Modal title={`${providerLabels[provider]} からファイルを選択`} onClose={onClose} wide><div className="modal-body"><div className="demo-banner"><span>DEMO</span>実ファイルには接続していません。STEP 5以降でGraph APIへ置き換えます。</div><div className="file-grid">{files.map((file) => <button key={file.id} className={`file-choice ${selected === file.id ? "selected" : ""}`} onClick={() => onSelect(file)}><span className={`file-icon big ${file.type === "PDF" ? "pdf" : "excel"}`}>{file.type === "PDF" ? "PDF" : file.type === "Excel" ? "XLS" : "PPT"}</span><span><strong>{file.name}</strong><small>{file.location}</small><small>更新：{formatDate(file.updatedAt)} ・ {file.updatedBy}</small></span><em>{selected === file.id ? "選択中" : "選択"}</em></button>)}</div></div><div className="modal-footer"><button className="quiet-button" onClick={onClose}>キャンセル</button></div></Modal>;
}

function TemplateSaveModal({ currentUser, route, onClose, onSave }: { currentUser: Approver; route: RouteMember[]; onClose: () => void; onSave: (template: RouteTemplate) => void }) {
  const [name, setName] = useState(""); const [description, setDescription] = useState("");
  return <Modal title="ルートを名前を付けて保存" onClose={onClose}><form onSubmit={(event) => { event.preventDefault(); if (!name.trim()) return; const timestamp = nowIso(); onSave({ id: uid("template"), name: name.trim(), description: description.trim(), createdBy: currentUser.name, createdAt: timestamp, updatedAt: timestamp, members: route.map(({ status: _status, completedAt: _completedAt, ...member }) => ({ ...member, id: uid("template-member") })) }); }}><div className="modal-body form-grid"><label><span>テンプレート名 <b>必須</b></span><input value={name} onChange={(event) => setName(event.target.value)} placeholder="例：部長承認ルート" autoFocus /></label><label><span>説明</span><textarea value={description} onChange={(event) => setDescription(event.target.value)} placeholder="このルートを使う場面を入力" rows={3} /></label><div className="save-preview"><strong>{route.length}名を保存します</strong><span>{route.map((item) => item.name).join(" → ")}</span></div></div><div className="modal-footer"><button type="button" className="quiet-button" onClick={onClose}>キャンセル</button><button type="submit" className="primary-button" disabled={!name.trim()}>保存する</button></div></form></Modal>;
}

function CasesPage({ cases, onOpen }: { cases: CirculationCase[]; onOpen: (accessKey: string) => void }) {
  const [tab, setTab] = useState<CaseTab>("waiting");
  const [query, setQuery] = useState("");
  const groups = useMemo(() => ({
    waiting: cases.filter((item) => item.state !== "completed" && (item.currentMemberId || item.returnedToStarter)),
    all: cases,
    returned: cases.filter((item) => item.state === "returned"),
    completed: cases.filter((item) => item.state === "completed"),
  }), [cases]);
  const visible = groups[tab].filter((item) => item.fileName.toLowerCase().includes(query.toLowerCase()) || item.id.toLowerCase().includes(query.toLowerCase()));
  return <div className="page"><PageHeading eyebrow="CIRCULATION CASES" title="案件一覧" description="案件を開くと、現在の確認者と専用URLを確認できます。" />
    <div className="tabs">{(["waiting", "all", "returned", "completed"] as CaseTab[]).map((key) => <button key={key} className={tab === key ? "active" : ""} onClick={() => setTab(key)}>{({ waiting: "承認待ち", all: "すべて", returned: "差し戻し", completed: "完了" } as Record<CaseTab, string>)[key]}<b>{groups[key].length}</b></button>)}</div>
    <section className="panel"><div className="toolbar"><div className="search-input large"><span>⌕</span><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="ファイル名・案件IDで検索" /></div><span className="result-count">{visible.length}件</span></div>
      {visible.length ? <div className="case-cards">{visible.map((item) => { const current = item.members.find((member) => member.id === item.currentMemberId); const done = item.members.filter((member) => member.status === "approved" || member.status === "stamped").length; return <button key={item.id} className="case-card" onClick={() => onOpen(item.id)}><span className={`file-icon big ${item.fileName.endsWith(".pdf") ? "pdf" : "excel"}`}>{item.fileName.endsWith(".pdf") ? "PDF" : "XLS"}</span><div className="case-card-main"><div><span className="case-id">{item.id}</span><StatusBadge state={item.state} /></div><h3>{item.fileName}</h3><p>{providerLabels[item.provider]} ・ 開始者 {item.initiatorName} ・ {formatDate(item.startedAt)}</p><div className="mini-progress"><span><i style={{ width: `${item.state === "completed" ? 100 : Math.round(done / item.members.length * 100)}%` }} /></span><small>{done} / {item.members.length} 完了</small></div></div><div className="current-owner"><small>{item.state === "completed" ? "完了日時" : "現在の担当"}</small><strong>{item.state === "completed" ? formatDate(item.updatedAt) : current?.name ?? "回覧開始者"}</strong></div><span className="open-arrow">›</span></button>; })}</div> : <EmptyState icon="▤" title="該当する案件はありません" text="別のタブまたは検索条件をお試しください。" />}
    </section>
  </div>;
}

function CaseDetail({ item, approvers, onBack, onChange, notify }: { item: CirculationCase; approvers: Approver[]; onBack: () => void; onChange: (item: CirculationCase) => void; notify: (message: string) => void }) {
  const [rejecting, setRejecting] = useState(false);
  const [editingRoute, setEditingRoute] = useState(false);
  const [emailDraft, setEmailDraft] = useState<EmailDraft>();
  const current = item.members.find((member) => member.id === item.currentMemberId)!;
  const canProcess = item.state !== "completed" && Boolean(current);
  const canRestart = item.state === "returned" && Boolean(item.returnedToStarter);
  const canEdit = item.state !== "completed";
  const reviewOwner = current ?? (item.returnedToStarter ? approvers.find((person) => person.id === item.initiatorId) : undefined);

  const copyCaseUrl = () => {
    navigator.clipboard.writeText(window.location.href)
      .then(() => notify("この案件の専用URLをコピーしました"))
      .catch(() => notify("ブラウザのアドレス欄からURLをコピーしてください"));
  };
  const openMailDraft = (to: string, subject: string, body: string) => {
    const recipient = approvers.find((person) => person.email === to) ?? item.members.find((member) => member.email === to);
    setEmailDraft({ to, recipientName: recipient?.name ?? to, subject, body, purpose: subject.includes("差し戻し") ? "差し戻し通知" : "回覧通知" });
  };
  const createEmail = () => {
    if (!reviewOwner || item.state === "completed") return;
    setEmailDraft({ to: reviewOwner.email, recipientName: reviewOwner.name, purpose: "承認依頼", subject: `【承認依頼】${item.fileName}`, body: `${reviewOwner.name}さん\n\n以下の案件が承認待ちです。\n案件名：${item.fileName}\n案件ID：${item.id}\n現在の確認者：${reviewOwner.name}\n対応内容：${current?.requiresStamp ? "捺印後に回付" : "内容確認後に回付"}\n\n案件URL：\n${window.location.href}\n\n内容を確認し、案件ページから回付してください。` });
  };

  const approve = () => {
    if (!current) return;
    const index = item.members.findIndex((member) => member.id === current.id);
    const timestamp = nowIso(); const next = item.members[index + 1];
    const members = item.members.map((member, memberIndex) => memberIndex === index ? { ...member, status: current.requiresStamp ? "stamped" as const : "approved" as const, completedAt: timestamp } : memberIndex === index + 1 ? { ...member, status: "current" as const } : member);
    const action = current.requiresStamp ? "捺印済み・回付" : "確認OK・回付";
    onChange({ ...item, members, state: next ? "circulating" : "completed", currentMemberId: next?.id, returnedToStarter: false, updatedAt: timestamp, history: [...item.history, { id: uid("history"), caseId: item.id, actionUserId: current.approverId, actionUserName: current.name, action, previousState: `${current.name}が処理中`, newState: next ? `${next.name}が処理中` : "回覧完了", createdAt: timestamp }] });
    notify(next ? `${next.name}さんへ回付しました` : "すべての回覧が完了しました");
    if (next) {
      setEmailDraft({ to: next.email, recipientName: next.name, purpose: "次の確認者への回付通知", subject: `【承認依頼】${item.fileName}`, body: `${next.name}さん\n\n${current.name}さんの処理が完了し、あなたの承認待ちになりました。\n\n案件名：${item.fileName}\n案件ID：${item.id}\n現在の確認者：${next.name}\n対応内容：${next.requiresStamp ? "文書へ捺印後、「捺印済み・回付」を押してください" : "内容確認後、「確認OK・回付」を押してください"}\n\n案件URL：\n${window.location.href}\n\nよろしくお願いいたします。` });
    } else {
      const initiator = approvers.find((person) => person.id === item.initiatorId);
      if (initiator) setEmailDraft({ to: initiator.email, recipientName: initiator.name, purpose: "回覧完了通知", subject: `【回覧完了】${item.fileName}`, body: `${initiator.name}さん\n\n以下の案件の承認回覧が完了しました。\n\n案件名：${item.fileName}\n案件ID：${item.id}\n最終確認者：${current.name}\n完了日時：${formatDate(timestamp)}\n\n案件URL：\n${window.location.href}\n\n回覧履歴は案件ページから確認できます。` });
    }
  };
  const restart = () => {
    const first = item.members[0]; const timestamp = nowIso();
    onChange({ ...item, state: "circulating", returnedToStarter: false, currentMemberId: first.id, updatedAt: timestamp, members: item.members.map((member, index) => ({ ...member, status: index === 0 ? "current" : "pending", completedAt: undefined })), history: [...item.history, { id: uid("history"), caseId: item.id, actionUserId: item.initiatorId, actionUserName: item.initiatorName, action: "再回覧", previousState: "回覧開始者へ差し戻し", newState: `${first.name}が処理中`, createdAt: timestamp }] });
    notify(`${first.name}さんから再回覧を開始しました`);
  };
  const doneCount = item.members.filter((member) => member.status === "approved" || member.status === "stamped").length;
  return <div className="page case-detail-page">
    <button className="back-button" onClick={onBack}>← 案件一覧へ戻る</button>
    <PageHeading eyebrow={item.id} title={item.fileName} description={`${providerLabels[item.provider]} ・ ${item.initiatorName}さんが ${formatDate(item.startedAt)} に開始`} actions={<><button className="secondary-button" onClick={copyCaseUrl}>URLをコピー</button>{item.state !== "completed" && <button className="primary-button" onClick={createEmail}>現在の確認者へメール</button>}<StatusBadge state={item.state} /></>} />
    <section className={`current-reviewer-banner ${item.state === "completed" ? "is-completed" : ""}`}>
      <div className="reviewer-symbol">{item.state === "completed" ? "✓" : "確"}</div>
      <div className="reviewer-copy"><span className="eyebrow">CURRENT REVIEWER</span><strong>{item.state === "completed" ? "現在の確認者：なし（回覧完了）" : reviewOwner ? `${reviewOwner.name}さんの承認待ち` : "現在の対応者"}</strong><p>{item.state === "completed" ? "この案件の承認回覧は完了しています。" : "案件URLを現在の確認者へメールで共有してください。"}</p></div>
      {reviewOwner && item.state !== "completed" && <div className="reviewer-person"><Avatar name={reviewOwner.name} /><div><strong>{reviewOwner.name}</strong><small>{reviewOwner.department} ・ {reviewOwner.email}</small></div></div>}
      {item.state !== "completed" && <span className="identity-badge open-access">承認待ち</span>}
    </section>
    <div className="detail-grid">
      <div className="detail-main">
        {canProcess && <section className={`action-panel ${current.requiresStamp ? "stamp" : "review"}`}><div className="action-heading"><span className="action-icon">{current.requiresStamp ? "印" : "✓"}</span><div><span className="eyebrow">YOUR ACTION</span><h2>{current.requiresStamp ? "捺印が必要です" : "内容を確認してください"}</h2><p>{current.requiresStamp ? "文書の捺印欄へ捺印し、保存後に回付してください。" : "文書の内容に問題がなければ、次の承認者へ回付してください。"}</p></div></div><div className="action-buttons"><button className="document-button" onClick={() => notify("デモ：実ファイルを開く動作はSTEP 7で接続します")}>↗ 文書を開く</button><span className="mail-auto-note">処理後、次の宛先へのメール下書きが開きます</span><span className="action-spacer" /><button className="danger-button" onClick={() => setRejecting(true)}>NG・差し戻し</button><button className="primary-button" onClick={approve}>{current.requiresStamp ? "捺印済み・回付" : "確認OK・回付"} →</button></div></section>}
        {canRestart && <section className="action-panel returned-action"><div className="action-heading"><span className="action-icon">↩</span><div><span className="eyebrow">RETURNED</span><h2>この案件は開始者へ差し戻されています</h2><p>コメントと文書を確認し、先頭の承認者から再回覧できます。</p></div></div><div className="action-buttons"><button className="document-button" onClick={() => notify("デモ：文書を開きました")}>↗ 文書を開く</button><span className="action-spacer" /><button className="primary-button" onClick={restart}>先頭から再回覧 →</button></div></section>}
        <section className="panel progress-panel"><div className="section-title"><div><span className="eyebrow">ROUTE PROGRESS</span><h2>回付ルートと進捗</h2></div>{canEdit && <button className="secondary-button" onClick={() => setEditingRoute(true)}>未処理ルートを編集</button>}</div>
          <div className="progress-summary"><strong>{item.state === "completed" ? "100" : Math.round(doneCount / item.members.length * 100)}<small>%</small></strong><span><i style={{ width: `${item.state === "completed" ? 100 : Math.round(doneCount / item.members.length * 100)}%` }} /></span><p>{doneCount}/{item.members.length}名 完了</p></div>
          <div className="timeline">{item.members.map((member, index) => { const completed = member.status === "approved" || member.status === "stamped"; return <div key={member.id} className={`timeline-item ${member.status}`}><div className="timeline-marker"><span>{completed ? "✓" : member.status === "current" ? "●" : index + 1}</span>{index < item.members.length - 1 && <i />}</div><Avatar name={member.name} /><div className="timeline-person"><strong>{member.name}</strong><small>{member.department} ・ {member.email}</small></div><span className={`route-kind ${member.requiresStamp ? "stamp-kind" : ""}`}>{member.requiresStamp ? "捺印必要" : "確認のみ"}</span><div className="timeline-state"><strong>{member.status === "stamped" ? "捺印済み" : member.status === "approved" ? "確認済み" : member.status === "current" ? "処理中" : "未処理"}</strong>{member.completedAt && <small>{formatDate(member.completedAt)}</small>}</div></div>; })}</div>
        </section>
        <section className="panel history-panel"><div className="section-title"><div><span className="eyebrow">AUDIT TRAIL</span><h2>回覧履歴</h2></div><span className="result-count">{item.history.length}件</span></div><div className="history-list">{[...item.history].reverse().map((entry) => <article key={entry.id}><span className={`history-mark ${entry.action.includes("NG") ? "ng" : ""}`}>{entry.action.includes("NG") ? "!" : "✓"}</span><div><div><strong>{entry.actionUserName}</strong><span>{entry.action}</span><time>{formatDate(entry.createdAt)}</time></div><p>{entry.previousState} → {entry.newState}</p>{entry.returnToUserName && <p className="return-line">戻り先：{entry.returnToUserName}</p>}{entry.comment && <blockquote>{entry.comment}</blockquote>}</div></article>)}</div></section>
      </div>
      <aside className="detail-aside"><section className="panel metadata-card"><span className={`file-icon huge ${item.fileName.endsWith(".pdf") ? "pdf" : "excel"}`}>{item.fileName.endsWith(".pdf") ? "PDF" : "XLS"}</span><h3>{item.fileName}</h3><button className="secondary-button full" onClick={() => notify("デモ：文書を開く動作はSTEP 7で接続します")}>↗ 文書を開く</button><dl><div><dt>保存場所</dt><dd>{providerLabels[item.provider]}</dd></div><div><dt>回覧開始者</dt><dd>{item.initiatorName}</dd></div><div><dt>開始日時</dt><dd>{formatDate(item.startedAt)}</dd></div><div><dt>現在の担当者</dt><dd>{current?.name ?? (item.state === "completed" ? "—" : item.initiatorName)}</dd></div><div><dt>案件状態</dt><dd>{stateLabel(item.state)}</dd></div></dl></section><div className="security-note"><span>⌾</span><div><strong>印影データは保存しません</strong><p>このアプリは印影・署名画像を保持しない設計です。</p></div></div></aside>
    </div>
    {rejecting && current && <RejectModal item={item} current={current} user={{ id: current.approverId, name: current.name }} onClose={() => setRejecting(false)} onSubmit={(changed) => { onChange(changed); setRejecting(false); notify("案件を差し戻しました"); const destination = changed.returnedToStarter ? approvers.find((person) => person.id === changed.initiatorId) : changed.members.find((member) => member.id === changed.currentMemberId); const latest = changed.history.at(-1); if (destination) openMailDraft(destination.email, `【差し戻し】${changed.fileName}`, `${destination.name}さん\n\n${current.name}さんから案件が差し戻されました。\n\n案件名：${changed.fileName}\n案件ID：${changed.id}\n戻り先：${destination.name}\n差し戻し理由：${latest?.comment ?? ""}\n\n案件URL：\n${window.location.href}\n\n内容を確認し、案件ページから対応してください。`); }} />}
    {editingRoute && <RouteEditModal item={item} approvers={approvers} onClose={() => setEditingRoute(false)} onSave={(changed) => { onChange(changed); setEditingRoute(false); notify("未処理の回付ルートを更新しました"); }} notify={notify} />}
    {emailDraft && <EmailComposerModal draft={emailDraft} onClose={() => setEmailDraft(undefined)} onSend={(message) => { setEmailDraft(undefined); notify(`${message.recipientName}さんへメールを送信しました（STEP 1デモ）`); }} />}
  </div>;
}

function RejectModal({ item, current, user, onClose, onSubmit }: { item: CirculationCase; current: RouteMember; user: { id: string; name: string }; onClose: () => void; onSubmit: (changed: CirculationCase) => void }) {
  const currentIndex = item.members.findIndex((member) => member.id === current.id);
  const prior = item.members.slice(0, currentIndex).filter((member) => member.status === "approved" || member.status === "stamped");
  const [target, setTarget] = useState("starter"); const [comment, setComment] = useState("");
  const submit = () => {
    if (!comment.trim()) return;
    const timestamp = nowIso(); const isStarter = target === "starter"; const returnMember = item.members.find((member) => member.approverId === target);
    const members = isStarter ? item.members : item.members.map((member, index) => {
      const returnIndex = item.members.findIndex((candidate) => candidate.approverId === target);
      if (index < returnIndex) return member;
      return { ...member, status: index === returnIndex ? "current" as const : "pending" as const, completedAt: undefined };
    });
    onSubmit({ ...item, state: "returned", returnedToStarter: isStarter, currentMemberId: isStarter ? undefined : returnMember?.id, members, updatedAt: timestamp, history: [...item.history, { id: uid("history"), caseId: item.id, actionUserId: user.id, actionUserName: user.name, action: "NG・差し戻し", comment: comment.trim(), returnToUserId: isStarter ? item.initiatorId : returnMember?.approverId, returnToUserName: isStarter ? item.initiatorName : returnMember?.name, previousState: `${current.name}が処理中`, newState: `${isStarter ? item.initiatorName : returnMember?.name}へ差し戻し`, createdAt: timestamp }] });
  };
  return <Modal title="NG・差し戻し" onClose={onClose}><div className="modal-body"><div className="warning-box"><span>!</span><p>差し戻し先を選び、理由を入力してください。処理済みの後続メンバーは未処理に戻ります。</p></div><fieldset className="return-options"><legend>戻り先</legend><label><input type="radio" name="target" value="starter" checked={target === "starter"} onChange={(event) => setTarget(event.target.value)} /><Avatar name={item.initiatorName} small /><span><strong>{item.initiatorName}</strong><small>回覧開始者</small></span></label>{prior.map((member) => <label key={member.id}><input type="radio" name="target" value={member.approverId} checked={target === member.approverId} onChange={(event) => setTarget(event.target.value)} /><Avatar name={member.name} small /><span><strong>{member.name}</strong><small>{member.department} ・ 処理済み</small></span></label>)}</fieldset><label className="comment-field"><span>コメント <b>必須</b></span><textarea value={comment} onChange={(event) => setComment(event.target.value)} placeholder="差し戻し理由を具体的に入力してください" rows={4} autoFocus /><small>{comment.length} / 500</small></label></div><div className="modal-footer"><button className="quiet-button" onClick={onClose}>キャンセル</button><button className="danger-button solid" disabled={!comment.trim()} onClick={submit}>差し戻す</button></div></Modal>;
}

function RouteEditModal({ item, approvers, onClose, onSave, notify }: { item: CirculationCase; approvers: Approver[]; onClose: () => void; onSave: (changed: CirculationCase) => void; notify: (message: string) => void }) {
  const [members, setMembers] = useState<RouteMember[]>(structuredClone(item.members));
  const currentIndex = members.findIndex((member) => member.id === item.currentMemberId);
  const normalize = (items: RouteMember[]) => items.map((member, index) => ({ ...member, sequence: index + 1 }));
  const add = (person: Approver) => setMembers(normalize([...members, { id: uid("case-member"), approverId: person.id, name: person.name, email: person.email, department: person.department, sequence: members.length + 1, requiresStamp: false, status: "pending" }]));
  const move = (index: number, direction: -1 | 1) => {
    const target = index + direction;
    if (target <= currentIndex || target >= members.length || members[target].status !== "pending") return;
    const changed = [...members]; [changed[index], changed[target]] = [changed[target], changed[index]]; setMembers(normalize(changed));
  };
  const changedSummary = JSON.stringify(members.map(({ id: _id, ...member }) => member)) !== JSON.stringify(item.members.map(({ id: _id, ...member }) => member));
  return <Modal title="未処理の回付ルートを編集" onClose={onClose} wide><div className="modal-body"><div className="demo-banner neutral"><span>編集範囲</span>処理済み承認者は履歴として固定されます。現在処理中の承認者は削除できません。</div><ApproverAutocomplete approvers={approvers} excludedIds={members.map((member) => member.approverId)} onSelect={add} label="未処理の末尾へ承認者を追加" /><div className="route-list edit-route-list">{members.map((member, index) => { const editable = member.status === "pending"; return <div key={member.id} className={`route-card ${!editable ? "locked" : ""}`}><span className="sequence">{index + 1}</span><Avatar name={member.name} /><div className="route-person"><strong>{member.name}</strong><small>{member.department} ・ {member.status === "pending" ? "未処理" : member.status === "current" ? "処理中（固定）" : "処理済み（固定）"}</small></div><label className="stamp-toggle"><input type="checkbox" checked={member.requiresStamp} disabled={!editable} onChange={() => setMembers(members.map((candidate) => candidate.id === member.id ? { ...candidate, requiresStamp: !candidate.requiresStamp } : candidate))} /><span>{member.requiresStamp ? "捺印必要" : "確認のみ"}</span></label><div className="move-actions"><button disabled={!editable || index - 1 <= currentIndex || members[index - 1]?.status !== "pending"} onClick={() => move(index, -1)}>↑</button><button disabled={!editable || index === members.length - 1 || members[index + 1]?.status !== "pending"} onClick={() => move(index, 1)}>↓</button><button className="remove" disabled={!editable} title={!editable ? "処理済み・処理中の承認者は削除できません" : "削除"} onClick={() => { if (!editable) { notify("現在処理中または処理済みの承認者は削除できません"); return; } setMembers(normalize(members.filter((candidate) => candidate.id !== member.id))); }}>×</button></div></div>; })}</div></div><div className="modal-footer"><button className="quiet-button" onClick={onClose}>キャンセル</button><button className="primary-button" disabled={!changedSummary} onClick={() => { const timestamp = nowIso(); onSave({ ...item, members, updatedAt: timestamp, history: [...item.history, { id: uid("history"), caseId: item.id, actionUserId: item.initiatorId, actionUserName: item.initiatorName, action: "回付ルート変更", comment: "未処理部分の追加・削除・順番・捺印要否を更新", previousState: `${item.members.length}名`, newState: `${members.length}名`, createdAt: timestamp }] }); }}>変更を保存</button></div></Modal>;
}
