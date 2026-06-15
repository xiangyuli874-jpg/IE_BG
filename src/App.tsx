import {
  ClipboardList,
  Copy,
  Download,
  GripVertical,
  LogOut,
  Loader2,
  Plus,
  RefreshCw,
  ShieldCheck,
  Trash2
} from "lucide-react";
import type { FormEvent, PointerEvent as ReactPointerEvent } from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { DEFAULT_PARAMETERS, LINE_CONFIGS, findLine } from "./data/presets";
import type { SampleValue, WorktimeGroup, WorktimeProcess, WorktimeSession } from "./types";

const RECENT_KEY = "gongshibiao.recentSessions";
const ACCESS_CODE_KEY = "gongshibiao.accessCode";
const ACCESS_CODE_HEADER = "X-Access-Code";
const SAVE_DELAY = 650;
const DEFAULT_TEMPLATE_MODEL_ID = "ordinary-washer-dryer";
const FALLBACK_MODEL_NAME = "洗烘结构";

type SaveState = "idle" | "saving" | "saved" | "error";

interface RecentSession {
  id: string;
  label: string;
  updatedAt: string;
}

function createProcess(order: number): WorktimeProcess {
  const id = `proc_${crypto.randomUUID().replace(/-/g, "").slice(0, 12)}`;
  return {
    id,
    order,
    stationNo: null,
    sourceKey: "",
    name: "",
    samples: [null, null, null, null, null],
    people: 1,
    remark: ""
  };
}

function readRecent(): RecentSession[] {
  try {
    const parsed = JSON.parse(localStorage.getItem(RECENT_KEY) || "[]") as RecentSession[];
    return Array.isArray(parsed) ? parsed.slice(0, 8) : [];
  } catch {
    return [];
  }
}

function sessionLabel(session: WorktimeSession) {
  const line = findLine(session.lineId);
  return `${line?.name ?? session.lineId} / ${displayModelName(session)}`;
}

function displayModelName(session: WorktimeSession) {
  return session.modelName?.trim() || FALLBACK_MODEL_NAME;
}

function rememberSession(session: WorktimeSession) {
  const recent = readRecent().filter((item) => item.id !== session.id);
  const next = [{ id: session.id, label: sessionLabel(session), updatedAt: session.updatedAt }, ...recent].slice(0, 8);
  localStorage.setItem(RECENT_KEY, JSON.stringify(next));
}

function forgetRecentSession(id: string) {
  const next = readRecent().filter((item) => item.id !== id);
  localStorage.setItem(RECENT_KEY, JSON.stringify(next));
  return next;
}

function numberOrNull(value: string): SampleValue {
  if (value.trim() === "") return null;
  const num = Number(value);
  return Number.isFinite(num) ? num : null;
}

function formatInput(value: number | null | undefined) {
  return value === null || value === undefined ? "" : String(value);
}

async function apiJson<T>(url: string, accessCode: string, init?: RequestInit): Promise<T> {
  const response = await fetch(url, {
    ...init,
    headers: {
      "Content-Type": "application/json",
      [ACCESS_CODE_HEADER]: accessCode,
      ...init?.headers
    }
  });
  const data = (await response.json().catch(() => null)) as T & { error?: string };
  if (!response.ok) throw new Error(data?.error || "请求失败");
  return data;
}

function useSessionFromUrl() {
  const params = new URLSearchParams(window.location.search);
  return params.get("session");
}

export default function App() {
  const [session, setSession] = useState<WorktimeSession | null>(null);
  const [accessCode, setAccessCode] = useState(() => localStorage.getItem(ACCESS_CODE_KEY) || "");
  const [accessInput, setAccessInput] = useState("");
  const [checkingAccess, setCheckingAccess] = useState(false);
  const [lineId, setLineId] = useState("a-line");
  const [modelName, setModelName] = useState("");
  const [loading, setLoading] = useState(true);
  const [creating, setCreating] = useState(false);
  const [exporting, setExporting] = useState<"group" | "all" | null>(null);
  const [saveState, setSaveState] = useState<SaveState>("idle");
  const [message, setMessage] = useState("");
  const [recent, setRecent] = useState<RecentSession[]>([]);
  const [dragProcessId, setDragProcessId] = useState<string | null>(null);
  const saveTimer = useRef<number | null>(null);
  const hydrated = useRef(false);

  const enabledLines = LINE_CONFIGS.filter((line) => line.enabled);

  const currentGroup = useMemo(() => {
    if (!session) return null;
    return session.groups.find((group) => group.id === session.currentGroupId) ?? session.groups[0] ?? null;
  }, [session]);

  const updateUrl = useCallback((id: string | null, mode: "push" | "replace" = "replace") => {
    const url = id ? `${window.location.pathname}?session=${encodeURIComponent(id)}` : window.location.pathname;
    if (`${window.location.pathname}${window.location.search}` === url) return;
    if (mode === "push") {
      window.history.pushState(null, "", url);
    } else {
      window.history.replaceState(null, "", url);
    }
  }, []);

  const requireAccessAgain = useCallback((notice = "访问码错误或已失效，请重新输入") => {
    localStorage.removeItem(ACCESS_CODE_KEY);
    setAccessCode("");
    setAccessInput("");
    setSession(null);
    setMessage(notice);
  }, []);

  const handleAccessSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    const code = accessInput.trim();
    if (!code) {
      setMessage("请输入访问码");
      return;
    }
    setCheckingAccess(true);
    setMessage("");
    try {
      await apiJson<{ ok: true }>("/api/access", code, { method: "POST", body: "{}" });
      localStorage.setItem(ACCESS_CODE_KEY, code);
      setAccessCode(code);
      setAccessInput("");
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "访问码校验失败");
    } finally {
      setCheckingAccess(false);
    }
  };

  const logoutAccess = () => {
    localStorage.removeItem(ACCESS_CODE_KEY);
    setAccessCode("");
    setAccessInput("");
    setSession(null);
    updateUrl(null);
    setMessage("已退出，请重新输入访问码");
  };

  const loadSession = useCallback(async (id: string, historyMode: "push" | "replace" | "none" = "replace") => {
    if (!accessCode) return;
    setLoading(true);
    setMessage("");
    try {
      const data = await apiJson<{ session: WorktimeSession }>(`/api/sessions/${id}`, accessCode);
      setSession(data.session);
      setLineId(data.session.lineId);
      setModelName(data.session.modelName ?? "");
      rememberSession(data.session);
      setRecent(readRecent());
      if (historyMode !== "none") updateUrl(data.session.id, historyMode);
    } catch (error) {
      const text = error instanceof Error ? error.message : "测量单读取失败";
      if (text.includes("访问码")) {
        requireAccessAgain(text);
      } else {
        setMessage(text);
      }
      setSession(null);
      if (historyMode !== "none") updateUrl(null, "replace");
    } finally {
      setLoading(false);
      hydrated.current = true;
    }
  }, [accessCode, requireAccessAgain, updateUrl]);

  useEffect(() => {
    const handlePopState = () => {
      const id = useSessionFromUrl();
      if (id) {
        void loadSession(id, "none");
      } else {
        setSession(null);
        setLoading(false);
        setMessage("");
      }
    };
    window.addEventListener("popstate", handlePopState);
    return () => window.removeEventListener("popstate", handlePopState);
  }, [loadSession]);

  useEffect(() => {
    if (!accessCode) {
      setLoading(false);
      hydrated.current = true;
      return;
    }
    setRecent(readRecent());
    const id = useSessionFromUrl();
    if (id) {
      void loadSession(id, "replace");
    } else {
      setLoading(false);
      hydrated.current = true;
    }
  }, [accessCode, loadSession]);

  const createNewSession = async () => {
    setCreating(true);
    setMessage("");
    try {
      const data = await apiJson<{ session: WorktimeSession }>("/api/sessions", accessCode, {
        method: "POST",
        body: JSON.stringify({ lineId, modelId: DEFAULT_TEMPLATE_MODEL_ID, modelName: modelName.trim() })
      });
      setSession(data.session);
      setModelName(data.session.modelName ?? "");
      rememberSession(data.session);
      setRecent(readRecent());
      updateUrl(data.session.id, "push");
    } catch (error) {
      const text = error instanceof Error ? error.message : "创建失败";
      if (text.includes("访问码")) requireAccessAgain(text);
      else setMessage(text);
    } finally {
      setCreating(false);
    }
  };

  const persist = useCallback(async (nextSession: WorktimeSession) => {
    setSaveState("saving");
    try {
      const data = await apiJson<{ session: WorktimeSession }>(`/api/sessions/${nextSession.id}`, accessCode, {
        method: "PUT",
        body: JSON.stringify(nextSession)
      });
      rememberSession(data.session);
      setRecent(readRecent());
      setSaveState("saved");
    } catch (error) {
      const text = error instanceof Error ? error.message : "";
      if (text.includes("访问码")) requireAccessAgain(text);
      setSaveState("error");
    }
  }, [accessCode, requireAccessAgain]);

  const patchSession = (updater: (draft: WorktimeSession) => WorktimeSession) => {
    setSession((current) => {
      if (!current) return current;
      const next = updater(structuredClone(current));
      next.updatedAt = new Date().toISOString();
      return next;
    });
  };

  useEffect(() => {
    if (!hydrated.current || !session) return;
    if (saveTimer.current) window.clearTimeout(saveTimer.current);
    setSaveState("idle");
    saveTimer.current = window.setTimeout(() => {
      void persist(session);
    }, SAVE_DELAY);
    return () => {
      if (saveTimer.current) window.clearTimeout(saveTimer.current);
    };
  }, [session, persist]);

  const updateParameter = (key: keyof WorktimeSession["parameters"], value: number | null) => {
    patchSession((draft) => {
      draft.parameters = {
        ...DEFAULT_PARAMETERS,
        ...draft.parameters,
        [key]: value
      };
      return draft;
    });
  };

  const setCurrentGroup = (groupId: string) => {
    patchSession((draft) => {
      draft.currentGroupId = groupId;
      return draft;
    });
  };

  const updateGroupProcesses = (groupId: string, updater: (processes: WorktimeProcess[]) => WorktimeProcess[]) => {
    patchSession((draft) => {
      draft.groups = draft.groups.map((group) => {
        if (group.id !== groupId) return group;
        const processes = updater(structuredClone(group.processes)).map((process, index) => ({
          ...process,
          order: index + 1,
          stationNo: index + 1
        }));
        return { ...group, processes };
      });
      return draft;
    });
  };

  const updateProcess = (groupId: string, processId: string, updater: (process: WorktimeProcess) => WorktimeProcess) => {
    updateGroupProcesses(groupId, (processes) =>
      processes.map((process) => process.id === processId ? updater(process) : process)
    );
  };

  const addProcess = (groupId: string) => {
    updateGroupProcesses(groupId, (processes) => [...processes, createProcess(processes.length + 1)]);
  };

  const deleteProcess = (groupId: string, processId: string) => {
    updateGroupProcesses(groupId, (processes) => processes.filter((process) => process.id !== processId));
  };

  const reorderProcess = (groupId: string, processId: string, targetId: string) => {
    if (processId === targetId) return;
    updateGroupProcesses(groupId, (processes) => {
      const from = processes.findIndex((process) => process.id === processId);
      const to = processes.findIndex((process) => process.id === targetId);
      if (from < 0 || to < 0) return processes;
      const next = [...processes];
      const [item] = next.splice(from, 1);
      next.splice(to, 0, item);
      return next;
    });
  };

  const exportWorkbook = async (scope: "group" | "all") => {
    if (!session) return;
    setExporting(scope);
    setMessage("");
    try {
      await persist(session);
      const response = await fetch("/api/export", {
        method: "POST",
        headers: { "Content-Type": "application/json", [ACCESS_CODE_HEADER]: accessCode },
        body: JSON.stringify({ sessionId: session.id, scope, groupId: currentGroup?.id })
      });
      if (!response.ok) {
        const data = await response.json().catch(() => null) as { error?: string } | null;
        throw new Error(data?.error || "导出失败");
      }
      const blob = await response.blob();
      const disposition = response.headers.get("Content-Disposition") || "";
      const encoded = disposition.match(/filename\*=UTF-8''([^;]+)/)?.[1];
      const fileName = encoded ? decodeURIComponent(encoded) : "工时测量与负荷山积.xlsx";
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = fileName;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    } catch (error) {
      const text = error instanceof Error ? error.message : "导出失败";
      if (text.includes("访问码")) requireAccessAgain(text);
      else setMessage(text);
    } finally {
      setExporting(null);
    }
  };

  const copyLink = async () => {
    if (!session) return;
    await navigator.clipboard.writeText(window.location.href);
    setMessage("测量单链接已复制");
  };

  const totalProcesses = session?.groups.reduce((sum, group) => sum + group.processes.length, 0) ?? 0;
  const deleteRecent = (id: string) => {
    setRecent(forgetRecentSession(id));
  };

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand">
          <ClipboardList size={28} aria-hidden="true" />
          <div>
            <h1>在线工时测量</h1>
            <p>{session ? sessionLabel(session) : "A线现场测时录入与山积图 Excel 导出"}</p>
          </div>
        </div>
        <div className="top-actions">
          {session && (
            <button className="icon-button" type="button" title="复制测量单链接" onClick={copyLink}>
              <Copy size={18} />
            </button>
          )}
          {accessCode && (
            <button className="icon-button" type="button" title="退出访问码" onClick={logoutAccess}>
              <LogOut size={18} />
            </button>
          )}
          <button className="icon-button" type="button" title="刷新当前测量单" onClick={() => session && loadSession(session.id)}>
            <RefreshCw size={18} />
          </button>
        </div>
      </header>

      {message && <div className="notice">{message}</div>}

      {!accessCode ? (
        <AccessPanel
          accessInput={accessInput}
          checking={checkingAccess}
          onInputChange={setAccessInput}
          onSubmit={handleAccessSubmit}
        />
      ) : loading ? (
        <section className="empty-state">
          <Loader2 className="spin" size={28} />
          <span>正在读取测量单</span>
        </section>
      ) : !session ? (
        <StartPanel
          lineId={lineId}
          modelName={modelName}
          recent={recent}
          creating={creating}
          onLineChange={setLineId}
          onModelNameChange={setModelName}
          onCreate={createNewSession}
          onOpen={(id) => loadSession(id, "push")}
          onForget={deleteRecent}
        />
      ) : (
        <>
          <section className="session-toolbar">
            <div className="metric">
              <span>班组</span>
              <strong>{session.groups.length}</strong>
            </div>
            <div className="metric">
              <span>工序</span>
              <strong>{totalProcesses}</strong>
            </div>
            <div className="metric">
              <span>保存</span>
              <strong className={`save-state ${saveState}`}>{saveLabel(saveState)}</strong>
            </div>
            <div className="toolbar-buttons">
              <button className="secondary" type="button" onClick={createNewSession} disabled={creating}>
                <Plus size={16} />
                新建
              </button>
              <button className="secondary" type="button" onClick={() => exportWorkbook("group")} disabled={Boolean(exporting)}>
                {exporting === "group" ? <Loader2 className="spin" size={16} /> : <Download size={16} />}
                导出当前班组
              </button>
              <button className="primary" type="button" onClick={() => exportWorkbook("all")} disabled={Boolean(exporting)}>
                {exporting === "all" ? <Loader2 className="spin" size={16} /> : <Download size={16} />}
                导出整机
              </button>
            </div>
          </section>

          <section className="parameter-panel">
            <NumberField label="计划单班产量" value={session.parameters.plannedOutput} onChange={(value) => updateParameter("plannedOutput", value)} />
            <NumberField label="班次工时(h)" value={session.parameters.shiftHours} onChange={(value) => updateParameter("shiftHours", value ?? 11)} />
          </section>

          <nav className="group-tabs" aria-label="班组切换">
            {session.groups.map((group) => (
              <button
                key={group.id}
                type="button"
                className={group.id === currentGroup?.id ? "active" : ""}
                onClick={() => setCurrentGroup(group.id)}
              >
                {group.name}
                <span>{group.processes.length}</span>
              </button>
            ))}
          </nav>

          {currentGroup && (
            <ProcessEditor
              group={currentGroup}
              dragProcessId={dragProcessId}
              setDragProcessId={setDragProcessId}
              onReorder={reorderProcess}
              onUpdate={updateProcess}
              onAdd={addProcess}
              onDelete={deleteProcess}
            />
          )}
        </>
      )}
    </main>
  );
}

function AccessPanel(props: {
  accessInput: string;
  checking: boolean;
  onInputChange: (value: string) => void;
  onSubmit: (event: FormEvent<HTMLFormElement>) => void;
}) {
  return (
    <section className="access-panel">
      <div className="access-card">
        <ShieldCheck size={34} aria-hidden="true" />
        <h2>输入访问码</h2>
        <form onSubmit={props.onSubmit}>
          <label>
            <span>访问码</span>
            <input
              autoComplete="current-password"
              inputMode="text"
              type="password"
              value={props.accessInput}
              onChange={(event) => props.onInputChange(event.target.value)}
            />
          </label>
          <button className="primary large" type="submit" disabled={props.checking}>
            {props.checking ? <Loader2 className="spin" size={18} /> : <ShieldCheck size={18} />}
            进入
          </button>
        </form>
      </div>
    </section>
  );
}

function StartPanel(props: {
  lineId: string;
  modelName: string;
  recent: RecentSession[];
  creating: boolean;
  onLineChange: (lineId: string) => void;
  onModelNameChange: (modelName: string) => void;
  onCreate: () => void;
  onOpen: (id: string) => void;
  onForget: (id: string) => void;
}) {
  return (
    <section className="start-panel">
      <div className="start-controls">
        <label>
          <span>线体</span>
          <select value={props.lineId} onChange={(event) => props.onLineChange(event.target.value)}>
            {LINE_CONFIGS.map((line) => (
              <option key={line.id} value={line.id} disabled={!line.enabled}>
                {line.name}{line.enabled ? "" : "（预留）"}
              </option>
            ))}
          </select>
        </label>
        <label>
          <span>型号名称</span>
          <input
            autoComplete="off"
            placeholder="例如 XQG100-XXXX"
            value={props.modelName}
            onChange={(event) => props.onModelNameChange(event.target.value)}
          />
        </label>
        <button className="primary large" type="button" onClick={props.onCreate} disabled={props.creating}>
          {props.creating ? <Loader2 className="spin" size={18} /> : <Plus size={18} />}
          新建测量单
        </button>
      </div>

      {props.recent.length > 0 && (
        <div className="recent-list">
          <h2>本机最近记录</h2>
          {props.recent.map((item) => (
            <div className="recent-item" key={item.id}>
              <button className="recent-open" type="button" onClick={() => props.onOpen(item.id)}>
                <span>{item.label}</span>
                <small>{new Date(item.updatedAt).toLocaleString()}</small>
              </button>
              <button className="recent-delete" type="button" title="删除本机记录" onClick={() => props.onForget(item.id)}>
                <Trash2 size={16} />
              </button>
            </div>
          ))}
        </div>
      )}
    </section>
  );
}

function NumberField(props: {
  label: string;
  value: number | null;
  step?: string;
  onChange: (value: number | null) => void;
}) {
  return (
    <label className="number-field">
      <span>{props.label}</span>
      <input
        type="number"
        inputMode="decimal"
        step={props.step ?? "1"}
        value={formatInput(props.value)}
        onChange={(event) => props.onChange(numberOrNull(event.target.value))}
      />
    </label>
  );
}

function ProcessEditor(props: {
  group: WorktimeGroup;
  dragProcessId: string | null;
  setDragProcessId: (id: string | null) => void;
  onReorder: (groupId: string, processId: string, targetId: string) => void;
  onUpdate: (groupId: string, processId: string, updater: (process: WorktimeProcess) => WorktimeProcess) => void;
  onAdd: (groupId: string) => void;
  onDelete: (groupId: string, processId: string) => void;
}) {
  const activeDragId = useRef<string | null>(null);

  const finishDrag = useCallback(() => {
    activeDragId.current = null;
    props.setDragProcessId(null);
  }, [props]);

  useEffect(() => {
    const handlePointerMove = (event: PointerEvent) => {
      const activeId = activeDragId.current;
      if (!activeId) return;
      const target = document.elementFromPoint(event.clientX, event.clientY);
      const row = target?.closest<HTMLElement>("[data-process-id]");
      const targetId = row?.dataset.processId;
      if (targetId && targetId !== activeId) {
        props.onReorder(props.group.id, activeId, targetId);
      }
    };

    window.addEventListener("pointermove", handlePointerMove, { passive: true });
    window.addEventListener("pointerup", finishDrag);
    window.addEventListener("pointercancel", finishDrag);
    return () => {
      window.removeEventListener("pointermove", handlePointerMove);
      window.removeEventListener("pointerup", finishDrag);
      window.removeEventListener("pointercancel", finishDrag);
    };
  }, [finishDrag, props]);

  const startDrag = (event: ReactPointerEvent<HTMLButtonElement>, processId: string) => {
    if (event.pointerType === "mouse" && event.button !== 0) return;
    event.preventDefault();
    activeDragId.current = processId;
    props.setDragProcessId(processId);
  };

  return (
    <section className="process-panel">
      <div className="process-header">
        <h2>{props.group.name}</h2>
        <button className="secondary" type="button" onClick={() => props.onAdd(props.group.id)}>
          <Plus size={16} />
          新增工序
        </button>
      </div>
      <div className="process-table">
        {props.group.processes.map((process, index) => (
          <div
            key={process.id}
            className={`process-row ${props.dragProcessId === process.id ? "dragging" : ""}`}
            data-process-id={process.id}
          >
            <button
              className="drag-handle"
              type="button"
              title="拖动排序"
              aria-label={`拖动第${index + 1}道工序排序`}
              onPointerDown={(event) => startDrag(event, process.id)}
            >
              <GripVertical size={17} />
            </button>
            <div className="process-index" aria-label="工序序号">{index + 1}</div>
            <label className="process-cell process-name-field">
              <span>工序名称</span>
              <input
                aria-label="工序名称"
                value={process.name}
                onChange={(event) => props.onUpdate(props.group.id, process.id, (item) => ({ ...item, name: event.target.value }))}
              />
            </label>
            {process.samples.map((sample, sampleIndex) => (
              <label className="process-cell sample-cell" key={sampleIndex}>
                <span>{`测量${sampleIndex + 1}`}</span>
                <input
                  aria-label={`测量${sampleIndex + 1}`}
                  type="number"
                  inputMode="decimal"
                  step="0.1"
                  value={formatInput(sample)}
                  onChange={(event) => props.onUpdate(props.group.id, process.id, (item) => {
                    const samples = [...item.samples] as WorktimeProcess["samples"];
                    samples[sampleIndex] = numberOrNull(event.target.value);
                    return { ...item, samples };
                  })}
                />
              </label>
            ))}
            <label className="process-cell people-cell">
              <span>人员</span>
              <input
                aria-label="人员"
                type="number"
                min="0"
                step="0.5"
                value={formatInput(process.people)}
                onChange={(event) => props.onUpdate(props.group.id, process.id, (item) => ({ ...item, people: numberOrNull(event.target.value) }))}
              />
            </label>
            <label className="process-cell remark-cell">
              <span>备注</span>
              <input
                aria-label="备注"
                value={process.remark}
                onChange={(event) => props.onUpdate(props.group.id, process.id, (item) => ({ ...item, remark: event.target.value }))}
              />
            </label>
            <button className="delete-process-button" type="button" title="删除" onClick={() => props.onDelete(props.group.id, process.id)}>
              <Trash2 size={16} />
            </button>
          </div>
        ))}
      </div>
    </section>
  );
}

function saveLabel(state: SaveState) {
  if (state === "saving") return "保存中";
  if (state === "saved") return "已保存";
  if (state === "error") return "失败";
  return "待保存";
}
