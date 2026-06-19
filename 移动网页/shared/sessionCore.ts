import { DEFAULT_PARAMETERS, findLine, findModel } from "../src/data/presets.js";
import type {
  SampleValue,
  WorktimeGroup,
  WorktimeParameters,
  WorktimeProcess,
  WorktimeSession
} from "../src/types.js";

const MAX_GROUPS = 24;
const MAX_PROCESSES = 600;
const MAX_REQUEST_BYTES = 1_000_000;
const DEFAULT_TEMPLATE_MODEL_ID = "ordinary-washer-dryer";
const FALLBACK_MODEL_NAME = "洗烘结构";

function makeId(prefix: string) {
  return `${prefix}_${crypto.randomUUID().replace(/-/g, "").slice(0, 18)}`;
}

function cloneParameters(parameters?: Partial<WorktimeParameters>): WorktimeParameters {
  return {
    ...DEFAULT_PARAMETERS,
    ...parameters
  };
}

function normalizeNumber(value: unknown): number | null {
  if (value === "" || value === null || value === undefined) return null;
  const num = Number(value);
  return Number.isFinite(num) ? num : null;
}

function normalizeText(value: unknown, maxLength = 80) {
  return typeof value === "string" ? value.trim().slice(0, maxLength) : "";
}

function normalizeSamples(samples: unknown): WorktimeProcess["samples"] {
  const list = Array.isArray(samples) ? samples : [];
  return [0, 1, 2, 3, 4].map((index) => normalizeNumber(list[index])) as WorktimeProcess["samples"];
}

function normalizeProcess(process: Partial<WorktimeProcess>, index: number): WorktimeProcess {
  const people = normalizeNumber(process.people);
  return {
    id: typeof process.id === "string" && process.id ? process.id : makeId("proc"),
    order: index + 1,
    stationNo: index + 1,
    sourceKey: typeof process.sourceKey === "string" ? process.sourceKey : "",
    name: typeof process.name === "string" ? process.name : "",
    samples: normalizeSamples(process.samples),
    people: people && people > 0 ? people : null,
    remark: typeof process.remark === "string" ? process.remark : ""
  };
}

function normalizeGroups(groups: unknown): WorktimeGroup[] {
  if (!Array.isArray(groups)) return [];
  const normalized = groups.slice(0, MAX_GROUPS).map((group, groupIndex) => {
    const rawGroup = group as Partial<WorktimeGroup>;
    const processes = Array.isArray(rawGroup.processes) ? rawGroup.processes : [];
    return {
      id: typeof rawGroup.id === "string" && rawGroup.id ? rawGroup.id : makeId("group"),
      name: typeof rawGroup.name === "string" ? rawGroup.name : `班组${groupIndex + 1}`,
      order: groupIndex + 1,
      processes: processes
        .slice(0, MAX_PROCESSES)
        .map((process, processIndex) => normalizeProcess(process as Partial<WorktimeProcess>, processIndex))
    };
  });

  const totalProcesses = normalized.reduce((sum, group) => sum + group.processes.length, 0);
  if (totalProcesses <= MAX_PROCESSES) return normalized;

  let remaining = MAX_PROCESSES;
  return normalized.map((group) => {
    const processes = group.processes.slice(0, Math.max(remaining, 0));
    remaining -= processes.length;
    return { ...group, processes };
  });
}

function assertRequestSize(session: WorktimeSession) {
  if (Buffer.byteLength(JSON.stringify(session), "utf8") > MAX_REQUEST_BYTES) {
    throw new Error("请求内容过大");
  }
}

export function isValidSessionId(id: string) {
  return /^wt_[a-f0-9]{12,}$/i.test(id);
}

export function createSession(lineId: string, modelId = DEFAULT_TEMPLATE_MODEL_ID, modelName = ""): WorktimeSession {
  const line = findLine(lineId);
  const model = findModel(lineId, modelId);
  if (!line?.enabled || !model?.enabled) {
    throw new Error("该线体或机型暂不可用");
  }

  const now = new Date().toISOString();
  const groups = structuredClone(model.groups).map((group, groupIndex) => ({
    ...group,
    order: groupIndex + 1,
    processes: group.processes.map((process, processIndex) => ({
      ...process,
      order: processIndex + 1,
      stationNo: processIndex + 1,
      name: "",
      samples: [null, null, null, null, null] as [
        SampleValue,
        SampleValue,
        SampleValue,
        SampleValue,
        SampleValue
      ],
      people: process.people ?? 1,
      remark: ""
    }))
  }));

  return {
    id: makeId("wt"),
    lineId,
    modelId,
    modelName: normalizeText(modelName) || FALLBACK_MODEL_NAME,
    currentGroupId: groups[0]?.id ?? "",
    parameters: cloneParameters(),
    groups,
    createdAt: now,
    updatedAt: now
  };
}

export function normalizeSession(session: WorktimeSession): WorktimeSession {
  assertRequestSize(session);
  const line = findLine(session.lineId);
  const model = findModel(session.lineId, session.modelId);
  if (!line?.enabled || !model?.enabled) {
    throw new Error("该线体或机型暂不可用");
  }

  const groups = normalizeGroups(session.groups);
  const currentGroupId = groups.some((group) => group.id === session.currentGroupId)
    ? session.currentGroupId
    : groups[0]?.id ?? "";

  return {
    id: typeof session.id === "string" && session.id ? session.id : makeId("wt"),
    lineId: session.lineId,
    modelId: session.modelId,
    modelName: normalizeText(session.modelName) || undefined,
    currentGroupId,
    parameters: cloneParameters(session.parameters),
    groups,
    createdAt: session.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
}
