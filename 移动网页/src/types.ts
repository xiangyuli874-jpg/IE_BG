export type ExportScope = "group" | "all";

export type SampleValue = number | null;

export interface WorktimeParameters {
  plannedOutput: number | null;
  shiftHours: number;
  utilization: number;
  allowance: number;
  skill: number;
  effort: number;
  environment: number;
  consistency: number;
}

export interface WorktimeProcess {
  id: string;
  order: number;
  stationNo: number | null;
  sourceKey: string;
  name: string;
  samples: [SampleValue, SampleValue, SampleValue, SampleValue, SampleValue];
  people: number | null;
  remark: string;
}

export interface WorktimeGroup {
  id: string;
  name: string;
  order: number;
  processes: WorktimeProcess[];
}

export interface WorktimeSession {
  id: string;
  lineId: string;
  modelId: string;
  modelName?: string;
  currentGroupId: string;
  parameters: WorktimeParameters;
  groups: WorktimeGroup[];
  createdAt: string;
  updatedAt: string;
}

export interface LineConfig {
  id: string;
  name: string;
  enabled: boolean;
  models: ModelConfig[];
}

export interface ModelConfig {
  id: string;
  name: string;
  enabled: boolean;
  sourceModelId?: string;
  groups: WorktimeGroup[];
}

export interface ExportRequest {
  sessionId: string;
  scope: ExportScope;
  groupId?: string;
}
