export interface IDashboardMetric {
  label: string;
  value: string;
  detail: string;
}

export interface IPromptCategory {
  id: string;
  name: string;
  promptCount: number;
  description: string;
}

export interface ICategorySummary {
  id: number;
  title: string;
  description?: string;
  promptCount?: number;
}

export interface IDepartmentSummary {
  id: number;
  title: string;
}

export interface IAiModelSummary {
  id: number;
  title: string;
}

export interface ITagSummary {
  id: number;
  title: string;
}

export interface IUserSummary {
  id: number;
  title: string;
  email?: string;
  department?: string;
}

export interface IPromptSummary {
  id: string;
  title: string;
  category: string;
  categoryId?: number;
  aiModel: string;
  aiModelId?: number;
  description: string;
  tags: string[];
  tagIds?: number[];
  featured: boolean;
  status: 'Published' | 'Draft';
  createdBy: string;
  createdById?: number;
  createdDate: string;
  averageRating: number;
  usageCount: number;
  documentUrl?: string;
  department?: string;
  departmentId?: number;
  isDeleted?: boolean;
}

export interface IPromptDetails extends IPromptSummary {
  promptText: string;
  modifiedDate: string;
  visibility?: string;
}

export interface IPromptWritePayload {
  title: string;
  category: string;
  aiModel: string;
  description: string;
  promptText: string;
  tags: string[];
  department?: string;
  visibility?: string;
  featured?: boolean;
  status: 'Published' | 'Draft';
  documentUrl?: string;
}

export interface IPromptFilters {
  searchTerm?: string;
  category?: string;
  aiModel?: string;
  department?: string;
  tag?: string;
  status?: 'Published' | 'Draft';
}

export interface IPromptRating {
  id: number;
  promptId: number;
  rating: number;
  comment?: string;
  createdBy: string;
  createdDate: string;
}

export interface IActivityLogEntry {
  id: number;
  action: string;
  promptId?: number;
  promptTitle?: string;
  actor: string;
  details?: string;
  createdDate: string;
}

export interface ISeedProgressEntry {
  label: string;
  created: number;
  skipped: number;
  failed: number;
  status: 'pending' | 'running' | 'complete' | 'failed';
}

export interface ISeedReport {
  categories: ISeedProgressEntry;
  models: ISeedProgressEntry;
  departments: ISeedProgressEntry;
  users: ISeedProgressEntry;
  tags: ISeedProgressEntry;
  prompts: ISeedProgressEntry;
  ratings: ISeedProgressEntry;
  activities: ISeedProgressEntry;
  promptAssets: ISeedProgressEntry;
  knowledgeBase: ISeedProgressEntry;
  trainingMaterials: ISeedProgressEntry;
  brandingAssets: ISeedProgressEntry;
  errors: string[];
}

export type AppView = 'landing' | 'dashboard';
export type AppStage = 'landing' | 'loading' | 'dashboard' | 'admin';
