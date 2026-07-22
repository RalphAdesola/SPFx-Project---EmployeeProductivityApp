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

export interface IPromptSummary {
  id: string;
  title: string;
  category: string;
  aiModel: string;
  description: string;
  tags: string[];
  featured: boolean;
  status: 'Published' | 'Draft';
  createdBy: string;
  createdDate: string;
  averageRating: number;
  usageCount: number;
}

export type AppView = 'landing' | 'dashboard';
export type AppStage = 'landing' | 'loading' | 'dashboard';
