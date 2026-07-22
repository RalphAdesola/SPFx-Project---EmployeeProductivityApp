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
  description: string;
  tags: string[];
  featured: boolean;
  status: 'Published' | 'Draft';
  createdBy: string;
  modified: string;
}

export interface IDashboardMetric {
  label: string;
  value: string;
  detail: string;
}

