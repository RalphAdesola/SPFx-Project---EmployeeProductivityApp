import * as React from 'react';
import type { IEmployeeProductivityAppProps } from './IEmployeeProductivityAppProps';
import type { IDashboardMetric, IPromptCategory, IPromptSummary } from './SharedTypes';
import LandingPage from './LandingPage';
import DashboardPage from './DashboardPage';
import LoadingScreen from './LoadingScreen';

export default function EmployeeProductivityApp(props: IEmployeeProductivityAppProps): React.ReactElement {
  const [stage, setStage] = React.useState<'landing' | 'loading' | 'dashboard'>('landing');
  const [searchTerm, setSearchTerm] = React.useState('');
  const [isLoadingVisible, setIsLoadingVisible] = React.useState(false);
  const loadingTimerRef = React.useRef<number | undefined>(undefined);
  const fadeTimerRef = React.useRef<number | undefined>(undefined);

  const dashboardMetrics: IDashboardMetric[] = [
    { label: 'Prompts', value: '248', detail: 'Across enterprise libraries' },
    { label: 'Categories', value: '18', detail: 'Curated prompt collections' },
    { label: 'Contributors', value: '42', detail: 'Trusted content owners' },
    { label: 'Featured', value: '12', detail: 'Recommended by governance' },
  ];

  const promptCategories: IPromptCategory[] = [
    { id: 'sales', name: 'Sales Enablement', promptCount: 34, description: 'Outbound, discovery, follow-up, and proposal prompts.' },
    { id: 'hr', name: 'HR & People Ops', promptCount: 21, description: 'Employee support, policy, onboarding, and communication prompts.' },
    { id: 'ops', name: 'Operations', promptCount: 28, description: 'Process automation, documentation, and reporting prompts.' },
    { id: 'it', name: 'IT Support', promptCount: 19, description: 'Incident response, troubleshooting, and user assistance prompts.' },
  ];

  const promptLibrary: IPromptSummary[] = [
    { id: 'p1', title: 'Executive Meeting Summary', category: 'Operations', aiModel: 'GPT-4o', description: 'Condense meeting notes into decisions, actions, and owners.', tags: ['summary', 'leadership', 'meeting'], featured: true, status: 'Published', createdBy: 'Victor Osahon', createdDate: 'Jul 19, 2026', averageRating: 4.9, usageCount: 124 },
    { id: 'p2', title: 'Customer Response Draft', category: 'Sales Enablement', aiModel: 'Claude 3.5 Sonnet', description: 'Generate a polished reply that addresses objections with empathy.', tags: ['customer', 'email', 'sales'], featured: true, status: 'Published', createdBy: 'Miracle Shodeinde', createdDate: 'Jul 18, 2026', averageRating: 4.8, usageCount: 98 },
    { id: 'p3', title: 'Policy Q&A Assistant', category: 'HR & People Ops', aiModel: 'GPT-4.1', description: 'Create clear responses for common internal policy questions.', tags: ['policy', 'hr', 'support'], featured: false, status: 'Draft', createdBy: 'Sijibomi Adeoye', createdDate: 'Jul 16, 2026', averageRating: 4.6, usageCount: 51 },
    { id: 'p4', title: 'Incident Update Generator', category: 'IT Support', aiModel: 'Gemini 2.5 Pro', description: 'Turn incident notes into a concise stakeholder update.', tags: ['incident', 'status', 'it'], featured: false, status: 'Published', createdBy: 'Godwin Onah', createdDate: 'Jul 15, 2026', averageRating: 4.7, usageCount: 76 },
    { id: 'p5', title: 'Microsoft Train the Trainer App in a Day', category: 'IT Support', aiModel: 'GPT-4o', description: 'Create a structured training prompt for internal Microsoft adoption workshops.', tags: ['training', 'microsoft', 'enablement'], featured: false, status: 'Published', createdBy: 'Victor Alfa', createdDate: 'Jul 20, 2026', averageRating: 4.8, usageCount: 42 },
  ];

  React.useEffect(() => {
    return () => {
      if (loadingTimerRef.current) {
        window.clearTimeout(loadingTimerRef.current);
      }
      if (fadeTimerRef.current) {
        window.clearTimeout(fadeTimerRef.current);
      }
    };
  }, []);

  const handleProceed = (): void => {
    setStage('loading');
    setIsLoadingVisible(true);

    if (loadingTimerRef.current) {
      window.clearTimeout(loadingTimerRef.current);
    }

    if (fadeTimerRef.current) {
      window.clearTimeout(fadeTimerRef.current);
    }

    fadeTimerRef.current = window.setTimeout(() => {
      setIsLoadingVisible(false);
    }, 4500);

    loadingTimerRef.current = window.setTimeout(() => {
      setStage('dashboard');
      setIsLoadingVisible(false);
    }, 5000);
  };

  if (stage === 'loading') {
    return <LoadingScreen isVisible={isLoadingVisible} />;
  }

  return stage === 'landing' ? (
    <LandingPage onProceed={handleProceed} />
  ) : (
    <DashboardPage
      displayName={props.userDisplayName}
      metrics={dashboardMetrics}
      categories={promptCategories}
      prompts={promptLibrary}
      searchTerm={searchTerm}
      onSearchTermChange={setSearchTerm}
      onBackToLanding={() => setStage('landing')}
    />
  );
}
