import * as React from 'react';
import type { IEmployeeProductivityAppProps } from './IEmployeeProductivityAppProps';
import type {
  IAiModelSummary,
  ICategorySummary,
  IDashboardMetric,
  IDepartmentSummary,
  IPromptDetails,
  IPromptFilters,
  IPromptWritePayload,
  IPromptSummary,
  ITagSummary  
} from './SharedTypes';

import LandingPage from './LandingPage';
import DashboardPage from './DashboardPage';
import LoadingScreen from './LoadingScreen';
import AdminPage from './AdminPage';
import SharePointService, { type ISharePointPromptWritePayload } from '../services/SharePointService';
import type { ISeedReport } from './SharedTypes';

const readFileAsArrayBuffer = (file: File): Promise<ArrayBuffer> => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = () => resolve(reader.result as ArrayBuffer);
  reader.onerror = () => reject(reader.error || new Error('Unable to read selected file.'));
  reader.readAsArrayBuffer(file);
});

export default function EmployeeProductivityApp(props: IEmployeeProductivityAppProps): React.ReactElement {
  const [stage, setStage] = React.useState<'landing' | 'loading' | 'dashboard' | 'admin'>('landing');
  const [searchTerm, setSearchTerm] = React.useState('');
  const [isLoadingVisible, setIsLoadingVisible] = React.useState(false);
  const [prompts, setPrompts] = React.useState<IPromptSummary[]>([]);
   const [categories, setCategories] = React.useState<ICategorySummary[]>([]);  const [departments, setDepartments] = React.useState<IDepartmentSummary[]>([]);
  const [models, setModels] = React.useState<IAiModelSummary[]>([]);
  const [tags, setTags] = React.useState<ITagSummary[]>([]);
  const [isLoadingData, setIsLoadingData] = React.useState(true);
  const [loadError, setLoadError] = React.useState<string | null>(null);
  const [canAccessAdmin, setCanAccessAdmin] = React.useState(false);
  const [adminReport, setAdminReport] = React.useState<ISeedReport | null>(null);
  const [adminLog, setAdminLog] = React.useState<string[]>([]);
  const [isAdminProcessing, setIsAdminProcessing] = React.useState(false);
  const loadingTimerRef = React.useRef<number | undefined>(undefined);
  const fadeTimerRef = React.useRef<number | undefined>(undefined);
  const service = React.useMemo(() => SharePointService.initialize(props.context), [props.context]);
  const loadPrompts = React.useCallback(async (filters?: IPromptFilters): Promise<void> => {
    const [promptItems, ratingItems] = await Promise.all([
      service.getPrompts(filters),
      service.getRatings()
    ]);

    setPrompts(promptItems.map((prompt) => {
      const promptRatings = ratingItems.filter((rating) => String(rating.promptId) === prompt.id);
      const averageRating = promptRatings.length
        ? promptRatings.reduce((sum, rating) => sum + rating.rating, 0) / promptRatings.length
        : prompt.averageRating;

      return {
        ...prompt,
        averageRating: Number(averageRating.toFixed(1))
      };
    }));
  }, [service]);

  const loadReferenceData = React.useCallback(async (): Promise<void> => {
    const [categoryItems, departmentItems, modelItems, tagItems] = await Promise.all([
      service.getCategories(),
      service.getDepartments(),
      service.getModels(),
      service.getTags()
    ]);

    setCategories(categoryItems);
    setDepartments(departmentItems);
    setModels(modelItems);
    setTags(tagItems);
  }, [service]);

  const resolveLookupId = React.useCallback(<TItem extends { id: number; title: string }>(items: TItem[], title: string | undefined): number | undefined => {
    const match = items.find((item) => item.title === title);
    return match?.id;
  }, []);

  const resolveLookupIds = React.useCallback((items: ITagSummary[], titles: string[] | undefined): number[] => {
    if (!titles) {
      return [];
    }

    return titles
      .map((title) => items.find((item) => item.title === title)?.id)
      .filter((value): value is number => typeof value === 'number');
  }, []);

  const mapFormToSharePointPayload = React.useCallback((payload: IPromptWritePayload): ISharePointPromptWritePayload => ({
    title: payload.title,
    description: payload.description,
    promptText: payload.promptText,
    categoryId: resolveLookupId(categories, payload.category) || 0,
    aiModelId: resolveLookupId(models, payload.aiModel) || 0,
    departmentId: resolveLookupId(departments, payload.department),
    tagIds: resolveLookupIds(tags, payload.tags),
    status: payload.status,
    visibility: payload.visibility || 'Organization',
    featured: Boolean(payload.featured),
    documentUrl: payload.documentUrl,
    promptOwnerId: undefined,
    lastReviewed: undefined
  }), [categories, departments, models, resolveLookupId, resolveLookupIds, tags]);

  const loadDashboardData = React.useCallback(async (filters?: IPromptFilters): Promise<void> => {
    setIsLoadingData(true);
    setLoadError(null);

    try {
      await Promise.all([
        loadPrompts(filters),
        loadReferenceData()
      ]);
    } catch (error) {
      setLoadError(error instanceof Error ? error.message : 'Unable to load SharePoint data.');
      setPrompts([]);
    } finally {
      setIsLoadingData(false);
    }
  }, [loadPrompts, loadReferenceData]);

  React.useEffect(() => {
    let active = true;

    const load = async (): Promise<void> => {
      setIsLoadingData(true);
      setLoadError(null);

      try {
        await Promise.all([loadPrompts(), loadReferenceData()]);

        if (!active) {
          return;
        }
      } catch (error) {
        if (!active) {
          return;
        }

        setLoadError(error instanceof Error ? error.message : 'Unable to load SharePoint data.');
      } finally {
        if (active) {
          setIsLoadingData(false);
        }
      }
    };

    void load();

    return () => {
      active = false;

      if (loadingTimerRef.current) {
        window.clearTimeout(loadingTimerRef.current);
      }

      if (fadeTimerRef.current) {
        window.clearTimeout(fadeTimerRef.current);
      }
    };
  }, [loadPrompts, loadReferenceData]);

  React.useEffect(() => {
    let active = true;

    const loadPermissions = async (): Promise<void> => {
      try {
        const allowed = await service.isCurrentUserAdmin();
        if (active) {
          setCanAccessAdmin(allowed);
        }
      } catch {
        if (active) {
          setCanAccessAdmin(false);
        }
      }
    };

    void loadPermissions();

    return () => {
      active = false;
    };
  }, [service]);

  React.useEffect(() => {
    if (stage === 'admin' && !canAccessAdmin) {
      setStage('dashboard');
    }
  }, [stage, canAccessAdmin]);

  const handleFiltersChange = React.useCallback((filters: IPromptFilters): void => {
    void loadDashboardData(filters);
  }, [loadDashboardData]);

  const handleCreatePrompt = React.useCallback(async (payload: IPromptWritePayload, file?: File): Promise<void> => {
    let documentUrl = payload.documentUrl;

    if (file) {
      const upload = await service.uploadPromptAsset(file.name, await readFileAsArrayBuffer(file));
      documentUrl = upload.serverRelativeUrl;
    }

    await service.createPrompt({
      ...mapFormToSharePointPayload({ ...payload, documentUrl }),
      documentUrl
    });
    await loadDashboardData();
  }, [service, loadDashboardData, mapFormToSharePointPayload]);

  const handleUpdatePrompt = React.useCallback(async (id: number, payload: IPromptWritePayload, file?: File): Promise<void> => {
    let documentUrl = payload.documentUrl;

    if (file) {
      const upload = await service.uploadPromptAsset(file.name, await readFileAsArrayBuffer(file));
      documentUrl = upload.serverRelativeUrl;
    }

    await service.updatePrompt(id, {
      ...mapFormToSharePointPayload({ ...payload, documentUrl }),
      documentUrl
    });
    await loadDashboardData();
  }, [service, loadDashboardData, mapFormToSharePointPayload]);

  const handleViewPrompt = React.useCallback(async (id: number): Promise<IPromptDetails | null> => {
    return service.getPromptById(id);
  }, [service]);

  const handleDeletePrompt = React.useCallback(async (id: number): Promise<void> => {
    await service.deletePrompt(id);
    await loadDashboardData();
  }, [service, loadDashboardData]);

  const handleCopyPrompt = React.useCallback(async (id: number): Promise<string> => {
    const prompt = await service.getPromptById(id);

    if (!prompt) {
      throw new Error('Prompt was not found.');
    }

    await service.incrementUsage(id);
    await loadDashboardData();
    return prompt.promptText;
  }, [service, loadDashboardData]);

  const handleSubmitRating = React.useCallback(async (id: number, rating: number): Promise<void> => {
    await service.submitRating(id, rating);
    await loadDashboardData();
  }, [service, loadDashboardData]);

  const handleSeedAllData = React.useCallback(async (): Promise<void> => {
    if (!canAccessAdmin) {
      setStage('dashboard');
      return;
    }

    setIsAdminProcessing(true);
    setAdminLog([]);
    try {
      const report = await service.seedDemoData({
        onProgress: (message) => setAdminLog((current) => [...current, message])
      });
      setAdminReport(report);
      await loadDashboardData();
    } finally {
      setIsAdminProcessing(false);
    }
  }, [canAccessAdmin, service, loadDashboardData]);

  const handleClearDemoData = React.useCallback(async (): Promise<void> => {
    if (!canAccessAdmin) {
      setStage('dashboard');
      return;
    }

    setIsAdminProcessing(true);
    setAdminLog([]);
    try {
      const report = await service.clearDemoData({
        onProgress: (message) => setAdminLog((current) => [...current, message])
      });
      setAdminReport(report);
      await loadDashboardData();
    } finally {
      setIsAdminProcessing(false);
    }
  }, [canAccessAdmin, service, loadDashboardData]);

  const dashboardMetrics: IDashboardMetric[] = React.useMemo(() => {
    const activePrompts = prompts.filter((prompt) => !prompt.isDeleted);
    const publishedPrompts = activePrompts.filter((prompt) => prompt.status === 'Published');
    const draftPrompts = activePrompts.filter((prompt) => prompt.status === 'Draft');
    const featuredPrompts = activePrompts.filter((prompt) => prompt.featured);
    const averageRating = activePrompts.length
      ? activePrompts.reduce((sum, prompt) => sum + prompt.averageRating, 0) / activePrompts.length
      : 0;
    const totalUsage = activePrompts.reduce((sum, prompt) => sum + prompt.usageCount, 0);
    const mostUsedPrompt = [...activePrompts].sort((left, right) => right.usageCount - left.usageCount)[0];
    const recentPrompt = [...activePrompts].sort((left, right) => new Date(right.createdDate).getTime() - new Date(left.createdDate).getTime())[0];
    const topRatedPrompt = [...activePrompts].sort((left, right) => right.averageRating - left.averageRating)[0];

    return [
      { label: 'Total Prompts', value: String(activePrompts.length), detail: 'Across enterprise libraries' },
      { label: 'Published Prompts', value: String(publishedPrompts.length), detail: 'Ready for use' },
      { label: 'Draft Prompts', value: String(draftPrompts.length), detail: 'In review or work-in-progress' },
      { label: 'Featured Prompts', value: String(featuredPrompts.length), detail: 'Highlighted by governance' },
      { label: 'Average Rating', value: averageRating ? averageRating.toFixed(1) : '0.0', detail: 'Across submitted ratings' },
      { label: 'Total Usage', value: String(totalUsage), detail: 'Prompt executions and copies' },
      { label: 'Most Used Prompt', value: mostUsedPrompt?.title || '—', detail: 'Highest usage count' },
      { label: 'Recently Added Prompt', value: recentPrompt?.title || '—', detail: 'Latest SharePoint record' },
      { label: 'Top Rated Prompt', value: topRatedPrompt?.title || '—', detail: 'Highest average rating' }
    ];
  }, [prompts]);

  const handleProceed = (): void => {
    setStage('loading');
    setIsLoadingVisible(true);

    if (loadingTimerRef.current) {
      window.clearTimeout(loadingTimerRef.current);
    }

    if (fadeTimerRef.current) {
      window.clearTimeout(fadeTimerRef.current);
    }

    fadeTimerRef.current = window.setTimeout(() => setIsLoadingVisible(false), 4500);
    loadingTimerRef.current = window.setTimeout(() => {
      setStage('dashboard');
      setIsLoadingVisible(false);
    }, 5000);
  };

  if (stage === 'loading') {
    return <LoadingScreen isVisible={isLoadingVisible} />;
  }

  if (stage === 'landing') {
    return <LandingPage onProceed={handleProceed} />;
  }

  if (stage === 'admin') {
    if (!canAccessAdmin) {
      return (
        <DashboardPage
          displayName={props.userDisplayName}
          metrics={dashboardMetrics}
          categories={categories}
          departments={departments}
          models={models}
          tags={tags}
          prompts={prompts}
          searchTerm={searchTerm}
          onSearchTermChange={setSearchTerm}
          onBackToLanding={() => setStage('landing')}
          isLoadingData={isLoadingData}
          loadError={loadError}
          onRefresh={() => void loadDashboardData()}
          onFiltersChange={handleFiltersChange}
          onCreatePrompt={handleCreatePrompt}
          onUpdatePrompt={handleUpdatePrompt}
          onViewPrompt={handleViewPrompt}
          onDeletePrompt={handleDeletePrompt}
          onCopyPrompt={handleCopyPrompt}
          onSubmitRating={handleSubmitRating}
          onNavigateToAdmin={() => setStage('admin')}
          canAccessAdmin={canAccessAdmin}
        />
      );
    }

    return (
      <AdminPage
        canAccess={canAccessAdmin}
        isProcessing={isAdminProcessing}
        progressLog={adminLog}
        report={adminReport}
        onSeedAllData={handleSeedAllData}
        onClearDemoData={handleClearDemoData}
        onBack={() => setStage('dashboard')}
      />
    );
  }

  return (
    <DashboardPage
      displayName={props.userDisplayName}
      metrics={dashboardMetrics}
      categories={categories}
      departments={departments}
      models={models}
      tags={tags}
      prompts={prompts}
      searchTerm={searchTerm}
      onSearchTermChange={setSearchTerm}
      onBackToLanding={() => setStage('landing')}
      isLoadingData={isLoadingData}
      loadError={loadError}
      onRefresh={() => void loadDashboardData()}
      onFiltersChange={handleFiltersChange}
      onCreatePrompt={handleCreatePrompt}
      onUpdatePrompt={handleUpdatePrompt}
      onViewPrompt={handleViewPrompt}
      onDeletePrompt={handleDeletePrompt}
      onCopyPrompt={handleCopyPrompt}
      onSubmitRating={handleSubmitRating}
      onNavigateToAdmin={() => setStage('admin')}
      canAccessAdmin={canAccessAdmin}
    />
  );
}
