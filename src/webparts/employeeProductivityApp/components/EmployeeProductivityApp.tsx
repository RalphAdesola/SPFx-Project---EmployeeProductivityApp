import * as React from 'react';
import type { IEmployeeProductivityAppProps } from './IEmployeeProductivityAppProps';
import type {
  IAiModelSummary,
  IAdminSummary,
  ICategorySummary,
  IDirectoryUser,
  IDashboardMetric,
  IDepartmentSummary,
  IPromptDetails,
  IPromptFilters,
  IPersonalPromptInsights,
  IPromptWritePayload,
  IPromptSummary,
  ITagSummary,
  IUserSummary
} from './SharedTypes';

import LandingPage from './LandingPage';
import DashboardPage from './DashboardPage';
import LoadingScreen from './LoadingScreen';
import AdminPage from './AdminPage';
import PromptAssistantPage from './PromptAssistantPage';
import SharePointService, { type ISharePointPromptWritePayload } from '../services/SharePointService';

export default function EmployeeProductivityApp(props: IEmployeeProductivityAppProps): React.ReactElement {
  const [stage, setStage] = React.useState<'landing' | 'loading' | 'dashboard' | 'admin' | 'prompt-assistant'>('landing');
  const [dashboardNavItem, setDashboardNavItem] = React.useState('dashboard');
  const [searchTerm, setSearchTerm] = React.useState('');
  const [isLoadingVisible, setIsLoadingVisible] = React.useState(false);
  const [prompts, setPrompts] = React.useState<IPromptSummary[]>([]);
   const [categories, setCategories] = React.useState<ICategorySummary[]>([]);  const [departments, setDepartments] = React.useState<IDepartmentSummary[]>([]);
  const [models, setModels] = React.useState<IAiModelSummary[]>([]);
  const [tags, setTags] = React.useState<ITagSummary[]>([]);
  const [users, setUsers] = React.useState<IUserSummary[]>([]);
  const [admins, setAdmins] = React.useState<IAdminSummary[]>([]);
  const [favoritePromptIds, setFavoritePromptIds] = React.useState<number[]>([]);
  const [currentUserId, setCurrentUserId] = React.useState<number | undefined>(undefined);
  const [isLoadingData, setIsLoadingData] = React.useState(true);
  const [loadError, setLoadError] = React.useState<string | null>(null);
  const [canAccessAdmin, setCanAccessAdmin] = React.useState(false);
  const [isAdminProcessing, setIsAdminProcessing] = React.useState(false);
  const loadingTimerRef = React.useRef<number | undefined>(undefined);
  const fadeTimerRef = React.useRef<number | undefined>(undefined);
  const service = React.useMemo(() => SharePointService.initialize(props.context), [props.context]);
  const loadPrompts = React.useCallback(async (filters?: IPromptFilters): Promise<void> => {
    const [promptItems, favoriteIds] = await Promise.all([
      service.getPrompts(filters),
      service.getCurrentUserFavoritePromptIds()
    ]);

    setFavoritePromptIds(favoriteIds);

    setPrompts(promptItems);
  }, [service]);

  const loadReferenceData = React.useCallback(async (): Promise<void> => {
    const [categoryItems, departmentItems, modelItems, tagItems, userItems] = await Promise.all([
      service.getCategories(),
      service.getDepartments(),
      service.getModels(),
      service.getTags(),
      service.getUsers()
    ]);

    setCategories(categoryItems);
    setDepartments(departmentItems);
    setModels(modelItems);
    setTags(tagItems);
    setUsers(userItems);
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

  const mapFormToSharePointPayload = React.useCallback((payload: IPromptWritePayload): ISharePointPromptWritePayload => {
    const categoryId = resolveLookupId(categories, payload.category);
    const aiModelId = resolveLookupId(models, payload.aiModel);

    if (!categoryId) {
      throw new Error('Select a valid Category from the SharePoint list before saving.');
    }

    return {
      title: payload.title,
      description: payload.description,
      promptText: payload.promptText,
      categoryId,
      aiModelId,
      departmentId: resolveLookupId(departments, payload.department),
      tagIds: resolveLookupIds(tags, payload.tags),
      status: payload.status,
      visibility: payload.visibility || 'Organization',
      featured: Boolean(payload.featured),
      documentUrl: payload.documentUrl,
      promptOwnerId: undefined,
      lastReviewed: undefined
    };
  }, [categories, departments, models, resolveLookupId, resolveLookupIds, tags]);

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
        const [, , userId] = await Promise.all([loadPrompts(), loadReferenceData(), service.getCurrentUserId()]);
        setCurrentUserId(userId);

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

  React.useEffect(() => {
    if (!canAccessAdmin) {
      setAdmins([]);
      return;
    }

    void service.getAdmins().then(setAdmins).catch(() => setAdmins([]));
  }, [canAccessAdmin, service]);

  const handleFiltersChange = React.useCallback((filters: IPromptFilters): void => {
    void loadDashboardData(filters);
  }, [loadDashboardData]);

  const handleCreatePrompt = React.useCallback(async (payload: IPromptWritePayload): Promise<void> => {
    await service.createPrompt({
      ...mapFormToSharePointPayload(payload)
    });
    await loadDashboardData();
  }, [service, loadDashboardData, mapFormToSharePointPayload]);

  const handleUpdatePrompt = React.useCallback(async (id: number, payload: IPromptWritePayload): Promise<void> => {
    await service.updatePrompt(id, {
      ...mapFormToSharePointPayload(payload)
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

  const handleAddAdmin = React.useCallback(async (email: string): Promise<void> => {
    await service.addAdmin(email);
    setAdmins(await service.getAdmins());
  }, [service]);

  const handleRemoveAdmin = React.useCallback(async (listItemId: number): Promise<void> => {
    await service.removeAdmin(listItemId);
    setAdmins(await service.getAdmins());
  }, [service]);

  const handleSearchDirectoryUsers = React.useCallback(async (searchText: string): Promise<IDirectoryUser[]> => {
    return service.searchDirectoryUsers(searchText);
  }, [service]);

  const handleSetPromptFavorite = React.useCallback(async (id: number, title: string, isFavorite: boolean): Promise<void> => {
    await service.setPromptFavorite(id, title, isFavorite);
    setFavoritePromptIds((current) => isFavorite
      ? Array.from(new Set([...current, id]))
      : current.filter((promptId) => promptId !== id));
  }, [service]);

  const myPrompts = React.useMemo(() => prompts.filter((prompt) => !prompt.isDeleted && prompt.createdById === currentUserId), [currentUserId, prompts]);

  const personalMetrics: IDashboardMetric[] = React.useMemo(() => {
    const publishedPrompts = myPrompts.filter((prompt) => prompt.status === 'Published');
    const draftPrompts = myPrompts.filter((prompt) => prompt.status === 'Draft');
    return [
      { label: 'Total Prompts', value: String(myPrompts.length), detail: 'Prompts you created' },
      { label: 'Published', value: String(publishedPrompts.length), detail: 'Available to your audience' },
      { label: 'Drafts', value: String(draftPrompts.length), detail: 'One saved draft allowed' }
    ];
  }, [myPrompts]);

  const personalInsights: IPersonalPromptInsights = React.useMemo(() => {
    const mostUsed = [...myPrompts].sort((left, right) => right.usageCount - left.usageCount)[0];
    const recentlyAdded = [...myPrompts].sort((left, right) => new Date(right.createdDate).getTime() - new Date(left.createdDate).getTime())[0];
    return {
      mostUsed: mostUsed?.title || 'No prompt activity yet',
      recentlyAdded: recentlyAdded?.title || 'No prompts created yet'
    };
  }, [myPrompts]);

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
          metrics={personalMetrics}
          categories={categories}
          departments={departments}
          models={models}
          tags={tags}
          prompts={prompts}
          favoritePromptIds={favoritePromptIds}
          myPrompts={myPrompts}
          personalInsights={personalInsights}
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
          onSetPromptFavorite={handleSetPromptFavorite}
          onNavigateToAdmin={() => setStage('admin')}
          onNavigateToPromptAssistant={() => setStage('prompt-assistant')}
          canAccessAdmin={canAccessAdmin}
        />
      );
    }

    return (
      <AdminPage
        canAccess={canAccessAdmin}
        isProcessing={isAdminProcessing}
        users={users}
        admins={admins}
        prompts={prompts.filter((prompt) => !prompt.isDeleted)}
        onAddAdmin={handleAddAdmin}
        onRemoveAdmin={handleRemoveAdmin}
        onSearchDirectoryUsers={handleSearchDirectoryUsers}
        onDeletePrompt={handleDeletePrompt}
        onBack={() => setStage('dashboard')}
      />
    );
  }

  if (stage === 'prompt-assistant') {
    return <PromptAssistantPage categories={categories} departments={departments} tags={tags} onCreatePrompt={handleCreatePrompt} onBack={() => { setDashboardNavItem('dashboard'); setStage('dashboard'); }} onSaveSuccess={() => { setDashboardNavItem('my-prompts'); setStage('dashboard'); }} />;
  }

  return (
    <DashboardPage
      displayName={props.userDisplayName}
      metrics={personalMetrics}
      categories={categories}
      departments={departments}
      models={models}
      tags={tags}
      prompts={prompts}
      favoritePromptIds={favoritePromptIds}
      myPrompts={myPrompts}
      personalInsights={personalInsights}
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
      onSetPromptFavorite={handleSetPromptFavorite}
      onNavigateToAdmin={() => setStage('admin')}
      onNavigateToPromptAssistant={() => setStage('prompt-assistant')}
      initialNavItem={dashboardNavItem}
      canAccessAdmin={canAccessAdmin}
    />
  );
}
