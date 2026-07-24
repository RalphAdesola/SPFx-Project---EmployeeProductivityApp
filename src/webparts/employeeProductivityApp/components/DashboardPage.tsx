import * as React from 'react';
import {
  Avatar,
  Badge,
  Body1,
  Body2,
  Button,
  Caption1,
  Card,
  CardHeader,
  Divider,
  Input,
  Text,
  Title1,
  Title2,
  Dropdown,
  Option,
  Textarea,
  Field,
  makeStyles
} from '@fluentui/react-components';
import {
  AddRegular,
  ChevronDoubleLeftRegular,
  ChevronDoubleRightRegular,
  ClockRegular,
  CopyRegular,
  DeleteRegular,
  EyeRegular,
  HeartRegular,
  HeartFilled,
  HomeRegular,
  PersonRegular,
  SearchRegular,
  SettingsRegular,
  CheckmarkCircleRegular,
  SignOutRegular
} from '@fluentui/react-icons';
import styles from './EmployeeProductivityApp.module.scss';
import type {
  IAiModelSummary,
  IDashboardMetric,
  IDepartmentSummary,
  ICategorySummary,
  IPersonalPromptInsights,
  IPromptDetails,
  IPromptFilters,
  IPromptSummary,
  IPromptWritePayload,
  ITagSummary
} from './SharedTypes';

export interface IDashboardPageProps {
  displayName: string;
  metrics: IDashboardMetric[];
  categories: ICategorySummary[];
  departments: IDepartmentSummary[];
  models: IAiModelSummary[];
  tags: ITagSummary[];
  prompts: IPromptSummary[];
  favoritePromptIds: number[];
  myPrompts: IPromptSummary[];
  personalInsights: IPersonalPromptInsights;
  searchTerm: string;
  onSearchTermChange: (value: string) => void;
  onBackToLanding: () => void;
  isLoadingData?: boolean;
  loadError?: string | null;
  onRefresh?: () => void;
  onFiltersChange: (filters: IPromptFilters) => void;
  onCreatePrompt: (payload: IPromptWritePayload) => Promise<void>;
  onUpdatePrompt: (id: number, payload: IPromptWritePayload) => Promise<void>;
  onViewPrompt: (id: number) => Promise<IPromptDetails | null>;
  onDeletePrompt: (id: number) => Promise<void>;
  onCopyPrompt: (id: number) => Promise<string>;
  onSetPromptFavorite: (id: number, title: string, isFavorite: boolean) => Promise<void>;
  onNavigateToAdmin: () => void;
  onNavigateToPromptAssistant: () => void;
  initialNavItem?: string;
  canAccessAdmin: boolean;
}

const useStyles = makeStyles({
  pageSurface: {
    backgroundColor: '#f8fafc'
  }
});

const navigationItems = [
  { key: 'dashboard', label: 'Dashboard', icon: <HomeRegular /> },
  { key: 'prompt-assistant', label: 'Prompt Assistant', icon: <AddRegular /> },
  { key: 'my-prompts', label: 'My Prompts', icon: <PersonRegular /> },
  { key: 'favorites', label: 'Favorites', icon: <HeartRegular /> },
  { key: 'logout', label: 'Log Out', icon: <SignOutRegular /> }
];

const visibilityOptions = ['Organization', 'Department', 'Private'];
const statusOptions: Array<'Published' | 'Draft'> = ['Published', 'Draft'];
const formatDate = (value: string): string => {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }

  const day = date.getDate() < 10 ? `0${date.getDate()}` : String(date.getDate());
  const monthNumber = date.getMonth() + 1;
  const month = monthNumber < 10 ? `0${monthNumber}` : String(monthNumber);
  return `${day}-${month}-${date.getFullYear()}`;
};

const copyTextToClipboard = async (value: string): Promise<void> => {
  if (navigator.clipboard && window.isSecureContext) {
    try {
      await navigator.clipboard.writeText(value);
      return;
    } catch {
      // Fall through to the browser-compatible copy mechanism.
    }
  }

  const textArea = document.createElement('textarea');
  textArea.value = value;
  textArea.setAttribute('readonly', '');
  textArea.style.position = 'fixed';
  textArea.style.opacity = '0';
  document.body.appendChild(textArea);
  textArea.select();

  const copied = document.execCommand('copy');
  document.body.removeChild(textArea);

  if (!copied) {
    throw new Error('Your browser blocked clipboard access. Copy the prompt text manually.');
  }
};

const emptyPromptForm: IPromptWritePayload = {
  title: '',
  category: '',
  aiModel: '',
  description: '',
  promptText: '',
  tags: [],
  department: '',
  visibility: 'Organization',
  featured: false,
  status: 'Draft',
  documentUrl: ''
};

export default function DashboardPage(props: IDashboardPageProps): React.ReactElement {
  const pageStyles = useStyles();
  const [sidebarCollapsed, setSidebarCollapsed] = React.useState(false);
  const [activeNavItem, setActiveNavItem] = React.useState(props.initialNavItem || 'dashboard');
  const [localSearch, setLocalSearch] = React.useState(props.searchTerm);
  const [selectedCategory, setSelectedCategory] = React.useState<string | undefined>(undefined);
  const [selectedModel, setSelectedModel] = React.useState<string | undefined>(undefined);
  const [selectedDepartment, setSelectedDepartment] = React.useState<string | undefined>(undefined);
  const [selectedTag, setSelectedTag] = React.useState<string | undefined>(undefined);
  const [selectedStatus, setSelectedStatus] = React.useState<'Published' | 'Draft' | undefined>(undefined);
  const [promptForm, setPromptForm] = React.useState<IPromptWritePayload>(emptyPromptForm);
  const [editingPromptId, setEditingPromptId] = React.useState<number | undefined>(undefined);
  const [selectedPrompt, setSelectedPrompt] = React.useState<IPromptDetails | null>(null);
  const [isSaving, setIsSaving] = React.useState(false);
  const [actionError, setActionError] = React.useState<string | null>(null);
  const [actionMessage, setActionMessage] = React.useState<string | null>(null);
  const hasActiveFilters = Boolean(localSearch.trim() || selectedCategory || selectedModel || selectedDepartment || selectedTag || selectedStatus);
  const categoryOptions = props.categories.map((category) => category.title);
  const modelOptions = props.models.map((model) => model.title);
  const departmentOptions = props.departments.map((department) => department.title);
  const tagOptions = props.tags.map((tag) => tag.title);
  const visibleNavigationItems = props.canAccessAdmin
    ? [...navigationItems, { key: 'admin', label: 'Administration', icon: <SettingsRegular /> }]
    : navigationItems;
  const isMyPromptsView = activeNavItem === 'my-prompts';
  const isFavoritesView = activeNavItem === 'favorites';
  const displayedPrompts = isFavoritesView
    ? props.prompts.filter((prompt) => props.favoritePromptIds.indexOf(Number(prompt.id)) >= 0)
    : isMyPromptsView ? props.myPrompts : props.prompts;

  React.useEffect(() => {
    if (props.initialNavItem) {
      setActiveNavItem(props.initialNavItem);
    }
  }, [props.initialNavItem]);

  React.useEffect(() => {
    props.onSearchTermChange(localSearch);
    props.onFiltersChange({
      searchTerm: localSearch,
      category: selectedCategory,
      aiModel: selectedModel,
      department: selectedDepartment,
      tag: selectedTag,
      status: selectedStatus
    });
  }, [localSearch, selectedCategory, selectedModel, selectedDepartment, selectedTag, selectedStatus, props.onSearchTermChange, props.onFiltersChange]);

  const updatePromptForm = <TKey extends keyof IPromptWritePayload,>(key: TKey, value: IPromptWritePayload[TKey]): void => {
    setPromptForm((current) => ({ ...current, [key]: value }));
  };

  const resetPromptForm = (): void => {
    setPromptForm(emptyPromptForm);
    setEditingPromptId(undefined);
    setActionError(null);
    setActionMessage(null);
  };

  const buildPayload = (status: 'Published' | 'Draft'): IPromptWritePayload => ({
    ...promptForm,
    status,
    tags: promptForm.tags.map((tag) => tag.trim()).filter(Boolean)
  });

  const savePrompt = async (status: 'Published' | 'Draft'): Promise<void> => {
    setIsSaving(true);
    setActionError(null);
    setActionMessage(null);

    try {
      const payload = buildPayload(status);

      if (!payload.title.trim() || !payload.category || !payload.promptText.trim()) {
        setActionError('Title, Category, and Prompt Text are required.');
        return;
      }

      if (editingPromptId) {
        await props.onUpdatePrompt(editingPromptId, payload);
        setActionMessage('Prompt updated successfully.');
      } else {
        await props.onCreatePrompt(payload);
        setActionMessage(status === 'Published' ? 'Prompt published successfully.' : 'Draft saved successfully.');
      }

      resetPromptForm();
    } catch (error) {
      setActionError(error instanceof Error ? error.message : 'Unable to save the prompt.');
    } finally {
      setIsSaving(false);
    }
  };

  const viewPrompt = async (promptId: string): Promise<void> => {
    setActionError(null);

    try {
      const prompt = await props.onViewPrompt(Number(promptId));
      setSelectedPrompt(prompt);
    } catch (error) {
      setActionError(error instanceof Error ? error.message : 'Unable to load prompt details.');
    }
  };

  const editPrompt = async (promptId: string): Promise<void> => {
    const prompt = await props.onViewPrompt(Number(promptId));

    if (!prompt) {
      setActionError('Prompt was not found.');
      return;
    }

    setEditingPromptId(Number(prompt.id));
    setPromptForm({
      title: prompt.title,
      category: prompt.category,
      aiModel: prompt.aiModel,
      description: prompt.description,
      promptText: prompt.promptText,
      tags: prompt.tags,
      department: prompt.department || '',
      visibility: prompt.visibility || 'Organization',
      featured: prompt.featured,
      status: prompt.status,
      documentUrl: prompt.documentUrl || ''
    });
    setSelectedPrompt(null);
    setActionMessage('Editing latest SharePoint prompt record.');
  };

  const copyPrompt = async (promptId: string): Promise<void> => {
    setActionError(null);

    try {
      const prompt = selectedPrompt?.id === promptId ? selectedPrompt : undefined;

      if (prompt?.promptText) {
        await copyTextToClipboard(prompt.promptText);
        setActionMessage('Prompt copied to your clipboard.');

        try {
          await props.onCopyPrompt(Number(promptId));
          setActionMessage('Prompt copied to your clipboard and usage logged.');
        } catch {
          setActionMessage('Prompt copied to your clipboard. Usage could not be logged.');
        }
        return;
      }

      const promptText = await props.onCopyPrompt(Number(promptId));
      await copyTextToClipboard(promptText);
      setActionMessage('Prompt copied to your clipboard and usage logged.');
    } catch (error) {
      setActionError(error instanceof Error ? error.message : 'Unable to copy the prompt.');
    }
  };

  const deletePrompt = async (promptId: string): Promise<void> => {
    setActionError(null);

    try {
      await props.onDeletePrompt(Number(promptId));
      setSelectedPrompt(null);
      setActionMessage('Prompt deleted from active library.');
    } catch (error) {
      setActionError(error instanceof Error ? error.message : 'Unable to delete the prompt.');
    }
  };

  const toggleFavorite = async (prompt: IPromptSummary): Promise<void> => {
    setActionError(null);

    try {
      const isFavorite = props.favoritePromptIds.indexOf(Number(prompt.id)) >= 0;
      await props.onSetPromptFavorite(Number(prompt.id), prompt.title, !isFavorite);
      setActionMessage(isFavorite ? 'Prompt removed from favorites.' : 'Prompt added to favorites.');
    } catch (error) {
      setActionError(error instanceof Error ? error.message : 'Unable to update favorites.');
    }
  };

  return (
    <section className={`${styles.employeeProductivityApp} ${pageStyles.pageSurface} ${styles.dashboardPage}`}>
      <aside className={`${styles.sidebar} ${sidebarCollapsed ? styles.sidebarCollapsed : ''}`}>
        <div className={styles.sidebarBrandRow}>
          <div className={styles.sidebarLogo}>EP</div>
          {!sidebarCollapsed && (
            <div className={styles.sidebarBrandText}>
              <Title2 className={styles.sidebarTitle}>Employee Productivity App</Title2>
              <Caption1 className={styles.sidebarSubtitle}>Prompt Library</Caption1>
            </div>
          )}
          <Button
            appearance="subtle"
            className={styles.sidebarCollapseButton}
            onClick={() => setSidebarCollapsed((current) => !current)}
            icon={sidebarCollapsed ? <ChevronDoubleRightRegular /> : <ChevronDoubleLeftRegular />}
            aria-label={sidebarCollapsed ? 'Expand sidebar' : 'Collapse sidebar'}
          />
        </div>

        <nav className={styles.sidebarNav} aria-label="Primary">
          {visibleNavigationItems.map((item) => (
            <button
              key={item.key}
              className={`${styles.sidebarNavItem} ${activeNavItem === item.key ? styles.sidebarNavItemActive : ''}`}
              type="button"
              onClick={() => {
                setActiveNavItem(item.key);
                if (item.key === 'admin') {
                  props.onNavigateToAdmin();
                }
                if (item.key === 'prompt-assistant') {
                  props.onNavigateToPromptAssistant();
                }
                if (item.key === 'logout') {
                  props.onBackToLanding();
                }
              }}
              aria-current={activeNavItem === item.key ? 'page' : undefined}
            >
              <span className={styles.sidebarNavIcon}>{item.icon}</span>
              {!sidebarCollapsed && <span>{item.label}</span>}
            </button>
          ))}
        </nav>

      </aside>

      <div className={styles.dashboardMain}>
        <header className={styles.topHeader}>
          <div className={styles.topHeaderLeft}>
            <Title1 className={styles.pageTitle}>{isFavoritesView ? 'Favorites' : isMyPromptsView ? 'My Prompts' : 'Prompt Library'}</Title1>
            <Caption1>{isFavoritesView ? 'Prompts you saved for quick access.' : isMyPromptsView ? 'Manage prompts you created and review their performance.' : 'Browse and reuse prompts curated across the organization.'}</Caption1>
          </div>

          <div className={styles.topHeaderActions}>
            <div className={styles.searchContainer}>
              <Input
                value={localSearch}
                onChange={(_, data) => setLocalSearch(data.value)}
                contentBefore={<SearchRegular />}
                placeholder="Search by title, description, category, tags, or model"
              />
            </div>

            <Avatar name={props.displayName} />
            <Button appearance="primary" icon={<AddRegular />} onClick={resetPromptForm}>Add Prompt</Button>
          </div>
        </header>

        <div className={styles.dashboardLayout}>
          <main className={styles.mainContent}>
            {props.loadError && (
              <Card className={styles.emptyStateCard}>
                <Body1>We could not load SharePoint data.</Body1>
                <Caption1>{props.loadError}</Caption1>
                {props.onRefresh && (
                  <Button appearance="primary" onClick={props.onRefresh}>
                    Retry
                  </Button>
                )}
              </Card>
            )}

            {props.isLoadingData && (
              <Card className={styles.emptyStateCard}>
                <Body1>Loading SharePoint data...</Body1>
                <Caption1>Please wait while we load your prompt library.</Caption1>
              </Card>
            )}

            {isMyPromptsView && (
              <>
                <section className={styles.summaryGrid}>
                  {props.metrics.map((metric) => (
                    <Card key={metric.label} className={styles.metricCard}>
                      <CardHeader header={<Body2>{metric.label}</Body2>} />
                      <Title2>{metric.value}</Title2>
                      <Caption1>{metric.detail}</Caption1>
                    </Card>
                  ))}
                </section>
                <section className={styles.personalInsights} aria-label="My prompt highlights">
                  <div><Caption1 className={styles.personalInsightLabel}>Most used:</Caption1><Body2>{props.personalInsights.mostUsed}</Body2></div>
                  <div><Caption1 className={styles.personalInsightLabel}>Recently added:</Caption1><Body2>{props.personalInsights.recentlyAdded}</Body2></div>
                </section>
              </>
            )}

            <section className={styles.controlsBar}>
              <div className={styles.filterField}>
                <Field label="Category">
                  <Dropdown
                    className={styles.whiteDropdown}
                    placeholder="All categories"
                    value={selectedCategory || ''}
                    selectedOptions={selectedCategory ? [selectedCategory] : []}
                    onOptionSelect={(_, data) => {
                      setSelectedCategory(data.optionValue ?? undefined);
                    }}
                  >
                    {categoryOptions.map((option) => (
                      <Option key={option} value={option}>
                        {option}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>
              </div>

              <div className={styles.filterField}>
                <Field label="Department">
                  <Dropdown
                    className={styles.whiteDropdown}
                    placeholder="All departments"
                    value={selectedDepartment || ''}
                    selectedOptions={selectedDepartment ? [selectedDepartment] : []}
                    onOptionSelect={(_, data) => setSelectedDepartment(data.optionValue ?? undefined)}
                  >
                    {departmentOptions.map((option) => (
                      <Option key={option} value={option}>{option}</Option>
                    ))}
                  </Dropdown>
                </Field>
              </div>

              <div className={styles.filterField}>
                <Field label="Tag">
                  <Dropdown
                    className={styles.whiteDropdown}
                    placeholder="All tags"
                    value={selectedTag || ''}
                    selectedOptions={selectedTag ? [selectedTag] : []}
                    onOptionSelect={(_, data) => setSelectedTag(data.optionValue ?? undefined)}
                  >
                    {tagOptions.map((option) => (
                      <Option key={option} value={option}>{option}</Option>
                    ))}
                  </Dropdown>
                </Field>
              </div>
            </section>

            <div className={styles.filterActions}>
              <Button
                appearance="secondary"
                onClick={() => {
                  setLocalSearch('');
                  setSelectedCategory(undefined);
                  setSelectedModel(undefined);
                  setSelectedDepartment(undefined);
                  setSelectedTag(undefined);
                  setSelectedStatus(undefined);
                }}
                disabled={!hasActiveFilters}
              >
                Clear Filters
              </Button>
            </div>

            {(actionError || actionMessage) && (
              <Card className={styles.emptyStateCard}>
                {actionError && <Body1>{actionError}</Body1>}
                {actionMessage && <Body1>{actionMessage}</Body1>}
              </Card>
            )}

            <section className={styles.promptGrid}>
              {displayedPrompts.length === 0 ? (
                <Card className={styles.emptyStateCard}>
                  <Body1>No prompts match your current filters.</Body1>
                  <Caption1>Try a different search term, category, or model.</Caption1>
                </Card>
              ) : (
                displayedPrompts.map((prompt) => (
                  <Card key={prompt.id} className={styles.promptCard}>
                    <div className={styles.promptCardTop}>
                      <div>
                        <div className={styles.promptTitleRow}>
                          <Body1 className={styles.promptTitle}>{prompt.title}</Body1>
                          <Badge appearance={prompt.status === 'Published' ? 'filled' : 'outline'}>
                            {prompt.status}
                          </Badge>
                          {prompt.featured && (
                            <Badge appearance="filled" color="brand">
                              Featured
                            </Badge>
                          )}
                        </div>
                        <Caption1>{prompt.category} · {prompt.createdBy}</Caption1>
                      </div>
                      <Button
                        appearance="subtle"
                        icon={props.favoritePromptIds.indexOf(Number(prompt.id)) >= 0 ? <HeartFilled /> : <HeartRegular />}
                        aria-label={props.favoritePromptIds.indexOf(Number(prompt.id)) >= 0 ? 'Remove from favorites' : 'Add to favorites'}
                        onClick={() => toggleFavorite(prompt)}
                      />
                    </div>

                    <Body2 className={styles.promptDescription}>{prompt.description}</Body2>

                    <div className={styles.promptMetaGrid}>
                      <div className={styles.promptMetaItem}>
                        <Text>Owner Department: {prompt.department || 'Not specified'}</Text>
                      </div>
                      <div className={styles.promptMetaItem}>
                        <Text>Created Date: {formatDate(prompt.createdDate)}</Text>
                      </div>
                    </div>

                    <div className={styles.promptTags}>
                      {prompt.tags.map((tag) => (
                        <Badge key={tag} appearance="outline">{tag}</Badge>
                      ))}
                    </div>

                    <div className={styles.promptCardFooter}>
                      <div className={styles.promptFooterLeft}>
                        <Caption1>Author: {prompt.createdBy}</Caption1>
                        <Caption1><ClockRegular /> {formatDate(prompt.createdDate)}</Caption1>
                      </div>
                      <div className={styles.promptFooterActions}>
                        <Button appearance="secondary" icon={<CopyRegular />} onClick={() => copyPrompt(prompt.id)}>Copy Prompt</Button>
                        {isMyPromptsView && <Button appearance="secondary" onClick={() => editPrompt(prompt.id)}>Edit</Button>}
                        {isMyPromptsView && <Button appearance="secondary" icon={<DeleteRegular />} onClick={() => deletePrompt(prompt.id)}>Delete</Button>}
                        <Button appearance="primary" icon={<EyeRegular />} onClick={() => viewPrompt(prompt.id)}>View Details</Button>
                      </div>
                    </div>
                  </Card>
                ))
              )}
            </section>
          </main>

          <aside className={styles.rightPanel}>
            <Card className={styles.sidePanelCard}>
              <CardHeader
                header={<Body1 style={{ fontWeight: 600 }}>{editingPromptId ? 'Edit Prompt' : 'Add Prompt'}</Body1>}
                description="Save directly to the SharePoint Prompt Library"
              />

              <div className={styles.formStack}>
                <Field label="Title">
                  <Input className={styles.whiteInput} value={promptForm.title} onChange={(_, data) => updatePromptForm('title', data.value)} placeholder="Enter prompt title" />
                </Field>

                <Field label="Category">
                  <Dropdown
                    className={styles.whiteDropdown}
                    placeholder="Select category"
                    value={promptForm.category}
                    selectedOptions={promptForm.category ? [promptForm.category] : []}
                    onOptionSelect={(_, data) => updatePromptForm('category', data.optionValue || '')}
                  >
                    {categoryOptions.map((option) => (
                      <Option key={option} value={option}>{option}</Option>
                    ))}
                  </Dropdown>
                </Field>

                <Field label="Recommended AI Tool (optional)" hint="Suggest a tool when it is especially effective for this prompt.">
                  <Dropdown
                    className={styles.whiteDropdown}
                    placeholder="Select a recommended tool"
                    value={promptForm.aiModel}
                    selectedOptions={promptForm.aiModel ? [promptForm.aiModel] : []}
                    onOptionSelect={(_, data) => updatePromptForm('aiModel', data.optionValue || '')}
                  >
                    {modelOptions.map((option) => (
                      <Option key={option} value={option}>{option}</Option>
                    ))}
                  </Dropdown>
                </Field>

                <Field label="Description">
                  <Textarea className={styles.whiteTextarea} value={promptForm.description} onChange={(_, data) => updatePromptForm('description', data.value)} resize="vertical" placeholder="Brief prompt purpose and when to use it" />
                </Field>

                <Field label="Prompt Text">
                  <Textarea className={styles.whiteTextarea} value={promptForm.promptText} onChange={(_, data) => updatePromptForm('promptText', data.value)} resize="vertical" placeholder="Paste the actual prompt content here" />
                </Field>

                <Field label="Tags">
                  <Dropdown
                    className={styles.whiteDropdown}
                    multiselect
                    placeholder="Select tags"
                    selectedOptions={promptForm.tags}
                    onOptionSelect={(_, data) => updatePromptForm('tags', data.selectedOptions)}
                  >
                    {tagOptions.map((option) => (
                      <Option key={option} value={option}>{option}</Option>
                    ))}
                  </Dropdown>
                </Field>

                <Field label="Owner Department">
                  <Dropdown
                    className={styles.whiteDropdown}
                    placeholder="Select owner department"
                    value={promptForm.department || ''}
                    selectedOptions={promptForm.department ? [promptForm.department] : []}
                    onOptionSelect={(_, data) => updatePromptForm('department', data.optionValue || '')}
                  >
                    {departmentOptions.map((option) => (
                      <Option key={option} value={option}>{option}</Option>
                    ))}
                  </Dropdown>
                </Field>

                <Field label="Visibility">
                  <Dropdown
                    className={styles.whiteDropdown}
                    placeholder="Select visibility"
                    value={promptForm.visibility || ''}
                    selectedOptions={promptForm.visibility ? [promptForm.visibility] : []}
                    onOptionSelect={(_, data) => updatePromptForm('visibility', data.optionValue || '')}
                  >
                    {visibilityOptions.map((option) => (
                      <Option key={option} value={option}>{option}</Option>
                    ))}
                  </Dropdown>
                </Field>

                <Button
                  appearance={promptForm.featured ? 'primary' : 'secondary'}
                  onClick={() => updatePromptForm('featured', !promptForm.featured)}
                >
                  {promptForm.featured ? 'Featured Prompt' : 'Mark as Featured'}
                </Button>
              </div>

              <Divider />

              <div className={styles.formActions}>
                <Button appearance="primary" icon={<CheckmarkCircleRegular />} disabled={isSaving} onClick={() => savePrompt(editingPromptId ? promptForm.status : 'Published')}>
                  {editingPromptId ? 'Save Edit' : 'Publish'}
                </Button>
                {!editingPromptId && <Button appearance="secondary" disabled={isSaving} onClick={() => savePrompt('Draft')}>Save Draft</Button>}
                <Button appearance="subtle" disabled={isSaving} onClick={resetPromptForm}>Cancel</Button>
              </div>
            </Card>

            {selectedPrompt && (
              <div className={styles.promptPreviewOverlay} role="presentation" onMouseDown={() => setSelectedPrompt(null)}>
                <section
                  className={styles.promptPreviewDialog}
                  role="dialog"
                  aria-modal="true"
                  aria-labelledby="prompt-preview-title"
                  onMouseDown={(event) => event.stopPropagation()}
                >
                  <div className={styles.promptPreviewHeader}>
                    <Title2 id="prompt-preview-title">{selectedPrompt.title}</Title2>
                    <Button appearance="subtle" onClick={() => setSelectedPrompt(null)}>Close</Button>
                  </div>
                  <div className={styles.promptPreviewContent}>
                    <div className={styles.promptPreviewBadges}>
                      <Badge appearance={selectedPrompt.status === 'Published' ? 'filled' : 'outline'}>{selectedPrompt.status}</Badge>
                      {selectedPrompt.featured && <Badge appearance="filled" color="brand">Featured</Badge>}
                      {selectedPrompt.tags.map((tag) => <Badge key={tag} appearance="outline">{tag}</Badge>)}
                    </div>
                    <Body2>{selectedPrompt.description}</Body2>
                    <div className={styles.promptPreviewMetaGrid}>
                      <Text>Category: {selectedPrompt.category}</Text>
                      <Text>Owner Department: {selectedPrompt.department || 'Not specified'}</Text>
                      <Text>Recommended AI Tool: {selectedPrompt.aiModel || 'Not specified'}</Text>
                      <Text>Visibility: {selectedPrompt.visibility || 'Organization'}</Text>
                      <Text>Author: {selectedPrompt.createdBy}</Text>
                      <Text>Created: {formatDate(selectedPrompt.createdDate)}</Text>
                      <Text>Last modified: {formatDate(selectedPrompt.modifiedDate)}</Text>
                    </div>
                    <div>
                      <Caption1>Prompt text</Caption1>
                      <Textarea className={`${styles.whiteTextarea} ${styles.promptPreviewTextarea}`} value={selectedPrompt.promptText} resize="vertical" readOnly />
                    </div>
                    {selectedPrompt.documentUrl && <Caption1>Sample asset: {selectedPrompt.documentUrl}</Caption1>}
                  </div>
                  <div className={styles.promptPreviewActions}>
                    <Button appearance="secondary" onClick={() => copyPrompt(selectedPrompt.id)}>Copy Prompt</Button>
                  </div>
                </section>
              </div>
            )}

          </aside>
        </div>
      </div>
    </section>
  );
}
