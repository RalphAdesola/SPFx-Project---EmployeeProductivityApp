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
  BoxRegular,
  ChevronDoubleLeftRegular,
  ChevronDoubleRightRegular,
  ClockRegular,
  CopyRegular,
  EyeRegular,
  FilterRegular,
  HeartRegular,
  HomeRegular,
  LightbulbRegular,
  MegaphoneRegular,
  CodeRegular,
  PersonRegular,
  SearchRegular,
  SettingsRegular,
  StarRegular,
  TagRegular,
  BoardRegular,
  HistoryRegular,
  LayerRegular,
  SparkleRegular,
  MoreHorizontalRegular,
  CheckmarkCircleRegular,
  SignOutRegular
} from '@fluentui/react-icons';
import styles from './EmployeeProductivityApp.module.scss';
import type {
  IAiModelSummary,
  IDashboardMetric,
  IDepartmentSummary,
  ICategorySummary,
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
  searchTerm: string;
  onSearchTermChange: (value: string) => void;
  onBackToLanding: () => void;
  isLoadingData?: boolean;
  loadError?: string | null;
  onRefresh?: () => void;
  onFiltersChange: (filters: IPromptFilters) => void;
  onCreatePrompt: (payload: IPromptWritePayload, file?: File) => Promise<void>;
  onUpdatePrompt: (id: number, payload: IPromptWritePayload, file?: File) => Promise<void>;
  onViewPrompt: (id: number) => Promise<IPromptDetails | null>;
  onDeletePrompt: (id: number) => Promise<void>;
  onCopyPrompt: (id: number) => Promise<string>;
  onSubmitRating: (id: number, rating: number) => Promise<void>;
  onNavigateToAdmin: () => void;
  canAccessAdmin: boolean;
}

const useStyles = makeStyles({
  pageSurface: {
    backgroundColor: '#f8fafc'
  }
});

const navigationItems = [
  { key: 'dashboard', label: 'Dashboard', icon: <HomeRegular /> },
  { key: 'my-prompts', label: 'My Prompts', icon: <PersonRegular /> },
  { key: 'shared-library', label: 'Shared Prompt Library', icon: <BoardRegular /> },
  { key: 'categories', label: 'Categories', icon: <LayerRegular /> },
  { key: 'favorites', label: 'Favorites', icon: <HeartRegular /> },
  { key: 'recent', label: 'Recently Used', icon: <HistoryRegular /> },
  { key: 'models', label: 'AI Models', icon: <CodeRegular /> },
  { key: 'tags', label: 'Tags', icon: <TagRegular /> },
  { key: 'settings', label: 'Settings', icon: <SettingsRegular /> }
];

const visibilityOptions = ['Organization', 'Department', 'Private'];
const statusOptions: Array<'Published' | 'Draft'> = ['Published', 'Draft'];
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
  const [activeNavItem, setActiveNavItem] = React.useState('dashboard');
  const [localSearch, setLocalSearch] = React.useState(props.searchTerm);
  const [selectedCategory, setSelectedCategory] = React.useState<string | undefined>(undefined);
  const [selectedModel, setSelectedModel] = React.useState<string | undefined>(undefined);
  const [selectedDepartment, setSelectedDepartment] = React.useState<string | undefined>(undefined);
  const [selectedTag, setSelectedTag] = React.useState<string | undefined>(undefined);
  const [selectedStatus, setSelectedStatus] = React.useState<'Published' | 'Draft' | undefined>(undefined);
  const [promptForm, setPromptForm] = React.useState<IPromptWritePayload>(emptyPromptForm);
  const [editingPromptId, setEditingPromptId] = React.useState<number | undefined>(undefined);
  const [selectedPrompt, setSelectedPrompt] = React.useState<IPromptDetails | null>(null);
  const [attachmentFile, setAttachmentFile] = React.useState<File | undefined>(undefined);
  const [isSaving, setIsSaving] = React.useState(false);
  const [actionError, setActionError] = React.useState<string | null>(null);
  const [actionMessage, setActionMessage] = React.useState<string | null>(null);
  const hasActiveFilters = Boolean(localSearch.trim() || selectedCategory || selectedModel || selectedDepartment || selectedTag || selectedStatus);
  const categoryOptions = props.categories.map((category) => category.title);
  const modelOptions = props.models.map((model) => model.title);
  const departmentOptions = props.departments.map((department) => department.title);
  const tagOptions = props.tags.map((tag) => tag.title);
  const visibleNavigationItems = props.canAccessAdmin
    ? [...navigationItems, { key: 'admin', label: 'Administration', icon: <SparkleRegular /> }]
    : navigationItems;

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
    setAttachmentFile(undefined);
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

      if (!payload.title.trim() || !payload.category || !payload.aiModel || !payload.promptText.trim()) {
        setActionError('Title, Category, AI Model, and Prompt Text are required.');
        return;
      }

      if (editingPromptId) {
        await props.onUpdatePrompt(editingPromptId, payload, attachmentFile);
        setActionMessage('Prompt updated successfully.');
      } else {
        await props.onCreatePrompt(payload, attachmentFile);
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
    setSelectedPrompt(prompt);
    setActionMessage('Editing latest SharePoint prompt record.');
  };

  const copyPrompt = async (promptId: string): Promise<void> => {
    setActionError(null);

    try {
      const promptText = await props.onCopyPrompt(Number(promptId));

      if (navigator.clipboard) {
        await navigator.clipboard.writeText(promptText);
      }

      setActionMessage('Prompt copied and usage logged.');
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
              }}
              aria-current={activeNavItem === item.key ? 'page' : undefined}
            >
              <span className={styles.sidebarNavIcon}>{item.icon}</span>
              {!sidebarCollapsed && <span>{item.label}</span>}
            </button>
          ))}
        </nav>

        <div className={styles.sidebarFooter}>
          {!sidebarCollapsed ? (
            <div className={styles.sidebarFooterStack}>
              <Card className={styles.sidebarFooterCard}>
                <Body2>SharePoint-ready architecture</Body2>
                <Caption1>Built for curated prompts, governance, and collaboration.</Caption1>
              </Card>
              <Button
                appearance="subtle"
                className={`${styles.sidebarNavItem} ${styles.logoutNavItem}`}
                onClick={props.onBackToLanding}
                icon={<SignOutRegular />}
              >
                <span>Log Out</span>
              </Button>
            </div>
          ) : (
            <div className={styles.sidebarFooterCompactStack}>
              <div className={styles.sidebarFooterCompact}>
                <SparkleRegular />
              </div>
              <Button
                appearance="subtle"
                className={`${styles.sidebarNavItem} ${styles.logoutNavItem} ${styles.logoutNavItemCompact}`}
                onClick={props.onBackToLanding}
                icon={<SignOutRegular />}
                aria-label="Log out"
              />
            </div>
          )}
        </div>
      </aside>

      <div className={styles.dashboardMain}>
        <header className={styles.topHeader}>
          <div className={styles.topHeaderLeft}>
            <Title1 className={styles.pageTitle}>Dashboard</Title1>
            <Caption1>Welcome back, {props.displayName || 'colleague'}</Caption1>
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

            <Button appearance="subtle" icon={<FilterRegular />}>Filter</Button>
            <Button appearance="subtle" icon={<MoreHorizontalRegular />}>Notifications</Button>
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

            <section className={styles.summaryGrid}>
              {props.metrics.map((metric) => (
                <Card key={metric.label} className={styles.metricCard}>
                  <CardHeader header={<Body2>{metric.label}</Body2>} />
                  <Title2>{metric.value}</Title2>
                  <Caption1>{metric.detail}</Caption1>
                </Card>
              ))}
            </section>

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
                <Field label="AI Model">
                  <Dropdown
                    className={styles.whiteDropdown}
                    placeholder="All models"
                    value={selectedModel || ''}
                    selectedOptions={selectedModel ? [selectedModel] : []}
                    onOptionSelect={(_, data) => setSelectedModel(data.optionValue ?? undefined)}
                  >
                    {modelOptions.map((option) => (
                      <Option key={option} value={option}>{option}</Option>
                    ))}
                  </Dropdown>
                </Field>
              </div>

              <div className={styles.filterField}>
                <Field label="Status">
                  <Dropdown
                    className={styles.whiteDropdown}
                    placeholder="Published / Draft"
                    value={selectedStatus || ''}
                    selectedOptions={selectedStatus ? [selectedStatus] : []}
                    onOptionSelect={(_, data) => setSelectedStatus((data.optionValue as 'Published' | 'Draft') || undefined)}
                  >
                    {statusOptions.map((option) => (
                      <Option key={option} value={option}>{option}</Option>
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
              {props.prompts.length === 0 ? (
                <Card className={styles.emptyStateCard}>
                  <Body1>No prompts match your current filters.</Body1>
                  <Caption1>Try a different search term, category, or model.</Caption1>
                </Card>
              ) : (
                props.prompts.map((prompt) => (
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
                      <Button appearance="subtle" icon={<EyeRegular />} onClick={() => viewPrompt(prompt.id)}>View Details</Button>
                    </div>

                    <Body2 className={styles.promptDescription}>{prompt.description}</Body2>

                    <div className={styles.promptMetaGrid}>
                      <div className={styles.promptMetaItem}>
                        <Caption1>AI Model</Caption1>
                        <Text>{prompt.aiModel}</Text>
                      </div>
                      <div className={styles.promptMetaItem}>
                        <Caption1>Created Date</Caption1>
                        <Text>{prompt.createdDate}</Text>
                      </div>
                      <div className={styles.promptMetaItem}>
                        <Caption1>Average Rating</Caption1>
                        <Text>{prompt.averageRating.toFixed(1)} / 5</Text>
                      </div>
                      <div className={styles.promptMetaItem}>
                        <Caption1>Usage Count</Caption1>
                        <Text>{prompt.usageCount}</Text>
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
                        <Caption1><ClockRegular /> {prompt.createdDate}</Caption1>
                      </div>
                      <div className={styles.promptFooterActions}>
                        <Button appearance="secondary" icon={<CopyRegular />} onClick={() => copyPrompt(prompt.id)}>Copy Prompt</Button>
                        <Button appearance="secondary" onClick={() => editPrompt(prompt.id)}>Edit</Button>
                        <Button appearance="secondary" onClick={() => deletePrompt(prompt.id)}>Delete</Button>
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

                <Field label="AI Model">
                  <Dropdown
                    className={styles.whiteDropdown}
                    placeholder="Select model"
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
                  <Input className={styles.whiteInput} value={promptForm.tags.join(', ')} onChange={(_, data) => updatePromptForm('tags', data.value.split(','))} placeholder="comma-separated tags" />
                </Field>

                <Field label="Department">
                  <Dropdown
                    className={styles.whiteDropdown}
                    placeholder="Select department"
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

                <Field label="Prompt Asset">
                  <input
                    className={styles.nativeFileInput}
                    type="file"
                    onChange={(event) => setAttachmentFile(event.currentTarget.files?.[0])}
                  />
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
                <Button appearance="primary" icon={<CheckmarkCircleRegular />} disabled={isSaving} onClick={() => savePrompt('Published')}>Publish</Button>
                <Button appearance="secondary" disabled={isSaving} onClick={() => savePrompt('Draft')}>Save Draft</Button>
                <Button appearance="subtle" disabled={isSaving} onClick={resetPromptForm}>Cancel</Button>
              </div>
            </Card>

            {selectedPrompt && (
              <Card className={styles.sidePanelCard}>
                <CardHeader
                  header={<Body1 style={{ fontWeight: 600 }}>{selectedPrompt.title}</Body1>}
                  description={`${selectedPrompt.category} · ${selectedPrompt.aiModel}`}
                />

                <div className={styles.formStack}>
                  <Body2>{selectedPrompt.description}</Body2>
                  <Textarea className={styles.whiteTextarea} value={selectedPrompt.promptText} resize="vertical" readOnly />
                  <Caption1>Author: {selectedPrompt.createdBy}</Caption1>
                  <Caption1>Modified: {selectedPrompt.modifiedDate}</Caption1>
                  <div className={styles.formActions}>
                    {[1, 2, 3, 4, 5].map((rating) => (
                      <Button key={rating} appearance="secondary" icon={<StarRegular />} onClick={() => props.onSubmitRating(Number(selectedPrompt.id), rating)}>
                        {rating}
                      </Button>
                    ))}
                  </div>
                  <div className={styles.formActions}>
                    <Button appearance="secondary" onClick={() => editPrompt(selectedPrompt.id)}>Edit</Button>
                    <Button appearance="secondary" onClick={() => copyPrompt(selectedPrompt.id)}>Copy</Button>
                    <Button appearance="subtle" onClick={() => setSelectedPrompt(null)}>Close</Button>
                  </div>
                </div>
              </Card>
            )}

            <Card className={styles.sidePanelCard}>
              <CardHeader
                header={<Body1 style={{ fontWeight: 600 }}>Quick Filters</Body1>}
                description="Useful shortcuts for future SharePoint list bindings"
              />

              <div className={styles.quickFilterStack}>
                <div className={styles.quickFilterChip}><BoxRegular /> My Prompts</div>
                <div className={styles.quickFilterChip}><StarRegular /> Favorites</div>
                <div className={styles.quickFilterChip}><MegaphoneRegular /> Shared Library</div>
                <div className={styles.quickFilterChip}><LightbulbRegular /> Recently Used</div>
              </div>
            </Card>
          </aside>
        </div>
      </div>
    </section>
  );
}
