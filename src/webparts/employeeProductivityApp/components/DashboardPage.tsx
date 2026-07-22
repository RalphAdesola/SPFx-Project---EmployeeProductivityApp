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
import type { IDashboardMetric, IPromptCategory, IPromptSummary } from './SharedTypes';

export interface IDashboardPageProps {
  displayName: string;
  metrics: IDashboardMetric[];
  categories: IPromptCategory[];
  prompts: IPromptSummary[];
  searchTerm: string;
  onSearchTermChange: (value: string) => void;
  onBackToLanding: () => void;
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

const categoryOptions = ['Sales Enablement', 'HR & People Ops', 'Operations', 'IT Support'];
const modelOptions = ['GPT-4.1', 'GPT-4o', 'Claude 3.5 Sonnet', 'Gemini 2.5 Pro'];
const departmentOptions = ['Sales', 'Human Resources', 'Operations', 'IT', 'Finance', 'Marketing'];
const visibilityOptions = ['Organization', 'Department', 'Private'];

export default function DashboardPage(props: IDashboardPageProps): React.ReactElement {
  const pageStyles = useStyles();
  const [sidebarCollapsed, setSidebarCollapsed] = React.useState(false);
  const [activeNavItem, setActiveNavItem] = React.useState('dashboard');
  const [localSearch, setLocalSearch] = React.useState(props.searchTerm);
  const [selectedCategory, setSelectedCategory] = React.useState<string | undefined>(undefined);
  const [selectedModel, setSelectedModel] = React.useState<string | undefined>(undefined);
  const hasActiveFilters = Boolean(localSearch.trim() || selectedCategory || selectedModel);

  React.useEffect(() => {
    props.onSearchTermChange(localSearch);
  }, [localSearch, props.onSearchTermChange]);

  const filteredPrompts = React.useMemo(() => {
    const normalized = localSearch.trim().toLowerCase();

    return props.prompts.filter((prompt) => {
      const searchableText = [
        prompt.title,
        prompt.description,
        prompt.category,
        prompt.aiModel,
        prompt.tags.join(' '),
        prompt.status
      ]
        .join(' ')
        .toLowerCase();

      const matchesText = !normalized || searchableText.includes(normalized);
      const matchesCategory = !selectedCategory || prompt.category === selectedCategory;
      const matchesModel = !selectedModel || prompt.aiModel === selectedModel;

      return matchesText && matchesCategory && matchesModel;
    });
  }, [localSearch, props.prompts, selectedCategory, selectedModel]);

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
          {navigationItems.map((item) => (
            <button
              key={item.key}
              className={`${styles.sidebarNavItem} ${activeNavItem === item.key ? styles.sidebarNavItemActive : ''}`}
              type="button"
              onClick={() => setActiveNavItem(item.key)}
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
            <Button appearance="primary" icon={<AddRegular />}>Add Prompt</Button>
          </div>
        </header>

        <div className={styles.dashboardLayout}>
          <main className={styles.mainContent}>
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
                  <Input className={styles.whiteInput} placeholder="Published / Draft" readOnly />
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
                }}
                disabled={!hasActiveFilters}
              >
                Clear Filters
              </Button>
            </div>

            <section className={styles.promptGrid}>
              {filteredPrompts.length === 0 ? (
                <Card className={styles.emptyStateCard}>
                  <Body1>No prompts match your current filters.</Body1>
                  <Caption1>Try a different search term, category, or model.</Caption1>
                </Card>
              ) : (
                filteredPrompts.map((prompt) => (
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
                      <Button appearance="subtle" icon={<EyeRegular />}>View Details</Button>
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
                        <Button appearance="secondary" icon={<CopyRegular />}>Copy Prompt</Button>
                        <Button appearance="primary" icon={<EyeRegular />}>View Details</Button>
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
                header={<Body1 style={{ fontWeight: 600 }}>Add Prompt</Body1>}
                description="Prepare a SharePoint-ready prompt record"
              />

              <div className={styles.formStack}>
                <Field label="Title">
                  <Input className={styles.whiteInput} placeholder="Enter prompt title" />
                </Field>

                <Field label="Category">
                  <Dropdown className={styles.whiteDropdown} placeholder="Select category">
                    {categoryOptions.map((option) => (
                      <Option key={option} value={option}>{option}</Option>
                    ))}
                  </Dropdown>
                </Field>

                <Field label="AI Model">
                  <Dropdown className={styles.whiteDropdown} placeholder="Select model">
                    {modelOptions.map((option) => (
                      <Option key={option} value={option}>{option}</Option>
                    ))}
                  </Dropdown>
                </Field>

                <Field label="Description">
                  <Textarea className={styles.whiteTextarea} resize="vertical" placeholder="Brief prompt purpose and when to use it" />
                </Field>

                <Field label="Prompt Text">
                  <Textarea className={styles.whiteTextarea} resize="vertical" placeholder="Paste the actual prompt content here" />
                </Field>

                <Field label="Tags">
                  <Input className={styles.whiteInput} placeholder="comma-separated tags" />
                </Field>

                <Field label="Department">
                  <Dropdown className={styles.whiteDropdown} placeholder="Select department">
                    {departmentOptions.map((option) => (
                      <Option key={option} value={option}>{option}</Option>
                    ))}
                  </Dropdown>
                </Field>

                <Field label="Visibility">
                  <Dropdown className={styles.whiteDropdown} placeholder="Select visibility">
                    {visibilityOptions.map((option) => (
                      <Option key={option} value={option}>{option}</Option>
                    ))}
                  </Dropdown>
                </Field>
              </div>

              <Divider />

              <div className={styles.formActions}>
                <Button appearance="primary" icon={<CheckmarkCircleRegular />}>Publish</Button>
                <Button appearance="secondary">Save Draft</Button>
                <Button appearance="subtle">Cancel</Button>
              </div>
            </Card>

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
