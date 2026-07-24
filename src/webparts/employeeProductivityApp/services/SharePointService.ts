import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { spfi, type SPFI } from '@pnp/sp';
import { SPFx } from '@pnp/sp/behaviors/spfx';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';

import type {
  IActivityLogEntry,
  IAiModelSummary,
  IAdminSummary,
  ICategorySummary,
  IDirectoryUser,
  IDepartmentSummary,
  IPromptDetails,
  IPromptFilters,
  IPromptSummary,
  ISeedReport,
  ITagSummary,
  IUserSummary
} from '../components/SharedTypes';

type ItemShape = Record<string, unknown> & { Id: number; Title?: string };
type PersonShape = { Title?: string; Email?: string; EMail?: string; Id?: number };
type LookupShape = { Id?: number; Title?: string };
type ItemAddShape = { data?: { Id?: number }; Id?: number };

export interface ISharePointPromptWritePayload {
  title: string;
  description: string;
  promptText: string;
  categoryId: number;
  aiModelId?: number;
  departmentId?: number;
  tagIds: number[];
  status: string;
  visibility: string;
  featured: boolean;
  usageCount?: number;
  documentUrl?: string;
  promptOwnerId?: number;
  lastReviewed?: string;
}

export interface ISeedDemoCallbacks {
  onProgress?: (entry: string) => void;
}

const LISTS = {
  prompts: 'Prompt Library',
  categories: 'Prompt Categories',
  departments: 'Departments',
  models: 'AI Models',
  tags: 'Tags',
  users: 'Users',
  admins: 'App Administrator',
  activities: 'Activity Log'
} as const;

const DEMO_PREFIX = 'DEMO - ';

type SeedProgressKey =
  | 'categories'
  | 'models'
  | 'departments'
  | 'users'
  | 'tags'
  | 'prompts'
  | 'activities';

const createProgressEntry = (label: string) => ({
  label,
  created: 0,
  skipped: 0,
  failed: 0,
  status: 'pending' as const
});

class SharePointService {
  private static instance: SharePointService | undefined;
  private sp: SPFI | undefined;
  private context: WebPartContext | undefined;

  public static initialize(context: WebPartContext): SharePointService {
    if (!SharePointService.instance) {
      SharePointService.instance = new SharePointService();
    }

    SharePointService.instance.sp = spfi().using(SPFx(context));
    SharePointService.instance.context = context;
    return SharePointService.instance;
  }

  public static getInstance(): SharePointService {
    if (!SharePointService.instance?.sp) {
      throw new Error('SharePointService is not initialized.');
    }

    return SharePointService.instance;
  }

  private get web() {
    if (!this.sp) {
      throw new Error('SharePointService is not initialized.');
    }

    return this.sp.web;
  }

  private toText(value: unknown): string {
    if (value === null || value === undefined) {
      return '';
    }

    if (typeof value === 'object') {
      const person = value as PersonShape;
      return person.Title || person.EMail || person.Email || '';
    }

    return String(value);
  }

  private toLookup(value: unknown): LookupShape | undefined {
    if (!value || typeof value !== 'object') {
      return undefined;
    }

    return value as LookupShape;
  }

  private toPerson(value: unknown): PersonShape | undefined {
    if (!value || typeof value !== 'object') {
      return undefined;
    }

    return value as PersonShape;
  }

  private async getGraphClient(): Promise<MSGraphClientV3> {
    if (!this.context) {
      throw new Error('SharePointService is not initialized.');
    }

    return this.context.msGraphClientFactory.getClient('3');
  }

  private toPeople(value: unknown): PersonShape[] {
    if (Array.isArray(value)) {
      return value.map((entry) => this.toPerson(entry)).filter((entry): entry is PersonShape => Boolean(entry));
    }

    if (value && typeof value === 'object' && Array.isArray((value as { results?: unknown[] }).results)) {
      return this.toPeople((value as { results: unknown[] }).results);
    }

    const person = this.toPerson(value);
    return person ? [person] : [];
  }

  private normalizeTags(value: unknown): string[] {
    if (Array.isArray(value)) {
      return value.map((entry) => this.toText(entry).trim()).filter(Boolean);
    }

    if (typeof value === 'string') {
      return value.split(/[;,]/).map((entry) => entry.trim()).filter(Boolean);
    }

    return [];
  }

  private normalizeLookupIds(value: unknown): number[] {
    if (Array.isArray(value)) {
      return value
        .map((entry) => this.toLookup(entry)?.Id)
        .filter((entry): entry is number => typeof entry === 'number');
    }

    const lookup = this.toLookup(value);
    return lookup?.Id ? [lookup.Id] : [];
  }

  private mapPrompt(item: ItemShape): IPromptSummary {
    const status = this.toText(item.Status) === 'Draft' ? 'Draft' : 'Published';
    const category = this.toLookup(item.Category);
    const aiModel = this.toLookup(item.AIModel);
    const department = this.toLookup(item.Department);
    const owner = this.toLookup(item.PromptOwner) || this.toLookup(item['PromptOwner']);

    return {
      id: String(item.Id),
      title: this.toText(item.Title),
      category: category?.Title || '',
      categoryId: category?.Id,
      aiModel: aiModel?.Title || '',
      aiModelId: aiModel?.Id,
      description: this.toText(item.Description),
      tags: this.normalizeTags(item.Tags),
      tagIds: this.normalizeLookupIds(item.Tags),
      featured: Boolean(item.Featured),
      status,
      createdBy: this.toText(owner || item.Author || item.CreatedBy),
      createdById: owner?.Id,
      createdDate: this.toText(item.Created || item.Modified),
      usageCount: Number(item.UsageCount || 0),
      documentUrl: item.DocumentUrl ? this.toText(item.DocumentUrl) : undefined,
      department: department?.Title,
      departmentId: department?.Id,
      isDeleted: Boolean(item.IsDeleted)
    };
  }

  private mapPromptDetails(item: ItemShape): IPromptDetails {
    return {
      ...this.mapPrompt(item),
      promptText: this.toText(item.PromptText),
      modifiedDate: this.toText(item.Modified || item.Created),
      visibility: item.Visibility ? this.toText(item.Visibility) : undefined
    };
  }

  private mapCategory(item: ItemShape): ICategorySummary {
    return {
      id: item.Id,
      title: this.toText(item.Title),
      description: item.Description ? this.toText(item.Description) : undefined,
      promptCount: item.PromptCount ? Number(item.PromptCount) : undefined
    };
  }

  private mapLookup(item: ItemShape): { id: number; title: string } {
    return {
      id: item.Id,
      title: this.toText(item.Title)
    };
  }

  private mapUser(item: ItemShape): IUserSummary {
    return {
      id: item.Id,
      title: this.toText(item.Title),
      email: item.EMail ? this.toText(item.EMail) : item.Email ? this.toText(item.Email) : undefined,
      department: item.Department ? this.toText(item.Department) : undefined
    };
  }

  private mapActivity(item: ItemShape): IActivityLogEntry {
    return {
      id: item.Id,
      action: this.toText(item.ActivityType),
      promptId: item.PromptId ? Number(item.PromptId) : undefined,
      promptTitle: item.PromptTitle ? this.toText(item.PromptTitle) : undefined,
      actor: this.toText(this.toPerson(item.Actor) || item.Author),
      details: item.Details ? this.toText(item.Details) : undefined,
      createdDate: this.toText(item.Created || item.Modified)
    };
  }

  public async isCurrentUserAdmin(): Promise<boolean> {
    const currentUser = await this.web.currentUser();
    const currentEmail = (currentUser.Email || (currentUser as { EMail?: string }).EMail || '').trim().toLowerCase();
    const currentUserId = Number(currentUser.Id || 0);

    if (!currentEmail && !currentUserId) {
      return false;
    }

    const items = await this.web.lists.getByTitle(LISTS.admins).items
      .select('Id,UserId,User/Id,User/Title,User/EMail')
      .expand('User')
      .top(5000)();

    return (items as ItemShape[]).some((item) => {
      return this.toPeople(item.User).some((adminUser) => {
        const adminEmail = (adminUser.EMail || adminUser.Email || '').trim().toLowerCase();
        return (Boolean(currentEmail) && adminEmail === currentEmail)
          || (Boolean(currentUserId) && Number(adminUser.Id || 0) === currentUserId);
      }) || Number(item.UserId || 0) === currentUserId;
    });
  }

  public async getAdmins(): Promise<IAdminSummary[]> {
    const items = await this.web.lists.getByTitle(LISTS.admins).items
      .select('Id,UserId,User/Id,User/Title,User/EMail')
      .expand('User')
      .top(5000)();

    return (items as ItemShape[])
      .flatMap((item) => this.toPeople(item.User).map((adminUser) => ({
        listItemId: item.Id,
        userId: Number(adminUser.Id || item.UserId || 0),
        title: adminUser.Title || adminUser.EMail || adminUser.Email || 'Administrator',
        email: adminUser.EMail || adminUser.Email
      })));
  }

  public async getCurrentUserId(): Promise<number> {
    const currentUser = await this.web.currentUser();
    const currentUserId = Number(currentUser.Id || 0);

    if (!currentUserId) {
      throw new Error('Unable to determine the current SharePoint user.');
    }

    return currentUserId;
  }

  public async addAdmin(email: string): Promise<void> {
    if (!await this.isCurrentUserAdmin()) {
      throw new Error('Only application administrators can add another administrator.');
    }

    const normalizedEmail = email.trim();
    if (!normalizedEmail) {
      throw new Error('Enter an administrator email address.');
    }

    const siteUser = await this.web.ensureUser(normalizedEmail);
    const existingItems = await this.web.lists.getByTitle(LISTS.admins).items
      .select('Id,UserId')
      .filter(`UserId eq ${siteUser.Id}`)
      .top(5000)();
    const existing = (existingItems as ItemShape[]);

    if (existing.length > 0) {
      throw new Error('This user has already been configured as an administrator.');
    }

    await this.web.lists.getByTitle(LISTS.admins).items.add({
      Title: normalizedEmail,
      UserId: siteUser.Id
    });
  }

  public async removeAdmin(listItemId: number): Promise<void> {
    if (!await this.isCurrentUserAdmin()) {
      throw new Error('Only application administrators can remove another administrator.');
    }

    const currentUserId = await this.getCurrentUserId();
    const adminItem = await this.web.lists.getByTitle(LISTS.admins).items
      .getById(listItemId)
      .select('Id,UserId')();

    if (Number((adminItem as ItemShape).UserId || 0) === currentUserId) {
      throw new Error('You cannot remove your own administrator access.');
    }

    await this.web.lists.getByTitle(LISTS.admins).items.getById(listItemId).delete();
  }

  public async searchDirectoryUsers(searchText: string): Promise<IDirectoryUser[]> {
    const normalizedSearch = searchText.trim();
    if (normalizedSearch.length < 2) {
      return [];
    }

    const escapedSearch = normalizedSearch.replace(/'/g, "''");
    const graphClient = await this.getGraphClient();
    const response = await graphClient
      .api('/users')
      .version('v1.0')
      .select('id,displayName,mail,userPrincipalName')
      .filter(`startsWith(displayName,'${escapedSearch}') or startsWith(mail,'${escapedSearch}') or startsWith(userPrincipalName,'${escapedSearch}')`)
      .top(10)
      .get() as { value?: Array<{ id?: string; displayName?: string; mail?: string; userPrincipalName?: string }> };

    return (response.value || [])
      .filter((user) => Boolean(user.id && user.displayName && (user.mail || user.userPrincipalName)))
      .map((user) => ({
        id: user.id as string,
        displayName: user.displayName as string,
        email: (user.mail || user.userPrincipalName) as string
      }));
  }

  public async hasCurrentUserDraft(excludePromptId?: number): Promise<boolean> {
    const currentUserId = await this.getCurrentUserId();

    const items = await this.web.lists.getByTitle(LISTS.prompts).items
      .select('Id,Status,IsDeleted,PromptOwnerId')
      .filter(`PromptOwnerId eq ${currentUserId} and Status eq 'Draft'`)
      .top(2)();

    return (items as ItemShape[]).some((item) => Number(item.Id) !== excludePromptId && !Boolean(item.IsDeleted));
  }

  private applyPromptFilters(prompts: IPromptSummary[], filters?: IPromptFilters): IPromptSummary[] {
    if (!filters) {
      return prompts;
    }

    const normalizedSearch = (filters.searchTerm || '').trim().toLowerCase();

    return prompts.filter((prompt) => {
      const searchableText = [
        prompt.title,
        prompt.description,
        prompt.category,
        prompt.aiModel,
        prompt.department,
        prompt.status,
        prompt.tags.join(' ')
      ].join(' ').toLowerCase();

      return (!normalizedSearch || searchableText.indexOf(normalizedSearch) >= 0)
        && (!filters.category || prompt.category === filters.category)
        && (!filters.aiModel || prompt.aiModel === filters.aiModel)
        && (!filters.department || prompt.department === filters.department)
        && (!filters.status || prompt.status === filters.status)
        && (!filters.tag || prompt.tags.indexOf(filters.tag) >= 0);
    });
  }

  private removeUndefinedValues(values: Record<string, unknown>): Record<string, unknown> {
    return Object.keys(values).reduce<Record<string, unknown>>((result, key) => {
      if (values[key] !== undefined) {
        result[key] = values[key];
      }

      return result;
    }, {});
  }

  private createReport(): ISeedReport {
    return {
      categories: createProgressEntry('Categories'),
      models: createProgressEntry('AI Models'),
      departments: createProgressEntry('Departments'),
      users: createProgressEntry('Users'),
      tags: createProgressEntry('Tags'),
      prompts: createProgressEntry('Prompts'),
      activities: createProgressEntry('Activity Logs'),
      errors: []
    };
  }

  private markProgress(report: ISeedReport, key: SeedProgressKey, status: 'pending' | 'running' | 'complete' | 'failed'): void {
    report[key].status = status;
  }

  private recordFailure(report: ISeedReport, key: SeedProgressKey, message: string): void {
    report[key].failed += 1;
    report.errors.push(message);
  }

  private async listItemExistsByTitle(listTitle: string, title: string): Promise<boolean> {
    const existing = await this.web.lists.getByTitle(listTitle).items
      .select('Id,Title')
      .filter(`Title eq '${title.replace(/'/g, "''")}'`)
      .top(1)();

    return existing.length > 0;
  }

  private async getLookupIdByTitle(listTitle: string, title: string): Promise<number | undefined> {
    const existing = await this.web.lists.getByTitle(listTitle).items
      .select('Id,Title')
            .filter(`Title eq '${title.replace(/'/g, "''")}'`)
      .top(1)();

    return existing.length > 0 ? Number((existing[0] as ItemShape).Id) : undefined;
  }

  private async addUniqueItem(listTitle: string, values: Record<string, unknown>, report: ISeedReport, key: SeedProgressKey, title: string): Promise<boolean> {
    try {
      if (await this.listItemExistsByTitle(listTitle, title)) {
        report[key].skipped += 1;
        return false;
      }

      await this.web.lists.getByTitle(listTitle).items.add(values);
      report[key].created += 1;
      return true;
    } catch (error) {
      this.recordFailure(report, key, `${key}: ${error instanceof Error ? error.message : String(error)}`);
      return false;
    }
  }

  public async seedDemoData(callbacks?: ISeedDemoCallbacks): Promise<ISeedReport> {
    const report = this.createReport();

    const categorySeed = [
      { title: `${DEMO_PREFIX}Sales`, label: 'Sales', description: 'Revenue growth, pipeline, and customer conversion prompts.' },
      { title: `${DEMO_PREFIX}Marketing`, label: 'Marketing', description: 'Campaign, content, and engagement prompts.' },
      { title: `${DEMO_PREFIX}Finance`, label: 'Finance', description: 'Budgeting, forecasting, and reporting prompts.' },
      { title: `${DEMO_PREFIX}HR`, label: 'HR', description: 'Employee lifecycle, policy, and culture prompts.' },
      { title: `${DEMO_PREFIX}IT Support`, label: 'IT Support', description: 'Incident handling and technical support prompts.' },
      { title: `${DEMO_PREFIX}Customer Service`, label: 'Customer Service', description: 'Response, escalation, and satisfaction prompts.' },
      { title: `${DEMO_PREFIX}Project Management`, label: 'Project Management', description: 'Planning, tracking, and status prompts.' },
      { title: `${DEMO_PREFIX}Executive`, label: 'Executive', description: 'Leadership briefings and decision support prompts.' },
      { title: `${DEMO_PREFIX}Legal`, label: 'Legal', description: 'Policy, contract, and compliance prompts.' },
      { title: `${DEMO_PREFIX}Operations`, label: 'Operations', description: 'Process, efficiency, and execution prompts.' }
    ];

    const modelSeed = ['GPT-4.1', 'GPT-4o', 'Claude 3.5 Sonnet', 'Gemini 2.5 Pro', 'Mistral Large'].map((title) => `${DEMO_PREFIX}${title}`);
    const departmentSeed = ['Sales', 'Marketing', 'Finance', 'Human Resources', 'IT', 'Operations', 'Legal', 'Executive Office'].map((title) => `${DEMO_PREFIX}${title}`);
    const tagSeed = ['strategy', 'customer', 'policy', 'automation', 'risk', 'reporting', 'follow-up', 'onboarding', 'incident', 'training'].map((title) => `${DEMO_PREFIX}${title}`);
    const userSeed = [
      { title: `${DEMO_PREFIX}Victoria Akindele`, email: 'akindele@lotusbetaanalytics.com', department: `${DEMO_PREFIX}Sales` },
      { title: `${DEMO_PREFIX}Victor Osahon`, email: 'victor.osahon@lotusbetaanalytics.com', department: `${DEMO_PREFIX}Operations` },
      { title: `${DEMO_PREFIX}Miracle Shodeinde`, email: 'miracle.shodeinde@lotusbetaanalytics.com', department: `${DEMO_PREFIX}Marketing` },
      { title: `${DEMO_PREFIX}Sijibomi Adeoye`, email: 'sijibomi.adeoye@lotusbetaanalytics.com', department: `${DEMO_PREFIX}HR` },
      { title: `${DEMO_PREFIX}Godwin Onah`, email: 'godwin.onah@lotusbetaanalytics.com', department: `${DEMO_PREFIX}IT` }
    ];

    const promptSeed: Array<{ title: string; category: string; aiModel: string; department: string; visibility: string; status: 'Published' | 'Draft'; featured: boolean; tags: string[]; description: string; promptText: string; usageCount: number; owner: string; lastReviewed: string; documentFile?: string; }> = [];

    const topics = [
      ['Sales', 'Pipeline Review Summary'],
      ['Sales', 'Discovery Call Follow-Up'],
      ['Sales', 'Proposal Personalization'],
      ['Sales', 'Renewal Risk Brief'],
      ['Marketing', 'Campaign Brief Generator'],
      ['Marketing', 'Brand Voice Refiner'],
      ['Marketing', 'Social Post Variants'],
      ['Marketing', 'Event Follow-Up Email'],
      ['Finance', 'Budget Variance Analysis'],
      ['Finance', 'Quarterly Forecast Narrative'],
      ['Finance', 'Expense Review Assistant'],
      ['Finance', 'Board Pack Summary'],
      ['HR', 'Onboarding Welcome Plan'],
      ['HR', 'Policy Clarification Reply'],
      ['HR', 'Performance Review Draft'],
      ['HR', 'Employee Announcement'],
      ['IT Support', 'Incident Response Update'],
      ['IT Support', 'Password Reset Response'],
      ['IT Support', 'Knowledge Article Draft'],
      ['IT Support', 'Root Cause Summary'],
      ['Customer Service', 'Escalation Response'],
      ['Customer Service', 'Churn Recovery Email'],
      ['Customer Service', 'CSAT Follow-Up'],
      ['Customer Service', 'Service Delay Notice'],
      ['Project Management', 'Weekly Status Report'],
      ['Project Management', 'Risk Register Update'],
      ['Project Management', 'Stakeholder Brief'],
      ['Project Management', 'Meeting Notes Summary'],
      ['Executive', 'Leadership Brief'],
      ['Executive', 'Decision Memo'],
      ['Executive', 'Town Hall Talking Points'],
      ['Executive', 'Quarterly OKR Review'],
      ['Legal', 'Contract Review Checklist'],
      ['Legal', 'Compliance Reply Draft'],
      ['Legal', 'Policy Risk Summary'],
      ['Legal', 'NDA Clause Explanation'],
      ['Operations', 'Process Improvement Idea'],
      ['Operations', 'SOP Draft'],
      ['Operations', 'Shift Handover Notes'],
      ['Operations', 'Capacity Planning Note'],
      ['Sales', 'Upsell Recommendation'],
      ['Marketing', 'Persona Messaging Matrix'],
      ['Finance', 'Invoice Follow-Up'],
      ['HR', 'Recruiter Outreach Draft'],
      ['IT Support', 'Change Advisory Summary'],
      ['Customer Service', 'Refund Response'],
      ['Project Management', 'Delivery Timeline Update'],
      ['Executive', 'Investor Update Draft'],
      ['Legal', 'Policy Exception Review'],
      ['Operations', 'Warehouse Exception Note']
    ];

    topics.slice(0, 10).forEach((topic, index) => {
      promptSeed.push({
        title: `DEMO - ${topic[1]}`,
        category: `${DEMO_PREFIX}${topic[0]}`,
        aiModel: modelSeed[index % modelSeed.length],
        department: departmentSeed[index % departmentSeed.length],
        visibility: index % 3 === 0 ? 'Organization' : index % 3 === 1 ? 'Department' : 'Private',
        status: index % 4 === 0 ? 'Draft' : 'Published',
        featured: index % 5 === 0,
        tags: [tagSeed[index % tagSeed.length], tagSeed[(index + 3) % tagSeed.length]],
        description: `Enterprise demo prompt for ${topic[0].toLowerCase()} teams.`,
        promptText: `You are helping with ${topic[1].toLowerCase()}. Produce a concise, professional response for internal use.`,
        usageCount: 10 + index * 3,
        owner: userSeed[index % userSeed.length].title,
        lastReviewed: new Date(Date.now() - (index % 90) * 86400000).toISOString()
      });
    });

    const categoryMap = new Map<string, number>();
    const modelMap = new Map<string, number>();
    const departmentMap = new Map<string, number>();
    const tagMap = new Map<string, number>();
    const userMap = new Map<string, number>();

    this.markProgress(report, 'categories', 'running');
    for (const entry of categorySeed) {
      try {
        const created = await this.addUniqueItem(LISTS.categories, { Title: entry.title, Description: entry.description }, report, 'categories', entry.title);
        if (created) {
          callbacks?.onProgress?.(`✔ Categories created: ${entry.title}`);
        }
        const id = await this.getLookupIdByTitle(LISTS.categories, entry.title);
        if (id) {
          categoryMap.set(entry.title, id);
        }
      } catch (error) {
        this.recordFailure(report, 'categories', String(error));
      }
    }
    this.markProgress(report, 'categories', 'complete');

    this.markProgress(report, 'models', 'running');
    for (const title of modelSeed) {
      const created = await this.addUniqueItem(LISTS.models, { Title: title }, report, 'models', title);
      if (created) {
        callbacks?.onProgress?.(`✔ AI Models created: ${title}`);
      }
      const id = await this.getLookupIdByTitle(LISTS.models, title);
      if (id) {
        modelMap.set(title, id);
      }
    }
    this.markProgress(report, 'models', 'complete');

    this.markProgress(report, 'departments', 'running');
    for (const title of departmentSeed) {
      const created = await this.addUniqueItem(LISTS.departments, { Title: title }, report, 'departments', title);
      if (created) {
        callbacks?.onProgress?.(`✔ Departments created: ${title}`);
      }
      const id = await this.getLookupIdByTitle(LISTS.departments, title);
      if (id) {
        departmentMap.set(title, id);
      }
    }
    this.markProgress(report, 'departments', 'complete');

    this.markProgress(report, 'users', 'running');
    for (const entry of userSeed) {
      const created = await this.addUniqueItem(LISTS.users, { Title: entry.title, EMail: entry.email, Department: entry.department }, report, 'users', entry.title);
      if (created) {
        callbacks?.onProgress?.(`✔ Users created: ${entry.title}`);
      }
      const id = await this.getLookupIdByTitle(LISTS.users, entry.title);
      if (id) {
        userMap.set(entry.title, id);
      }
    }
    this.markProgress(report, 'users', 'complete');

    this.markProgress(report, 'tags', 'running');
    for (const title of tagSeed) {
      const created = await this.addUniqueItem(LISTS.tags, { Title: title }, report, 'tags', title);
      if (created) {
        callbacks?.onProgress?.(`✔ Tags created: ${title}`);
      }
      const id = await this.getLookupIdByTitle(LISTS.tags, title);
      if (id) {
        tagMap.set(title, id);
      }
    }
    this.markProgress(report, 'tags', 'complete');

    this.markProgress(report, 'prompts', 'running');
    for (const entry of promptSeed) {
      try {
        if (await this.listItemExistsByTitle(LISTS.prompts, entry.title)) {
          report.prompts.skipped += 1;
          continue;
        }

        const response = await this.web.lists.getByTitle(LISTS.prompts).items.add({
          Title: entry.title,
          Description: entry.description,
          PromptText: entry.promptText,
          CategoryId: categoryMap.get(entry.category),
          AIModelId: modelMap.get(entry.aiModel),
          DepartmentId: departmentMap.get(entry.department),
          TagsId: entry.tags
            .map((tag) => tagMap.get(tag))
            .filter((id): id is number => typeof id === 'number'),
          Status: entry.status,
          Visibility: entry.visibility,
          Featured: entry.featured,
          UsageCount: entry.usageCount,
          DocumentUrl: '',
          IsDeleted: false,
          PromptOwnerId: userMap.get(entry.owner),
          LastReviewed: entry.lastReviewed
        });

        report.prompts.created += 1;
        callbacks?.onProgress?.(`✔ Prompts created: ${entry.title}`);
      } catch (error) {
        this.recordFailure(report, 'prompts', `Prompts: ${error instanceof Error ? error.message : String(error)}`);
      }
    }
    this.markProgress(report, 'prompts', 'complete');

    const prompts = await this.getPrompts();

    this.markProgress(report, 'activities', 'running');
    const activityActions = ['Prompt Created', 'Prompt Updated', 'Prompt Viewed', 'Prompt Copied', 'Prompt Favorited', 'Prompt Shared'];
    const currentUser = await this.web.currentUser();
    for (let index = 0; index < 90; index += 1) {
      try {
        const prompt = prompts[index % prompts.length];
        if (!prompt) {
          report.activities.skipped += 1;
          continue;
        }
        const when = new Date(Date.now() - index * 24 * 60 * 60 * 1000).toISOString();
        const title = `${DEMO_PREFIX}${activityActions[index % activityActions.length]} ${index + 1}`;
        if (await this.listItemExistsByTitle(LISTS.activities, title)) {
          report.activities.skipped += 1;
          continue;
        }

        await this.web.lists.getByTitle(LISTS.activities).items.add({
          Title: title,
          ActivityType: activityActions[index % activityActions.length],
          PromptId: Number(prompt.id),
          PromptTitle: prompt.title,
          ActorId: currentUser.Id,
          Details: `Demo history entry created at ${when}`
        });
        report.activities.created += 1;
      } catch (error) {
        this.recordFailure(report, 'activities', `Activities: ${error instanceof Error ? error.message : String(error)}`);
      }
    }
    this.markProgress(report, 'activities', 'complete');

    return report;
  }

  public async clearDemoData(callbacks?: ISeedDemoCallbacks): Promise<ISeedReport> {
    const report = this.createReport();
    const demoFilter = `startswith(Title,'${DEMO_PREFIX.replace(/'/g, "''")}')`;
    const listKeys: Array<[SeedProgressKey, string]> = [
      ['prompts', LISTS.prompts],
      ['categories', LISTS.categories],
      ['models', LISTS.models],
      ['departments', LISTS.departments],
      ['users', LISTS.users],
      ['tags', LISTS.tags],
      ['activities', LISTS.activities]
    ];

    for (const [key, listTitle] of listKeys) {
      this.markProgress(report, key, 'running');
      try {
        const items = await this.web.lists.getByTitle(listTitle).items.select('Id,Title').filter(demoFilter).top(5000)();
        for (const item of items as ItemShape[]) {
          await this.web.lists.getByTitle(listTitle).items.getById(item.Id).delete();
          report[key].created += 1;
          callbacks?.onProgress?.(`✔ Cleared ${listTitle}: ${item.Title || item.Id}`);
        }
      } catch (error) {
        this.recordFailure(report, key, `${listTitle}: ${error instanceof Error ? error.message : String(error)}`);
      }
      this.markProgress(report, key, 'complete');
    }

    return report;
  }

  public async getPrompts(filters?: IPromptFilters): Promise<IPromptSummary[]> {
    const items = await this.web.lists.getByTitle(LISTS.prompts).items
      .select('Id,Title,Description,PromptText,Category/Id,Category/Title,AIModel/Id,AIModel/Title,Department/Id,Department/Title,Tags/Id,Tags/Title,Featured,Status,PromptOwner/Id,PromptOwner/Title,PromptOwner/EMail,Created,Modified,UsageCount,DocumentUrl,IsDeleted')
      .expand('Category,AIModel,Department,Tags,PromptOwner')
      .top(5000)();

    const prompts = (items as ItemShape[])
      .filter((item) => !item.IsDeleted)
      .map((item) => this.mapPrompt(item));

    return this.applyPromptFilters(prompts, filters);
  }

  public async getPromptById(id: number, logView = true): Promise<IPromptDetails | null> {
    const item = await this.web.lists.getByTitle(LISTS.prompts).items.getById(id)
      .select('Id,Title,Description,PromptText,Category/Id,Category/Title,AIModel/Id,AIModel/Title,Department/Id,Department/Title,Tags/Id,Tags/Title,Featured,Status,Visibility,PromptOwner/Id,PromptOwner/Title,PromptOwner/EMail,LastReviewed,Created,Modified,UsageCount,DocumentUrl,IsDeleted')
      .expand('Category,AIModel,Department,Tags,PromptOwner')();

    if ((item as ItemShape).IsDeleted) {
      return null;
    }

    const prompt = this.mapPromptDetails(item as ItemShape);
    if (logView) {
      await this.logActivity('Prompt Viewed', id, prompt.title, `Viewed prompt "${prompt.title}".`);
    }

    return prompt;
  }

  public async createPrompt(payload: ISharePointPromptWritePayload): Promise<IPromptDetails> {
    if (payload.status === 'Draft' && await this.hasCurrentUserDraft()) {
      throw new Error('You can only have one saved draft at a time. Publish or delete your existing draft before saving another.');
    }

    const currentUser = await this.web.currentUser();
    const created = await this.web.lists.getByTitle(LISTS.prompts).items.add(this.removeUndefinedValues({
      Title: payload.title,
      Description: payload.description,
      PromptText: payload.promptText,
      CategoryId: payload.categoryId,
      AIModelId: payload.aiModelId,
      DepartmentId: payload.departmentId,
      TagsId: payload.tagIds || [],
      Status: payload.status,
      Visibility: payload.visibility,
      Featured: Boolean(payload.featured),
      UsageCount: Number(payload.usageCount || 0),
      DocumentUrl: payload.documentUrl || '',
      IsDeleted: false,
      PromptOwnerId: payload.promptOwnerId || currentUser.Id,
      LastReviewed: payload.lastReviewed
    }));

    const createdItem = created as ItemAddShape;
    const promptId = Number(createdItem.data?.Id || createdItem.Id);
    const prompt = await this.getPromptById(promptId, false);

    if (!prompt) {
      throw new Error('Unable to reload the new prompt after creation.');
    }

    await this.logActivity('Prompt Created', promptId, prompt.title, `Created prompt "${prompt.title}".`);
    return prompt;
  }

  public async updatePrompt(id: number, payload: Partial<ISharePointPromptWritePayload>): Promise<IPromptDetails> {
    if (payload.status === 'Draft' && await this.hasCurrentUserDraft(id)) {
      throw new Error('You can only have one saved draft at a time. Publish or delete your existing draft before saving another.');
    }

    await this.web.lists
  .getByTitle(LISTS.prompts)
  .items
  .getById(id)
  .update(
    this.removeUndefinedValues({
      Title: payload.title,
      Description: payload.description,
      PromptText: payload.promptText,
      CategoryId: payload.categoryId,
      AIModelId: payload.aiModelId,
      DepartmentId: payload.departmentId,
      TagsId: payload.tagIds,
      Status: payload.status,
      Visibility: payload.visibility,
      Featured: payload.featured,
      UsageCount: payload.usageCount,
      DocumentUrl: payload.documentUrl,
      PromptOwnerId: payload.promptOwnerId,
      LastReviewed: payload.lastReviewed
    })
  );

    const updated = await this.getPromptById(id, false);
    if (!updated) {
      throw new Error('Unable to reload the prompt after update.');
    }

    await this.logActivity('Prompt Updated', id, updated.title, `Updated prompt "${updated.title}".`);
    return updated;
  }

  public async deletePrompt(id: number): Promise<void> {
    const prompt = await this.getPromptById(id, false);
    const currentUser = await this.web.currentUser();
    const isAdministrator = await this.isCurrentUserAdmin();

    if (!prompt) {
      throw new Error('Prompt was not found.');
    }

    if (!isAdministrator && prompt.createdById !== currentUser.Id) {
      throw new Error('You can only delete prompts that you created.');
    }

    await this.web.lists.getByTitle(LISTS.prompts).items.getById(id).update({
      IsDeleted: true,
      Status: 'Draft'
    });

    await this.logActivity('Prompt Deleted', id, prompt?.title, `Soft deleted prompt ${id}.`);
  }

  public async incrementUsage(id: number): Promise<void> {
    const prompt = await this.getPromptById(id, false);

    if (!prompt) {
      return;
    }

    await this.web.lists.getByTitle(LISTS.prompts).items.getById(id).update({
      UsageCount: prompt.usageCount + 1
    });

    await this.logActivity('Prompt Copied', id, prompt.title, `Copied prompt "${prompt.title}".`);
  }

  public async getCategories(): Promise<ICategorySummary[]> {
    const items = await this.web.lists.getByTitle(LISTS.categories).items.select('Id,Title,Description,PromptCount').top(5000)();
    return (items as ItemShape[]).map((item) => this.mapCategory(item));
  }

  public async getDepartments(): Promise<IDepartmentSummary[]> {
    const items = await this.web.lists.getByTitle(LISTS.departments).items.select('Id,Title').top(5000)();
    return (items as ItemShape[]).map((item) => this.mapLookup(item));
  }

  public async getModels(): Promise<IAiModelSummary[]> {
    const items = await this.web.lists.getByTitle(LISTS.models).items.select('Id,Title').top(5000)();
    return (items as ItemShape[]).map((item) => this.mapLookup(item));
  }

  public async getTags(): Promise<ITagSummary[]> {
    const items = await this.web.lists.getByTitle(LISTS.tags).items.select('Id,Title').top(5000)();
    return (items as ItemShape[]).map((item) => this.mapLookup(item));
  }

  public async getUsers(): Promise<IUserSummary[]> {
    try {
      const items = await this.web.lists.getByTitle(LISTS.users).items.select('Id,Title,EMail,Department').top(5000)();
      return (items as ItemShape[]).map((item) => this.mapUser(item));
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      if (message.indexOf(`List '${LISTS.users}' does not exist`) >= 0) {
        return [];
      }

      throw error;
    }
  }

  public async getActivities(): Promise<IActivityLogEntry[]> {
    const items = await this.web.lists.getByTitle(LISTS.activities).items
      .select('Id,ActivityType,PromptId,PromptTitle,Actor/Id,Actor/Title,Actor/EMail,Details,Author/Title,Created,Modified')
      .expand('Actor,Author')
      .top(5000)();

    return (items as ItemShape[]).map((item) => this.mapActivity(item));
  }

  public async getCurrentUserFavoritePromptIds(): Promise<number[]> {
    const currentUser = await this.web.currentUser();
    const actorNames = [currentUser.Title, currentUser.Email]
      .filter((value): value is string => Boolean(value))
      .map((value) => value.trim().toLowerCase());
    const activities = await this.getActivities();
    const favorites = new Set<number>();

    activities
      .filter((activity) => actorNames.indexOf(activity.actor.trim().toLowerCase()) >= 0)
      .sort((left, right) => new Date(left.createdDate).getTime() - new Date(right.createdDate).getTime())
      .forEach((activity) => {
        if (!activity.promptId) {
          return;
        }

        if (activity.action === 'Prompt Favorited') {
          favorites.add(activity.promptId);
        }

        if (activity.action === 'Prompt Unfavorited') {
          favorites.delete(activity.promptId);
        }
      });

    return Array.from(favorites);
  }

  public async setPromptFavorite(promptId: number, promptTitle: string, isFavorite: boolean): Promise<void> {
    await this.logActivity(
      isFavorite ? 'Prompt Favorited' : 'Prompt Unfavorited',
      promptId,
      promptTitle,
      isFavorite ? `Favorited prompt "${promptTitle}".` : `Removed prompt "${promptTitle}" from favorites.`
    );
  }

  public async logActivity(action: string, promptId?: number, promptTitle?: string, details?: string): Promise<void> {
    try {
      const currentUser = await this.web.currentUser();

      await this.web.lists.getByTitle(LISTS.activities).items.add(this.removeUndefinedValues({
        Title: action,
        ActivityType: action,
        PromptId: promptId,
        PromptTitle: promptTitle || '',
        ActorId: currentUser.Id,
        Details: details || ''
      }));
    } catch (error) {
      console.warn('Unable to write an Activity Log record.', error);
    }
  }

}

export default SharePointService;
