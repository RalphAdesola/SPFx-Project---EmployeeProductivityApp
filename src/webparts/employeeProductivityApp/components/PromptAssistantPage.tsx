import * as React from 'react';
import * as mammoth from 'mammoth';
import * as pdfjsLib from 'pdfjs-dist';
import { Body1, Button, Card, Caption1, Dropdown, Field, Input, Option, Textarea, Title1, Title2 } from '@fluentui/react-components';
import { CopyRegular } from '@fluentui/react-icons';
import styles from './EmployeeProductivityApp.module.scss';
import type { ICategorySummary, IDepartmentSummary, IPromptWritePayload, ITagSummary } from './SharedTypes';

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
  'pdfjs-dist/build/pdf.worker.min.js',
  import.meta.url
).toString();

export interface IPromptAssistantPageProps {
  categories: ICategorySummary[];
  departments: IDepartmentSummary[];
  tags: ITagSummary[];
  onCreatePrompt: (payload: IPromptWritePayload) => Promise<void>;
  onBack: () => void;
  onSaveSuccess: () => void;
}

const outputPreferences = {
  summary: {
    label: 'Decision-ready summary',
    instruction: 'Present the result as an executive-ready summary with clear headings, decisions, risks, owners, and next steps.'
  },
  minutes: {
    label: 'Structured meeting record',
    instruction: 'Where the source contains a meeting, present the result as a structured record of attendees, discussion points, decisions, action items, and follow-up items.'
  },
  actions: {
    label: 'Action plan',
    instruction: 'Present action items in a table with the stated or suggested owner, due date, dependencies, and any unresolved questions.'
  },
  reusable: {
    label: 'Reusable delivery brief',
    instruction: 'Present the result as a reusable working brief with context, inputs, expected output, constraints, and next actions.'
  }
} as const;

type OutputPreference = keyof typeof outputPreferences;

interface IContextProfile {
  keywords: readonly string[];
  promptTitle: string;
  role: string;
  goal: (source: string) => string;
  instructions: readonly string[];
}

const includesAny = (source: string, keywords: readonly string[]): boolean => keywords.some((keyword) => source.indexOf(keyword) >= 0);

const getProjectDeliveryGoal = (source: string): string => {
  const normalizedSource = source.toLowerCase();
  const isDiscovery = includesAny(normalizedSource, ['discovery', 'stakeholder workshop', 'requirements workshop', 'requirements gathering']);
  const isApplication = includesAny(normalizedSource, ['application', 'system', 'solution', 'platform', 'portal']);
  const isCustomer = includesAny(normalizedSource, ['customer', 'client']);

  if (isDiscovery) {
    const subject = `${isCustomer ? 'customer ' : ''}${isApplication ? 'application ' : ''}discovery meeting`;
    return `Transform the ${subject} into an implementation-ready project brief covering the business problem, application requirements, agreed scope, proposed solution, decisions, risks, dependencies, action items, and next steps required to move the solution into design and development.`;
  }

  return 'Transform the source material into an implementation-ready delivery brief covering the business problem, requirements, scope, solution approach, decisions, risks, dependencies, ownership, and next steps for the project team.';
};

const contextProfiles: readonly IContextProfile[] = [
  {
    keywords: ['discovery', 'requirements', 'stakeholder', 'project', 'implementation', 'user story', 'solution design', 'scope', 'deliverable', 'acceptance criteria'],
    promptTitle: 'Project Delivery Brief',
    role: 'Project Management and Solution Delivery',
    goal: getProjectDeliveryGoal,
    instructions: [
      'Extract the business problem, current-state challenges, desired future state, and expected benefits.',
      'Identify functional and non-functional requirements, proposed technologies, and solution components.',
      'Separate confirmed decisions from assumptions, open questions, constraints, risks, and dependencies.',
      'Extract action items with owners and deadlines, then recommend next steps for design, development, testing, and implementation.',
      'Highlight information that still needs clarification from the customer or stakeholders.'
    ]
  },
  {
    keywords: ['employee', 'recruitment', 'onboarding', 'performance review', 'leave', 'payroll', 'workforce', 'people operations', 'human resources'],
    promptTitle: 'People Operations Brief',
    role: 'Human Resources and People Operations',
    goal: () => 'Transform the source material into a clear people-operations brief covering employee impact, policy implications, responsibilities, decisions, risks, required communications, and next actions.',
    instructions: [
      'Identify the relevant employee groups, policy requirements, responsibilities, and compliance considerations.',
      'Extract confirmed decisions, employee impacts, dependencies, unresolved questions, and required approvals.',
      'Recommend practical communication, implementation, and follow-up actions without inventing policy requirements.'
    ]
  },
  {
    keywords: ['budget', 'forecast', 'revenue', 'expense', 'financial', 'invoice', 'cost', 'profit', 'variance', 'cash flow'],
    promptTitle: 'Finance Analysis Brief',
    role: 'Finance and Business Analysis',
    goal: () => 'Analyse the source material and produce a decision-ready finance and business analysis brief covering financial drivers, assumptions, variances, risks, approvals, and recommended actions.',
    instructions: [
      'Extract financial figures, time periods, assumptions, variances, and business drivers stated in the source material.',
      'Identify financial risks, dependencies, required decisions, approvals, and follow-up actions.',
      'Clearly distinguish stated facts from estimates, assumptions, or information requiring validation.'
    ]
  },
  {
    keywords: ['incident', 'outage', 'service desk', 'ticket', 'root cause', 'severity', 'sla', 'downtime', 'technical issue', 'production issue'],
    promptTitle: 'IT Service Brief',
    role: 'IT Service Management',
    goal: () => 'Transform the source material into an actionable IT service management brief covering the incident or service issue, business impact, technical findings, ownership, recovery actions, risks, and prevention steps.',
    instructions: [
      'Identify the affected services, users, business impact, incident timeline, severity, and current service status.',
      'Separate confirmed technical findings from hypotheses, unknowns, and items requiring investigation.',
      'Extract recovery actions, owners, dependencies, escalation requirements, and recommended prevention or follow-up steps.'
    ]
  },
  {
    keywords: ['customer', 'client', 'opportunity', 'prospect', 'proposal', 'pipeline', 'deal', 'account manager', 'sales', 'quotation'],
    promptTitle: 'Customer Engagement Brief',
    role: 'Sales and Customer Engagement',
    goal: () => 'Analyse the source material and produce a structured customer engagement brief covering customer needs, commercial opportunity, proposed solution, commitments, risks, stakeholders, and next actions.',
    instructions: [
      'Extract customer needs, pain points, desired outcomes, stakeholders, commercial context, and solution expectations.',
      'Identify commitments, decisions, objections, risks, dependencies, and information needed before the next customer interaction.',
      'Recommend clear next actions for the account, sales, and delivery teams without inventing customer commitments.'
    ]
  },
  {
    keywords: ['campaign', 'brand', 'audience', 'content', 'marketing', 'launch', 'social media', 'engagement', 'messaging', 'market research'],
    promptTitle: 'Marketing Strategy Brief',
    role: 'Marketing and Communications',
    goal: () => 'Transform the source material into a marketing and communications brief covering audience needs, campaign objectives, messaging, channels, measures of success, risks, and next actions.',
    instructions: [
      'Identify target audiences, campaign objectives, brand or messaging requirements, channels, and success measures.',
      'Extract approved decisions, dependencies, deadlines, risks, and items that require stakeholder confirmation.',
      'Recommend coordinated next actions for content, campaign, and communications delivery.'
    ]
  },
  {
    keywords: ['contract', 'legal', 'compliance', 'regulation', 'clause', 'agreement', 'liability', 'privacy', 'gdpr', 'terms and conditions'],
    promptTitle: 'Legal Review Brief',
    role: 'Legal and Compliance',
    goal: () => 'Analyse the source material and produce a structured legal and compliance review brief covering obligations, risks, approvals, exceptions, open questions, and required next actions.',
    instructions: [
      'Identify stated obligations, parties, dates, approvals, compliance requirements, risks, and exceptions.',
      'Separate explicit contractual or regulatory statements from assumptions and points requiring qualified legal review.',
      'Do not provide legal advice or invent legal obligations that are not supported by the source material.'
    ]
  },
  {
    keywords: ['process', 'operations', 'workflow', 'handover', 'capacity', 'supplier', 'procurement', 'quality', 'procedure', 'service delivery'],
    promptTitle: 'Operations Improvement Brief',
    role: 'Operations and Process Improvement',
    goal: () => 'Transform the source material into an operational improvement brief covering the current process, issues, responsibilities, constraints, performance considerations, risks, and recommended next actions.',
    instructions: [
      'Map the stated process, handoffs, responsibilities, pain points, dependencies, controls, and constraints.',
      'Identify improvement opportunities, risks, decisions, and items requiring owner or stakeholder confirmation.',
      'Recommend practical next actions that are directly supported by the source material.'
    ]
  }
];

const fallbackContextProfile: IContextProfile = {
  keywords: [],
  promptTitle: 'Business Analysis Brief',
  role: 'Business Analysis and Operational Planning',
  goal: () => 'Transform the source material into a structured, decision-ready business brief that identifies its purpose, key facts, decisions, risks, ownership, dependencies, open questions, and next actions.',
  instructions: [
    'Identify the purpose, key facts, stakeholders, decisions, ownership, dependencies, and constraints stated in the source material.',
    'Separate confirmed information from assumptions, risks, and open questions that require clarification.',
    'Recommend practical next actions that are supported by the source material.'
  ]
};

const inferContextProfile = (source: string): IContextProfile => {
  const normalizedSource = source.toLowerCase();
  let selectedProfile = fallbackContextProfile;
  let highestScore = 0;

  contextProfiles.forEach((profile) => {
    const score = profile.keywords.reduce((total, keyword) => total + (normalizedSource.indexOf(keyword) >= 0 ? 1 : 0), 0);
    if (score > highestScore) {
      selectedProfile = profile;
      highestScore = score;
    }
  });

  return selectedProfile;
};

const buildContextAwarePrompt = (source: string, outcome: OutputPreference): string => {
  const profile = inferContextProfile(source);
  const instructions = [
    ...profile.instructions,
    outputPreferences[outcome].instruction,
    'Do not invent requirements, commitments, facts, or decisions that are not supported by the source material.',
    'Use a concise, professional tone and explicitly identify assumptions or gaps that need clarification.'
  ];

  return [
    `Act as an enterprise ${profile.role} assistant.`,
    '',
    `Goal: ${profile.goal(source)}`,
    '',
    'Source material:',
    source,
    '',
    'Instructions:',
    ...instructions.map((instruction) => `- ${instruction}`)
  ].join('\n');
};

type AssistantFormField = 'title' | 'category' | 'department' | 'tags' | 'prompt';

const assistantFormFieldLabels: Record<AssistantFormField, string> = {
  title: 'Prompt title',
  category: 'Category',
  department: 'Owner department',
  tags: 'Tags',
  prompt: 'Prepared prompt'
};

const readFileAsArrayBuffer = (file: File): Promise<ArrayBuffer> => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = () => resolve(reader.result as ArrayBuffer);
  reader.onerror = () => reject(new Error('Unable to read the selected file.'));
  reader.readAsArrayBuffer(file);
});

const extractDocumentText = async (file: File): Promise<string> => {
  const buffer = await readFileAsArrayBuffer(file);
  const extension = file.name.split('.').pop()?.toLowerCase();

  if (extension === 'docx') {
    const result = await mammoth.extractRawText({ arrayBuffer: buffer });
    return result.value;
  }

  if (extension === 'pdf') {
    const document = await pdfjsLib.getDocument({ data: new Uint8Array(buffer) }).promise;
    const pages: string[] = [];
    for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
      const page = await document.getPage(pageNumber);
      const content = await page.getTextContent();
      pages.push((content.items as Array<{ str?: string }>).map((item) => item.str || '').join(' '));
    }
    return pages.join('\n\n');
  }

  throw new Error('Select a DOCX or PDF file.');
};

const copyText = async (text: string): Promise<void> => {
  if (navigator.clipboard && window.isSecureContext) {
    await navigator.clipboard.writeText(text);
    return;
  }
  const element = document.createElement('textarea');
  element.value = text;
  element.style.position = 'fixed';
  element.style.opacity = '0';
  document.body.appendChild(element);
  element.select();
  const copied = document.execCommand('copy');
  document.body.removeChild(element);
  if (!copied) throw new Error('Your browser blocked clipboard access.');
};

export default function PromptAssistantPage(props: IPromptAssistantPageProps): React.ReactElement {
  const [source, setSource] = React.useState('');
  const [outcome, setOutcome] = React.useState<OutputPreference>('summary');
  const [prompt, setPrompt] = React.useState('');
  const [title, setTitle] = React.useState('');
  const [category, setCategory] = React.useState('');
  const [department, setDepartment] = React.useState('');
  const [tags, setTags] = React.useState<string[]>([]);
  const [message, setMessage] = React.useState<string | undefined>();
  const [error, setError] = React.useState<string | undefined>();
  const [saving, setSaving] = React.useState(false);
  const [isSaveConfirmationVisible, setIsSaveConfirmationVisible] = React.useState(false);
  const [validationErrors, setValidationErrors] = React.useState<AssistantFormField[]>([]);
  const [isExtracting, setIsExtracting] = React.useState(false);
  const [selectedFileName, setSelectedFileName] = React.useState<string | undefined>();

  const prepare = (): void => {
    setError(undefined);
    if (source.trim().length < 20) {
      setError('Paste meeting notes or a document excerpt before preparing a prompt.');
      return;
    }
    const profile = inferContextProfile(source);
    setPrompt(buildContextAwarePrompt(source, outcome));
    setTitle((current) => current || profile.promptTitle);
    setMessage('Prepared prompt is ready to copy or save as a draft.');
  };

  const handleFileSelected = async (event: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const input = event.currentTarget;
    const file = input.files?.[0];
    if (!file) return;
    input.value = '';
    setError(undefined);
    setMessage(undefined);
    setIsExtracting(true);
    try {
      const extractedText = await extractDocumentText(file);
      if (!extractedText.trim()) throw new Error('No readable text was found in this document.');
      setSource(extractedText);
      setSelectedFileName(file.name);
      setMessage(`Text extracted from ${file.name}. Review it, then prepare your prompt.`);
    } catch (extractionError) {
      setError(extractionError instanceof Error ? extractionError.message : 'Unable to extract text from the document.');
      setSelectedFileName(undefined);
    } finally {
      setIsExtracting(false);
    }
  };

  const getMissingFields = (): AssistantFormField[] => {
    const missing: AssistantFormField[] = [];
    if (!title.trim()) missing.push('title');
    if (!category) missing.push('category');
    if (!department) missing.push('department');
    if (!tags.length) missing.push('tags');
    if (!prompt) missing.push('prompt');
    return missing;
  };

  const hasValidationError = (field: AssistantFormField): boolean => validationErrors.indexOf(field) >= 0;

  const requestSave = (): void => {
    setError(undefined);
    const missingFields = getMissingFields();
    setValidationErrors(missingFields);
    if (missingFields.length) {
      setError(`Complete all required fields: ${missingFields.map((field) => assistantFormFieldLabels[field]).join(', ')}.`);
      return;
    }

    setIsSaveConfirmationVisible(true);
  };

  const save = async (): Promise<void> => {
    setError(undefined);
    setIsSaveConfirmationVisible(false);
    setSaving(true);
    try {
      await props.onCreatePrompt({ title: title.trim(), category, aiModel: '', description: `Prepared with Prompt Assistant using ${inferContextProfile(source).role} context inferred from the source material.`, promptText: prompt, tags, department, visibility: 'Organization', featured: false, status: 'Draft' });
      setMessage('Prompt saved as a draft in the Prompt Library.');
      props.onSaveSuccess();
    } catch (saveError) {
      setError(saveError instanceof Error ? saveError.message : 'Unable to save the prompt.');
    } finally { setSaving(false); }
  };

  return <section className={`${styles.employeeProductivityApp} ${styles.dashboardPage}`}><div className={styles.dashboardMain}>
    <div className={styles.adminHeader}><Title1>Prompt Assistant</Title1><Caption1>Paste meeting notes or document text to prepare a reusable prompt. Your pasted text remains in this browser session unless you save the resulting prompt.</Caption1><div><Button appearance="secondary" onClick={props.onBack}>Back to Dashboard</Button></div></div>
    {(message || error) && <Card className={`${styles.emptyStateCard} ${error ? styles.adminErrorNotification : styles.adminSuccessNotification}`}><Body1>{error || message}</Body1></Card>}
    <div className={styles.promptAssistantLayout}>
      <Card className={styles.sidePanelCard}><Title2>Prepare a prompt</Title2><Caption1 className={styles.assistantNotice}>AI generation will connect here when the approved Azure OpenAI endpoint is available but for now, the AI prepares a structured prompt you can just use in your Copilot, ChatGpt, Claude etc.</Caption1><div className={styles.formStack}>
        <Field label="Preferred output"><Dropdown value={outcome} selectedOptions={[outcome]} onOptionSelect={(_, data) => setOutcome((data.optionValue || 'summary') as OutputPreference)}>{(Object.keys(outputPreferences) as OutputPreference[]).map((key) => <Option key={key} value={key}>{outputPreferences[key].label}</Option>)}</Dropdown></Field>
        <Field label="Upload a DOCX or PDF"><input className={styles.assistantFileInput} type="file" accept=".docx,.pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/pdf" disabled={isExtracting} onChange={(event) => void handleFileSelected(event)} />{selectedFileName && <Caption1>Loaded: {selectedFileName}</Caption1>}</Field>
        <Field label="Meeting notes or document text"><Textarea value={source} onChange={(_, data) => setSource(data.value)} resize="vertical" placeholder="Paste notes, a transcript, or a document excerpt here." /></Field>
        <Button appearance="primary" onClick={prepare}>Prepare Prompt</Button>
      </div></Card>
      <Card className={styles.sidePanelCard}><Title2>Review and save</Title2><div className={styles.formStack}>
        <Field label="Prompt title" required validationMessage={hasValidationError('title') ? 'Prompt title is required.' : undefined} validationState={hasValidationError('title') ? 'error' : 'none'}><Input value={title} onChange={(_, data) => { setTitle(data.value); setValidationErrors((current) => current.filter((field) => field !== 'title')); }} /></Field>
        <Field label="Category" required validationMessage={hasValidationError('category') ? 'Category is required.' : undefined} validationState={hasValidationError('category') ? 'error' : 'none'}><Dropdown placeholder="Select category" value={category} selectedOptions={category ? [category] : []} onOptionSelect={(_, data) => { setCategory(data.optionValue || ''); setValidationErrors((current) => current.filter((field) => field !== 'category')); setError(undefined); }}>{props.categories.map((item) => <Option key={item.id} value={item.title}>{item.title}</Option>)}</Dropdown></Field>
        <Field label="Owner department" required validationMessage={hasValidationError('department') ? 'Owner department is required.' : undefined} validationState={hasValidationError('department') ? 'error' : 'none'}><Dropdown placeholder="Select department" value={department} selectedOptions={department ? [department] : []} onOptionSelect={(_, data) => { setDepartment(data.optionValue || ''); setValidationErrors((current) => current.filter((field) => field !== 'department')); }}>{props.departments.map((item) => <Option key={item.id} value={item.title}>{item.title}</Option>)}</Dropdown></Field>
        <Field label="Tags" required validationMessage={hasValidationError('tags') ? 'Select at least one tag.' : undefined} validationState={hasValidationError('tags') ? 'error' : 'none'}><Dropdown multiselect placeholder="Select tags" selectedOptions={tags} onOptionSelect={(_, data) => { setTags(data.selectedOptions); setValidationErrors((current) => current.filter((field) => field !== 'tags')); }}>{props.tags.map((item) => <Option key={item.id} value={item.title}>{item.title}</Option>)}</Dropdown></Field>
        <Field label="Prepared prompt" required validationMessage={hasValidationError('prompt') ? 'Prepare a prompt before saving.' : undefined} validationState={hasValidationError('prompt') ? 'error' : 'none'}><Textarea className={styles.assistantPromptText} value={prompt} readOnly resize="vertical" placeholder="Your prepared prompt will appear here." /></Field>
        {error && <Caption1 className={styles.assistantSaveError}>{error}</Caption1>}
        <div className={styles.promptPreviewActions}><Button appearance="secondary" icon={<CopyRegular />} disabled={!prompt} onClick={() => void copyText(prompt).then(() => setMessage('Prepared prompt copied to your clipboard.')).catch((copyError: Error) => setError(copyError.message))}>Copy Prompt</Button><Button appearance="primary" disabled={saving} onClick={requestSave}>Save Draft</Button></div>
      </div></Card>
    </div>
    {isSaveConfirmationVisible && <div className={styles.promptPreviewOverlay} role="presentation" onMouseDown={() => setIsSaveConfirmationVisible(false)}><section className={styles.promptConfirmationDialog} role="dialog" aria-modal="true" aria-labelledby="save-draft-confirmation-title" onMouseDown={(event) => event.stopPropagation()}><Title2 id="save-draft-confirmation-title">Are you sure?</Title2><Body1>Do you want to save this prompt as a draft?</Body1><div className={styles.promptPreviewActions}><Button appearance="secondary" onClick={() => setIsSaveConfirmationVisible(false)}>No</Button><Button appearance="primary" disabled={saving} onClick={() => void save()}>Yes</Button></div></section></div>}
  </div></section>;
}
