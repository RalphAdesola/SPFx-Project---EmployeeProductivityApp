import * as React from 'react';
import * as mammoth from 'mammoth';
import * as pdfjsLib from 'pdfjs-dist';
import { Body1, Button, Card, Caption1, Dropdown, Field, Input, Option, Textarea, Title1, Title2 } from '@fluentui/react-components';
import { CopyRegular } from '@fluentui/react-icons';
import styles from './EmployeeProductivityApp.module.scss';
import type { ICategorySummary, IDepartmentSummary, IPromptWritePayload, ITagSummary } from './SharedTypes';

export interface IPromptAssistantPageProps {
  categories: ICategorySummary[];
  departments: IDepartmentSummary[];
  tags: ITagSummary[];
  onCreatePrompt: (payload: IPromptWritePayload) => Promise<void>;
  onBack: () => void;
  onSaveSuccess: () => void;
}

const outcomes = {
  summary: 'Summarize key discussion points, decisions, action items with owners and dates, open risks, and next steps.',
  minutes: 'Create formal meeting minutes with attendees, discussion points, decisions, action items, risks, and next steps.',
  actions: 'Extract action items with suggested owners, stated due dates, dependencies, and unresolved questions.',
  reusable: 'Create a clear reusable AI prompt with purpose, context, inputs, and expected output.'
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
  const [outcome, setOutcome] = React.useState('summary');
  const [prompt, setPrompt] = React.useState('');
  const [title, setTitle] = React.useState('');
  const [category, setCategory] = React.useState('');
  const [department, setDepartment] = React.useState('');
  const [tags, setTags] = React.useState<string[]>([]);
  const [message, setMessage] = React.useState<string | undefined>();
  const [error, setError] = React.useState<string | undefined>();
  const [saving, setSaving] = React.useState(false);
  const [isExtracting, setIsExtracting] = React.useState(false);
  const [selectedFileName, setSelectedFileName] = React.useState<string | undefined>();

  const prepare = (): void => {
    setError(undefined);
    if (source.trim().length < 20) {
      setError('Paste meeting notes or a document excerpt before preparing a prompt.');
      return;
    }
    setPrompt(['Act as an enterprise productivity assistant.', '', `Goal: ${outcome}.`, '', 'Source material:', source.trim(), '', 'Instructions:', outcomes[outcome as keyof typeof outcomes], 'Use a concise, professional tone and identify assumptions.'].join('\n'));
    setTitle((current) => current || `${outcome} prompt`);
    setMessage('Prepared prompt is ready to copy or save as a draft.');
  };

  const handleFileSelected = async (event: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file = event.target.files?.[0];
    if (!file) return;
    setError(undefined);
    setMessage(undefined);
    setIsExtracting(true);
    try {
      const extractedText = await extractDocumentText(file);
      if (!extractedText.trim()) throw new Error('No readable text was found in this document.');
      setSource(extractedText.trim());
      setSelectedFileName(file.name);
      setMessage(`Text extracted from ${file.name}. Review it, then prepare your prompt.`);
    } catch (extractionError) {
      setError(extractionError instanceof Error ? extractionError.message : 'Unable to extract text from the document.');
      setSelectedFileName(undefined);
    } finally {
      setIsExtracting(false);
      event.target.value = '';
    }
  };

  const save = async (): Promise<void> => {
    setError(undefined);
    if (!prompt || !title.trim() || !category) {
      setError('Prepare the prompt, give it a title, and select a category before saving.');
      return;
    }
    setSaving(true);
    try {
      await props.onCreatePrompt({ title: title.trim(), category, aiModel: '', description: `Prepared with Prompt Assistant for ${outcome}.`, promptText: prompt, tags, department, visibility: 'Organization', featured: false, status: 'Draft' });
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
      <Card className={styles.sidePanelCard}><Title2>Prepare a prompt</Title2><Caption1 className={styles.assistantNotice}>AI generation will connect here when the approved Azure OpenAI endpoint is available. For now, this prepares a structured prompt you can use in Copilot.</Caption1><div className={styles.formStack}>
        <Field label="Desired outcome"><Dropdown value={outcome} selectedOptions={[outcome]} onOptionSelect={(_, data) => setOutcome(data.optionValue || 'summary')}>{Object.keys(outcomes).map((key) => <Option key={key} value={key}>{key}</Option>)}</Dropdown></Field>
        <Field label="Upload a DOCX or PDF"><input className={styles.assistantFileInput} type="file" accept=".docx,.pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/pdf" disabled={isExtracting} onChange={(event) => void handleFileSelected(event)} />{selectedFileName && <Caption1>Loaded: {selectedFileName}</Caption1>}</Field>
        <Field label="Meeting notes or document text"><Textarea value={source} onChange={(_, data) => setSource(data.value)} resize="vertical" placeholder="Paste notes, a transcript, or a document excerpt here." /></Field>
        <Button appearance="primary" onClick={prepare}>Prepare Prompt</Button>
      </div></Card>
      <Card className={styles.sidePanelCard}><Title2>Review and save</Title2><div className={styles.formStack}>
        <Field label="Prompt title"><Input value={title} onChange={(_, data) => setTitle(data.value)} /></Field>
        <Field label="Category" required><Dropdown placeholder="Select category" value={category} selectedOptions={category ? [category] : []} onOptionSelect={(_, data) => setCategory(data.optionValue || '')}>{props.categories.map((item) => <Option key={item.id} value={item.title}>{item.title}</Option>)}</Dropdown></Field>
        <Field label="Owner department"><Dropdown placeholder="Select department" value={department} selectedOptions={department ? [department] : []} onOptionSelect={(_, data) => setDepartment(data.optionValue || '')}>{props.departments.map((item) => <Option key={item.id} value={item.title}>{item.title}</Option>)}</Dropdown></Field>
        <Field label="Tags"><Dropdown multiselect placeholder="Select tags" selectedOptions={tags} onOptionSelect={(_, data) => setTags(data.selectedOptions)}>{props.tags.map((item) => <Option key={item.id} value={item.title}>{item.title}</Option>)}</Dropdown></Field>
        <Field label="Prepared prompt"><Textarea className={styles.assistantPromptText} value={prompt} readOnly resize="vertical" placeholder="Your prepared prompt will appear here." /></Field>
        <div className={styles.promptPreviewActions}><Button appearance="secondary" icon={<CopyRegular />} disabled={!prompt} onClick={() => void copyText(prompt).then(() => setMessage('Prepared prompt copied to your clipboard.')).catch((copyError: Error) => setError(copyError.message))}>Copy Prompt</Button><Button appearance="primary" disabled={!prompt || saving} onClick={() => void save()}>Save Draft</Button></div>
      </div></Card>
    </div>
  </div></section>;
}
