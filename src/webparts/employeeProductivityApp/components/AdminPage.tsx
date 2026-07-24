import * as React from 'react';
import { Body1, Body2, Button, Card, CardHeader, Caption1, Textarea, Title1 } from '@fluentui/react-components';
import { PlayRegular, DeleteRegular, ArrowClockwiseRegular } from '@fluentui/react-icons';
import styles from './EmployeeProductivityApp.module.scss';
import type { ISeedReport } from './SharedTypes';

export interface IAdminPageProps {
  canAccess: boolean;
  isProcessing: boolean;
  progressLog: string[];
  report: ISeedReport | null;
  onSeedAllData: () => Promise<void>;
  onClearDemoData: () => Promise<void>;
  onBack: () => void;
}

export default function AdminPage(props: IAdminPageProps): React.ReactElement {
  if (!props.canAccess) {
    return (
      <section className={`${styles.employeeProductivityApp} ${styles.dashboardPage}`}>
        <div className={styles.dashboardMain}>
          <Card className={styles.emptyStateCard}>
            <CardHeader header={<Title1>Admin</Title1>} />
            <Body1>You do not have permission to access this page.</Body1>
            <Button appearance="secondary" onClick={props.onBack}>Back</Button>
          </Card>
        </div>
      </section>
    );
  }

  return (
    <section className={`${styles.employeeProductivityApp} ${styles.dashboardPage}`}>
      <div className={styles.dashboardMain}>
        <Title1>Admin</Title1>
        <Card className={styles.sidePanelCard}>
          <CardHeader header={<Body1 style={{ fontWeight: 600 }}>Seed Demo Data</Body1>} />
          <Caption1>This operation creates sample records for demonstration purposes only.</Caption1>

          <div className={styles.formActions}>
            <Button appearance="primary" icon={<PlayRegular />} disabled={props.isProcessing} onClick={() => void props.onSeedAllData()}>
              Seed All Data
            </Button>
            <Button appearance="secondary" icon={<DeleteRegular />} disabled={props.isProcessing} onClick={() => void props.onClearDemoData()}>
              Clear Demo Data
            </Button>
            <Button appearance="subtle" icon={<ArrowClockwiseRegular />} disabled={props.isProcessing}>
              View Progress
            </Button>
          </div>

          <div className={styles.formStack}>
            {props.progressLog.length === 0 ? (
              <Textarea className={styles.whiteTextarea} readOnly value="Progress will appear here." />
            ) : (
              <Textarea className={styles.whiteTextarea} readOnly value={props.progressLog.join('\n')} />
            )}
          </div>

          {props.report && (
            <div className={styles.promptMetaGrid}>
              {Object.entries(props.report).filter(([key]) => key !== 'errors').map(([key, entry]) => (
                <Card key={key} className={styles.promptMetaItem}>
                  <Body2>{entry.label}</Body2>
                  <Caption1>Created: {entry.created}</Caption1>
                  <Caption1>Skipped: {entry.skipped}</Caption1>
                  <Caption1>Failed: {entry.failed}</Caption1>
                </Card>
              ))}
            </div>
          )}

          <Button appearance="subtle" onClick={props.onBack}>Back to Dashboard</Button>
        </Card>
      </div>
    </section>
  );
}
