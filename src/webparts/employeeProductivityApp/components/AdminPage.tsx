import * as React from 'react';
import { Body1, Body2, Button, Card, CardHeader, Caption1, Input, Title1 } from '@fluentui/react-components';
import { AddRegular, DeleteRegular } from '@fluentui/react-icons';
import styles from './EmployeeProductivityApp.module.scss';
import type { IAdminSummary, IDirectoryUser, IPromptSummary, IUserSummary } from './SharedTypes';

export interface IAdminPageProps {
  canAccess: boolean;
  isProcessing: boolean;
  users: IUserSummary[];
  admins: IAdminSummary[];
  prompts: IPromptSummary[];
  onAddAdmin: (email: string) => Promise<void>;
  onRemoveAdmin: (listItemId: number) => Promise<void>;
  onSearchDirectoryUsers: (searchText: string) => Promise<IDirectoryUser[]>;
  onDeletePrompt: (id: number) => Promise<void>;
  onBack: () => void;
}

export default function AdminPage(props: IAdminPageProps): React.ReactElement {
  const [adminEmail, setAdminEmail] = React.useState('');
  const [directoryUsers, setDirectoryUsers] = React.useState<IDirectoryUser[]>([]);
  const [isSearchingDirectory, setIsSearchingDirectory] = React.useState(false);
  const [isSavingAdmin, setIsSavingAdmin] = React.useState(false);
  const [isDeletingPromptId, setIsDeletingPromptId] = React.useState<number | undefined>(undefined);
  const [isRemovingAdminId, setIsRemovingAdminId] = React.useState<number | undefined>(undefined);
  const [isAdminListExpanded, setIsAdminListExpanded] = React.useState(false);
  const [isAddAdminExpanded, setIsAddAdminExpanded] = React.useState(false);
  const [message, setMessage] = React.useState<string | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);

  const userRows = React.useMemo(() => {
    const promptCounts = new Map<string, number>();
    props.prompts.forEach((prompt) => {
      const ownerKey = prompt.createdBy.trim().toLowerCase();
      promptCounts.set(ownerKey, (promptCounts.get(ownerKey) || 0) + 1);
    });

    const rows = new Map<string, { title: string; email?: string; department?: string; promptCount: number }>();
    props.users.forEach((user) => {
      const key = user.title.trim().toLowerCase();
      rows.set(key, {
        title: user.title,
        email: user.email,
        department: user.department,
        promptCount: promptCounts.get(key) || 0
      });
    });

    props.prompts.forEach((prompt) => {
      const key = prompt.createdBy.trim().toLowerCase();
      if (!rows.has(key)) {
        rows.set(key, { title: prompt.createdBy, promptCount: promptCounts.get(key) || 0 });
      }
    });

    return Array.from(rows.values()).sort((left, right) => right.promptCount - left.promptCount || left.title.localeCompare(right.title));
  }, [props.prompts, props.users]);

  const addAdmin = async (): Promise<void> => {
    setError(undefined);
    setMessage(undefined);
    setIsSavingAdmin(true);

    try {
      await props.onAddAdmin(adminEmail);
      setAdminEmail('');
      setMessage('Administrator access has been granted.');
    } catch (addError) {
      setError(addError instanceof Error ? addError.message : 'Unable to add the administrator.');
    } finally {
      setIsSavingAdmin(false);
    }
  };

  const searchDirectory = async (): Promise<void> => {
    setError(undefined);
    setMessage(undefined);
    setIsSearchingDirectory(true);

    try {
      setDirectoryUsers(await props.onSearchDirectoryUsers(adminEmail));
    } catch (searchError) {
      setError(searchError instanceof Error ? searchError.message : 'Unable to search the organization directory.');
      setDirectoryUsers([]);
    } finally {
      setIsSearchingDirectory(false);
    }
  };

  const deletePrompt = async (prompt: IPromptSummary): Promise<void> => {
    setError(undefined);
    setMessage(undefined);
    setIsDeletingPromptId(Number(prompt.id));

    try {
      await props.onDeletePrompt(Number(prompt.id));
      setMessage(`"${prompt.title}" was removed from the active prompt library.`);
    } catch (deleteError) {
      setError(deleteError instanceof Error ? deleteError.message : 'Unable to delete the prompt.');
    } finally {
      setIsDeletingPromptId(undefined);
    }
  };

  const removeAdmin = async (admin: IAdminSummary): Promise<void> => {
    setError(undefined);
    setMessage(undefined);
    setIsRemovingAdminId(admin.listItemId);

    try {
      await props.onRemoveAdmin(admin.listItemId);
      setMessage(`${admin.title} no longer has administrator access.`);
    } catch (removeError) {
      setError(removeError instanceof Error ? removeError.message : 'Unable to remove administrator access.');
    } finally {
      setIsRemovingAdminId(undefined);
    }
  };

  if (!props.canAccess) {
    return (
      <section className={`${styles.employeeProductivityApp} ${styles.dashboardPage}`}>
        <div className={styles.dashboardMain}>
          <Card className={styles.emptyStateCard}>
            <CardHeader header={<Title1>Administration</Title1>} />
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
        <div className={styles.adminHeader}>
          <Title1>Administration</Title1>
          <Caption1>Manage application administrators, users, and prompt governance.</Caption1>
          <div>
            <Button appearance="secondary" onClick={props.onBack}>Back to Dashboard</Button>
          </div>
        </div>

        {(message || error) && (
          <Card className={`${styles.emptyStateCard} ${error ? styles.adminErrorNotification : styles.adminSuccessNotification}`}>
            <Body1>{error || message}</Body1>
          </Card>
        )}

        <div className={styles.dashboardLayout}>
          <main className={styles.mainContent}>
            <Card className={styles.sidePanelCard}>
              <CardHeader header={<Body1 style={{ fontWeight: 600 }}>Users and Prompt Ownership</Body1>} description="Prompt totals are calculated from the Prompt Library." />
              <div className={styles.userOwnershipGrid}>
                {userRows.length === 0 ? (
                  <Caption1>No users or prompts are available yet.</Caption1>
                ) : userRows.map((user) => (
                  <div key={`${user.title}-${user.email || ''}`} className={styles.userOwnershipCard}>
                    <Body2>{user.title}</Body2>
                    <Caption1 className={styles.adminMetaLine}>Prompts created: {user.promptCount}</Caption1>
                  </div>
                ))}
              </div>
            </Card>

            <Card className={styles.sidePanelCard}>
              <CardHeader header={<Body1 style={{ fontWeight: 600 }}>Prompt Governance</Body1>} description="Only administrators can soft-delete prompts." />
              <div className={styles.promptGrid}>
                {props.prompts.length === 0 ? (
                  <Caption1>No active prompts are available.</Caption1>
                ) : props.prompts.map((prompt) => (
                  <div key={prompt.id} className={styles.promptCardTop}>
                    <div>
                      <Body2>{prompt.title}</Body2>
                      <Caption1 className={styles.adminMetaLine}>{prompt.category} · {prompt.createdBy} · {prompt.status}</Caption1>
                    </div>
                    <Button
                      appearance="secondary"
                      icon={<DeleteRegular />}
                      disabled={props.isProcessing || isDeletingPromptId === Number(prompt.id)}
                      onClick={() => void deletePrompt(prompt)}
                    >
                      Delete
                    </Button>
                  </div>
                ))}
              </div>
            </Card>
          </main>

          <aside className={styles.rightPanel}>
            <Card className={styles.sidePanelCard}>
              <div className={styles.adminDirectoryHeader} onClick={() => setIsAdminListExpanded((current) => !current)}>
                <div>
                  <Body1 className={styles.adminDirectoryTitle}>Administrators - {props.admins.length}</Body1>
                  <Caption1>Click to {isAdminListExpanded ? 'collapse' : 'view'} administrators</Caption1>
                </div>
                <Button
                  appearance="subtle"
                  icon={<AddRegular />}
                  aria-label="Add administrator"
                  onClick={(event) => {
                    event.stopPropagation();
                    setIsAddAdminExpanded((current) => !current);
                  }}
                >
                  Add Admin
                </Button>
              </div>

              {isAddAdminExpanded && (
                <div className={styles.formStack}>
                  <Caption1>Search Microsoft 365 directory, then select a user to grant access.</Caption1>
                  <Input value={adminEmail} onChange={(_, data) => setAdminEmail(data.value)} placeholder="Search name or email" />
                  <Button appearance="secondary" disabled={isSearchingDirectory || adminEmail.trim().length < 2} onClick={() => void searchDirectory()}>
                    Search directory
                  </Button>
                  {directoryUsers.map((user) => (
                    <Button key={user.id} appearance="subtle" onClick={() => { setAdminEmail(user.email); setDirectoryUsers([]); }}>
                      {user.displayName} · {user.email}
                    </Button>
                  ))}
                  <Button appearance="primary" icon={<AddRegular />} disabled={isSavingAdmin || !adminEmail.trim()} onClick={() => void addAdmin()}>
                    Add Admin
                  </Button>
                </div>
              )}

              {isAdminListExpanded && (
                <div className={styles.adminDirectoryList}>
                  {props.admins.length === 0 ? (
                    <Caption1>No administrators were found.</Caption1>
                  ) : props.admins.map((admin) => (
                    <div key={`${admin.listItemId}-${admin.email || admin.title}`} className={styles.promptMetaItem}>
                      <div className={styles.promptCardTop}>
                        <div>
                          <Body2>{admin.title}</Body2>
                          {admin.email && <Caption1 className={styles.adminMetaLine}>{admin.email}</Caption1>}
                        </div>
                        <Button
                          appearance="subtle"
                          icon={<DeleteRegular />}
                          aria-label={`Remove ${admin.title} as administrator`}
                          disabled={isRemovingAdminId === admin.listItemId}
                          onClick={() => void removeAdmin(admin)}
                        />
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </Card>
          </aside>
        </div>
      </div>
    </section>
  );
}
