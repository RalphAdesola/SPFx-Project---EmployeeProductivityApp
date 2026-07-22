import * as React from 'react';
import { Button, Title1 } from '@fluentui/react-components';
import { ArrowRightRegular, SparkleRegular } from '@fluentui/react-icons';
import styles from './EmployeeProductivityApp.module.scss';

export interface ILandingPageProps {
  onProceed: () => void;
}

export default function LandingPage(props: ILandingPageProps): React.ReactElement {
  return (
    <section className={styles.landingPage}>
      <div className={styles.landingCanvas}>
        <div className={styles.landingGlowLeft} />
        <div className={styles.landingGlowRight} />

        <div className={styles.landingTopBar}>
          <div className={styles.landingTopLeftBrand}>Employee Productivity App</div>
          <div className={styles.landingTopRightCard} aria-label="Top productivity prompt combinations">
            <SparkleRegular className={styles.landingTopRightIcon} />
            <Title1 className={styles.landingTopRightTitle}>Top Productivity Prompt Combinations</Title1>
          </div>
        </div>

        <div className={styles.landingHero}>
          <Title1 className={styles.landingTitle}>One Workspace. Every Prompt. Every Team.</Title1>
          <div className={styles.landingActions}>
            <Button className={styles.landingButton} appearance="primary" size="large" icon={<ArrowRightRegular />} onClick={props.onProceed}>
              Get Started
            </Button>
          </div>
        </div>
      </div>
    </section>
  );
}
