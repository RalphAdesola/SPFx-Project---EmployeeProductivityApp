import * as React from 'react';
import styles from './EmployeeProductivityApp.module.scss';

export interface ILoadingScreenProps {
  isVisible: boolean;
}

export default function LoadingScreen(props: ILoadingScreenProps): React.ReactElement {
  return (
    <section
      className={`${styles.loadingPage} ${props.isVisible ? styles.loadingPageVisible : styles.loadingPageHidden}`}
      aria-busy="true"
      aria-live="polite"
      aria-label="Loading screen"
    >
      <div className={`${styles.loadingBackdrop} ${props.isVisible ? styles.loadingBackdropVisible : styles.loadingBackdropHidden}`} />

      <div className={`${styles.loadingCard} ${props.isVisible ? styles.loadingCardVisible : styles.loadingCardHidden}`}>
        <div className={styles.loadingLogoRow}>
          <div className={styles.loadingOrb}>
            <span />
            <span />
            <span />
            <span />
            <span />
            <span />
          </div>
        </div>

        <div className={styles.loadingCopy}>
          <p className={styles.loadingTitle}>Preparing your Intelligent workspace.</p>
          <p className={styles.loadingSubtitle}>Please wait...</p>
        </div>
      

        <div className={styles.loadingProgressTrack} aria-hidden="true">
          <div className={styles.loadingProgressBar} />
        </div>
      </div>
    </section>
  );
}
