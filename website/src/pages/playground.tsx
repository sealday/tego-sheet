import BrowserOnly from '@docusaurus/BrowserOnly';
import Layout from '@theme/Layout';
import type { ReactElement } from 'react';
import { Playground } from '../components/playground/playground';
import styles from './playground.module.css';

export function PlaygroundLoadingState(): ReactElement {
  return (
    <main className={styles.loading} aria-label="Loading playground" aria-busy="true">
      <div className={styles.loadingHeader} />
      <div className={styles.loadingGrid}>
        <div />
        <div />
      </div>
    </main>
  );
}

export default function PlaygroundRoute(): ReactElement {
  return (
    <Layout
      title="Interactive spreadsheet playground"
      description="Try controlled, uncontrolled, custom chrome, locale, and legacy JSON tego-sheet examples."
    >
      <BrowserOnly fallback={<PlaygroundLoadingState />}>{() => <Playground />}</BrowserOnly>
    </Layout>
  );
}
