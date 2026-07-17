import Link from '@docusaurus/Link';
import Layout from '@theme/Layout';
import type { ReactElement } from 'react';
import styles from './index.module.css';

export default function Home(): ReactElement {
  return (
    <Layout
      title="Typed React spreadsheet"
      description="A typed React spreadsheet for real application workflows"
    >
      <main className={styles.main}>
        <section className={styles.hero}>
          <p className={styles.eyebrow}>React · TypeScript · Canvas</p>
          <h1>tego-sheet</h1>
          <p className={styles.tagline}>
            A typed React spreadsheet for real application workflows.
          </p>
          <div className={styles.actions}>
            <Link className={styles.primaryAction} to="/docs/getting-started/quick-start">
              Quick Start
            </Link>
            <Link className={styles.secondaryAction} to="/playground">
              Playground
            </Link>
          </div>
          <p className={styles.supportingLink}>
            Ready to install?{' '}
            <Link to="/docs/getting-started/installation">Read the installation guide</Link>.
          </p>
        </section>
      </main>
    </Layout>
  );
}
