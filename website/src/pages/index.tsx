import Link from '@docusaurus/Link';
import Layout from '@theme/Layout';
import { useState, type ReactElement } from 'react';
import { HomepagePreview } from '../components/homepage-preview';
import styles from './index.module.css';

const controlledExample = `import { useState } from 'react';
import { TegoSheet, type WorkbookData } from 'tego-sheet';
import 'tego-sheet/styles.css';

const initialWorkbook: WorkbookData = [
  { name: 'Budget', rows: { len: 100 }, cols: { len: 20 } },
];

export function BudgetSheet() {
  const [value, setValue] = useState(initialWorkbook);

  return <TegoSheet value={value} onChange={setValue} />;
}`;

const capabilities = [
  {
    number: '01',
    title: 'Typed React state',
    description: 'Choose a controlled value or an isolated uncontrolled defaultValue.',
  },
  {
    number: '02',
    title: 'Canvas performance',
    description: 'Keep large, interactive worksheets responsive in application layouts.',
  },
  {
    number: '03',
    title: 'Public slots and locales',
    description: 'Compose typed chrome and ship isolated English, Chinese, German, or Dutch UI.',
  },
  {
    number: '04',
    title: 'Compatible workbook JSON',
    description: 'Load and serialize the established sparse workbook data shape.',
  },
] as const;

const resources = [
  {
    label: 'Usage guides',
    description: 'State, refs, callbacks, custom chrome, locales, and layout.',
    to: '/docs/getting-started/quick-start',
  },
  {
    label: 'API reference',
    description: 'Generated directly from the exported TypeScript surface.',
    to: '/docs/api',
  },
  {
    label: 'Playground',
    description: 'Test five public integration modes in the browser.',
    to: '/playground',
  },
  {
    label: 'GitHub',
    description: 'Read the source, report issues, and follow releases.',
    href: 'https://github.com/sealday/tego-sheet',
  },
] as const;

function InstallCommand(): ReactElement {
  const [status, setStatus] = useState('');

  const copyInstallCommand = async (): Promise<void> => {
    try {
      await navigator.clipboard.writeText('npm install tego-sheet');
      setStatus('Install command copied');
    } catch {
      setStatus('Could not copy the install command');
    }
  };

  return (
    <div className={styles.installCommand}>
      <code>npm install tego-sheet</code>
      <button type="button" onClick={copyInstallCommand} aria-label="Copy install command">
        Copy
      </button>
      <span className={styles.srStatus} aria-live="polite">
        {status}
      </span>
    </div>
  );
}

export default function Home(): ReactElement {
  return (
    <Layout
      title="Typed React spreadsheet"
      description="A typed React spreadsheet for real application workflows"
    >
      <main className={styles.main}>
        <section className={styles.hero}>
          <div className={styles.heroCopy}>
            <p className={styles.eyebrow}>React · TypeScript · Canvas</p>
            <h1>Spreadsheet UI that belongs in your React application.</h1>
            <p className={styles.tagline}>
              tego-sheet combines a typed component boundary with a fast Canvas workspace and a
              workbook format your application can own.
            </p>
            <div className={styles.actions}>
              <Link className={styles.primaryAction} to="/docs/getting-started/quick-start">
                Start building
              </Link>
              <Link className={styles.secondaryAction} to="/playground">
                Open Playground
              </Link>
            </div>
          </div>
          <div className={styles.install}>
            <p className={styles.sectionLabel}>Install</p>
            <InstallCommand />
            <p>React 19 · TypeScript declarations included · MIT licensed</p>
          </div>
        </section>

        <section className={styles.exampleSection} aria-labelledby="controlled-example-title">
          <div className={styles.sectionIntro}>
            <p className={styles.sectionLabel}>A predictable boundary</p>
            <h2 id="controlled-example-title">Keep workbook state in React.</h2>
            <p>
              Pass a controlled value, accept each onChange snapshot, and use the same JSON shape
              everywhere else in your application.
            </p>
            <Link to="/docs/concepts/controlled-and-uncontrolled">
              Understand state ownership →
            </Link>
          </div>
          <pre className={styles.codeBlock} aria-label="Controlled React example">
            <code>{controlledExample}</code>
          </pre>
        </section>

        <section className={styles.capabilities} aria-labelledby="capabilities-title">
          <div className={styles.capabilityHeading}>
            <p className={styles.sectionLabel}>Built for application work</p>
            <h2 id="capabilities-title">One public surface. Four durable foundations.</h2>
          </div>
          <ol className={styles.capabilityList}>
            {capabilities.map((capability) => (
              <li key={capability.title}>
                <span>{capability.number}</span>
                <h3>{capability.title}</h3>
                <p>{capability.description}</p>
              </li>
            ))}
          </ol>
        </section>

        <section className={styles.previewSection} aria-labelledby="preview-title">
          <div className={styles.previewIntro}>
            <p className={styles.sectionLabel}>Public component, live</p>
            <h2 id="preview-title">Try the worksheet directly.</h2>
            <p>
              This compact viewport uses the same package entry and workbook data you will use in
              production. Use the Playground for full-size testing and inspectors.
            </p>
          </div>
          <HomepagePreview />
        </section>

        <section className={styles.resources} aria-labelledby="resources-title">
          <div className={styles.resourceHeading}>
            <p className={styles.sectionLabel}>Continue exploring</p>
            <h2 id="resources-title">From first render to complete integration.</h2>
          </div>
          <div className={styles.resourceList}>
            {resources.map((resource) =>
              'href' in resource ? (
                <a key={resource.label} href={resource.href}>
                  <span>
                    <strong>{resource.label}</strong>
                    <small>{resource.description}</small>
                  </span>
                  <b aria-hidden="true">↗</b>
                </a>
              ) : (
                <Link key={resource.label} to={resource.to}>
                  <span>
                    <strong>{resource.label}</strong>
                    <small>{resource.description}</small>
                  </span>
                  <b aria-hidden="true">→</b>
                </Link>
              ),
            )}
          </div>
        </section>
      </main>
    </Layout>
  );
}
