import Link from '@docusaurus/Link';
import Layout from '@theme/Layout';
import type { ReactElement } from 'react';
import { groupRoadmapItems, roadmapPhases } from '../data/roadmap';
import styles from './roadmap.module.css';

export default function Roadmap(): ReactElement {
  const groups = groupRoadmapItems();

  return (
    <Layout
      title="Product roadmap"
      description="The planned product direction for tego-sheet, led by spreadsheet template printing and document generation."
    >
      <main className={styles.main}>
        <header className={styles.hero}>
          <div>
            <p className={styles.eyebrow}>Product direction</p>
            <h1>Product roadmap</h1>
          </div>
          <div className={styles.heroCopy}>
            <p>
              Template printing and document generation come first. Build spreadsheet templates,
              pass structured data, preview deterministic pages, and output only the intended sheet
              or range.
            </p>
            <p>
              Phases communicate capability dependencies and implementation order, not promised
              release dates. Every item below is planned and links to an implementation-level
              Mini-RFC.
            </p>
          </div>
        </header>

        <div className={styles.phases}>
          {roadmapPhases.map((phase) => (
            <section
              className={styles.phase}
              data-roadmap-phase={phase.id}
              key={phase.id}
              aria-labelledby={`roadmap-phase-${phase.id}`}
            >
              <header className={styles.phaseHeader}>
                <span>Phase {phase.id}</span>
                <p>{phase.label}</p>
                <h2 id={`roadmap-phase-${phase.id}`}>{phase.title}</h2>
                <p>{phase.summary}</p>
              </header>

              <ol className={styles.items}>
                {groups.get(phase.id)?.map((item) => (
                  <li data-roadmap-item={item.id} key={item.id}>
                    <div className={styles.status}>
                      <i aria-hidden="true" />
                      Planned
                    </div>
                    <Link to={item.designTo}>{item.title}</Link>
                    <p>{item.summary}</p>
                  </li>
                ))}
              </ol>
            </section>
          ))}
        </div>

        <footer className={styles.footer}>
          <p>
            The component and SDK own document behavior, extension protocols, and integration UI.
            Storage, accounts, collaboration backends, permissions services, and AI providers stay
            host-owned.
          </p>
          <Link to="/docs/roadmap">Read the complete Roadmap definition →</Link>
        </footer>
      </main>
    </Layout>
  );
}
