import Layout from '@theme/Layout';
import type { ReactElement } from 'react';

export default function PlaygroundRoute(): ReactElement {
  return (
    <Layout title="Playground" description="Interactive tego-sheet examples">
      <main className="container margin-vert--xl">
        <h1>Playground</h1>
        <p>The interactive presets are being prepared.</p>
      </main>
    </Layout>
  );
}
