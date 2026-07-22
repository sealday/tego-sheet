# Docusaurus Product Roadmap Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Publish the approved 33-item product Roadmap and 34 Mini-RFCs through the existing tego-sheet Docusaurus GitHub Pages site.

**Architecture:** Store display content in one typed module under `website/src/data`, render a dedicated `/roadmap` page, and expose the detailed Mini-RFCs as normal Docusaurus docs. Extend the existing navbar, documentation sidebar, browser tests, and visual regression suite; keep deployment owned by the existing GitHub Actions Pages workflow.

**Tech Stack:** Docusaurus 3, React 19, TypeScript 6, CSS Modules, Vitest, Playwright, GitHub Actions Pages.

## Global Constraints

- Keep the product positioned as an embeddable React component and TypeScript SDK.
- Every displayed item is `planned`; status indicators are non-interactive.
- Template printing and document generation is the primary direction.
- Phase means dependency order, not a promised date.
- Roadmap page labels and links come from one typed data source.
- Detailed designs live under `/docs/roadmap/*` and remain link-checked by Docusaurus.
- Deployment continues through `.github/workflows/ci.yml`; do not commit `website/build`.

---

### Task 1: Define the Roadmap contract with tests

**Files:**
- Create: `tests/unit/website/roadmap.test.ts`
- Modify: `tests/docs/docs.spec.ts`
- Create: `website/src/data/roadmap.ts`

**Interfaces:**
- Produces: `RoadmapPhaseId`, `RoadmapPhase`, `RoadmapItem`, `roadmapPhases`, `roadmapItems`, and `groupRoadmapItems()`.
- Requires: exactly 5 phases, 33 unique planned items, and Docusaurus links beginning with `/docs/roadmap/`.

- [x] Write the unit and browser assertions before creating the data/page modules.
- [x] Run them and confirm missing-module/navigation failures.
- [x] Implement immutable typed data from `website/docs/roadmap/index.md`.
- [x] Rerun the focused tests and confirm green.

### Task 2: Render and integrate the Roadmap page

**Files:**
- Create: `website/src/pages/roadmap.tsx`
- Create: `website/src/pages/roadmap.module.css`
- Modify: `website/docusaurus.config.ts`
- Modify: `website/src/pages/index.tsx`

**Interfaces:**
- Route: `/roadmap`.
- Heading: `Product roadmap`.
- Output: five semantic phase sections and 33 planned design links.

- [x] Render phase summaries and non-interactive planned cards with Docusaurus `Link`.
- [x] Add `Roadmap` to the navbar and homepage resources.
- [x] Add responsive 5/2/1-column behavior, visible focus, reduced-motion compatibility, and no horizontal overflow.
- [x] Run docs typecheck and browser tests.

### Task 3: Publish the Mini-RFC documentation set

**Files:**
- Create: `website/docs/roadmap/index.md`
- Create: `website/docs/roadmap/foundation.md`
- Create: `website/docs/roadmap/template-printing.md`
- Create: `website/docs/roadmap/formulas-data.md`
- Create: `website/docs/roadmap/analysis-visualization.md`
- Create: `website/docs/roadmap/extensibility.md`
- Create: `website/docs/roadmap/host-integrations.md`
- Modify: `website/sidebar-structure.ts`
- Create: `docs/superpowers/specs/2026-07-22-product-roadmap-design.md`

**Interfaces:**
- Sidebar category: `Product Roadmap`.
- Canonical overview ID: `roadmap/index`.
- Every page must build with `onBrokenLinks: 'throw'`.

- [x] Add the overview and six capability-domain documents.
- [x] Add an explicit sidebar category in dependency order.
- [x] Build Docusaurus and resolve every broken link as a release blocker.

### Task 4: Visual regression and publication

**Files:**
- Modify: `tests/docs-visual/docs-visual.spec.ts`
- Create: `tests/docs-visual/docs-visual.spec.ts-snapshots/roadmap-desktop-darwin.png`

**Interfaces:**
- Visual route: `/roadmap` at 1440×900, light scheme, reduced motion.
- Public URL: `https://sealday.github.io/tego-sheet/roadmap`.

- [x] Add and approve the Roadmap full-page snapshot.
- [x] Run format, lint, typecheck, unit, docs browser, docs visual, package, and production build gates.
- [ ] Commit with Conventional Commit syntax and push the feature branch.
- [ ] Merge to `main`, wait for the Pages workflow deployment, and verify the public Roadmap DOM contains 33 planned items without console errors or horizontal overflow.
