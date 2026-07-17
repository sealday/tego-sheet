# tego-sheet Docusaurus 文档站与 GitHub Pages 设计

日期：2026-07-17

## 背景

`tego-sheet` 已提供 React、TypeScript 和 Vite 构建的电子表格组件、全屏本地 demo、迁移文档与完整质量门禁，但尚无面向使用者的公开文档站。项目需要一个部署到 GitHub Pages 的英文站点，同时承载手写使用文档、从 TypeScript 公共入口自动生成的 API Reference，以及可以直接操作不同 React 模式的在线 Playground。

## 目标

1. 在 `https://sealday.github.io/tego-sheet/` 提供静态文档站。
2. 以英文作为首发语言，并保留以后增加中文内容的结构空间。
3. 从 `src/index.ts` 的公开导出自动生成完整 API Reference。
4. 提供安装、核心概念、使用指南与迁移指南。
5. 提供五种预设模式的交互 Playground，用户无需编辑代码即可测试组件。
6. 将文档构建、浏览器验证、视觉验证和 Pages 发布纳入现有 GitHub CI 门禁。

## 非目标

- 首版不提供在线代码编辑器或任意 JavaScript 执行环境。
- 首版不启用博客、文档版本快照或中英文双语内容。
- 首版不接入 Algolia 或第三方站内搜索服务；导航、目录与浏览器查找足以覆盖当前内容规模。
- 不为 Playground 添加服务端存储、登录或跨访问持久化。
- 不公开 controller、Canvas engine 或其他内部模块。
- 不改变 `tego-sheet` 的公共 API、workbook JSON 或既有功能行为。
- 不把文档站作为移动端电子表格编辑能力的产品承诺。

## 选型

采用 Docusaurus 3、TypeDoc、`typedoc-plugin-markdown` 和 `docusaurus-plugin-typedoc`。

Docusaurus 是 React 应用，MDX 可以直接嵌入 React 组件，因此可以使用同一套 React 运行时加载 `TegoSheet`。Docusaurus 原生支持静态构建、文档导航、MDX、GitHub Pages 项目子路径和后续可选的国际化。TypeDoc 的 Docusaurus 集成可以在 Docusaurus 开发与构建过程中生成兼容的 Markdown 和侧边栏数据。

未选择 Astro Starlight。Starlight 的静态性能和内置 Pagefind 搜索有优势，但交互 Demo 需要额外的 Astro React integration、客户端 hydration 边界和另一套 TypeDoc 集成。当前站点的核心交互是大型 React Canvas 组件，保持 React 单一 UI 技术栈能减少集成风险。

未选择自建 Vite React 文档站。自建方案需要自行维护文档路由、侧边栏、目录、搜索、代码高亮、SEO 与失效链接检查，长期成本高于需求带来的收益。

参考：

- [Docusaurus MDX and React](https://www.docusaurus.io/docs/markdown-features/react)
- [Docusaurus deployment](https://www.docusaurus.io/docs/deployment)
- [TypeDoc Docusaurus integration](https://typedoc-plugin-markdown.org/plugins/docusaurus)
- [TypeDoc validation](https://typedoc.org/documents/Options.Validation.html)
- [GitHub Pages custom workflows](https://docs.github.com/en/pages/getting-started-with-github-pages/using-custom-workflows-with-github-pages)

## 仓库结构

站点源码位于仓库根目录下的 `website/`，依赖与 lockfile 继续由根目录 `package.json` 管理，不建立第二套包管理边界。

```text
website/
  docusaurus.config.ts
  sidebars.ts
  docs/
    getting-started/
    concepts/
    guides/
    migration/
    api/                 # TypeDoc 生成，不提交
  src/
    components/
      playground/
    css/
    pages/
      index.tsx
      playground.tsx
  static/
```

根目录增加以下命令：

- `docs:start`：构建组件库并启动 Docusaurus 开发服务器；TypeDoc 同步生成 API 文档。
- `docs:build`：构建组件库、生成 API Reference，并生成静态站点。
- `docs:serve`：本地服务静态构建结果。
- `test:docs`：运行文档站浏览器测试。
- `test:docs-visual`：运行文档站关键页面视觉测试。

文档站通过包名 `tego-sheet` 和已声明的 locale 子路径导入组件，不直接导入 `src/` 内部文件。`docs:start` 与 `docs:build` 必须先生成 `dist`，保证 Playground 验证的是公开 package exports，而不是内部实现。

## 页面与导航

全局导航固定为：

1. `Docs`
2. `API`
3. `Playground`
4. `GitHub`

首页包含产品定位、安装命令、最小 React 示例、核心能力和紧凑的实时表格预览。首页主要行动按钮分别进入 Quick Start 与 Playground。

文档侧边栏结构：

- `Getting Started`
  - Installation
  - Quick Start
  - Styling and Sizing
- `Core Concepts`
  - Controlled and Uncontrolled Workbooks
  - Workbook Data
  - Refs and Commands
  - Callbacks and Errors
- `Guides`
  - Custom Toolbar and Sheet Tabs
  - Locales
  - Validation and Filtering
  - Frozen Panes and Layout
  - Printing
- `Migration`
  - From x-data-spreadsheet
- `API Reference`
  - Components
  - Props
  - Handles
  - Workbook and Event Types
  - Slots, Locales, and Errors

手写文档与生成的 API Reference 使用同一视觉语言、左侧导航与右侧页内目录。TypeDoc 生成独立 sidebar 数据，`sidebars.ts` 将其挂载在 `API Reference` 分组。首版不冻结文档版本；只有项目出现需要同时维护的已发布不兼容版本时才启用 Docusaurus versioning。

## 视觉方向

站点以 Docusaurus 的可访问文档布局为基础，使用现有 demo 的深海军蓝、灰蓝与明亮蓝作为品牌色。首页与 Playground 可以使用深色工作区，正文文档保持高对比度的浅色阅读表面，并提供 Docusaurus 深色模式。

避免重度渐变、卡片堆叠和装饰性动画。视觉重点是代码、表格和公共 API。Playground 使用全宽自定义页面布局，不显示文档侧边栏；文档页保持标准三栏布局。

## API Reference 生成

TypeDoc 的唯一 entry point 是 `src/index.ts`。生成过程读取公共 re-export，并排除 private、protected、internal 与未导出的实现细节。

启用以下验证原则：

- 未解析的 `@link`、无效相对路径和未导出引用产生验证错误。
- 公共组件、函数、类型、接口、类和相关公开成员必须具有 TSDoc。
- 验证警告按错误处理。
- 极少数无法合理写文档的声明只能通过带理由的显式 allowlist 排除。

`website/docs/api/` 加入 Git 忽略规则，不提交生成文件。开发服务器和生产构建每次重新生成 API Reference，使源码、声明文件和公开文档保持一致。

## Playground

Playground 位于 `/playground`，通过 Docusaurus 客户端边界加载。服务端输出稳定的加载占位，避免在 SSR 阶段访问 Canvas、DOM 或浏览器 API。

页面包含五个 preset：

1. `Uncontrolled`：使用 `defaultValue`，允许自由编辑并观察回调。
2. `Controlled`：父组件拥有 `value`，接受 `onChange` 返回的新 workbook。
3. `Custom Chrome`：展示 typed toolbar renderer、自定义 sheet tabs 和公开 slot actions。
4. `Locales`：在 English、简体中文、German 和 Dutch 字典之间切换，并证明 locale 按组件隔离。
5. `Legacy JSON`：载入兼容的 sparse workbook JSON，并检查序列化结果。

模式由 `?mode=` 查询参数表示，使链接可分享，并支持浏览器前进与后退。非法值回退到 `uncontrolled` 并规范化 URL。

每次模式切换都会卸载旧的 `TegoSheet` 并创建新的隔离实例，同时清空旧事件。`Reset mode` 恢复当前 preset 的初始 workbook，不刷新页面。每个 preset 由独立的 fixture factory 和组件边界负责，禁止在同一挂载周期切换 controlled/uncontrolled 模式。

桌面布局右侧显示 inspector，窄屏折叠到表格下方。Inspector 包含：

- 当前模式说明和使用到的公共 API。
- 最近 50 条公开回调事件。
- 只读、可复制的当前 workbook JSON。
- 指向对应使用指南的链接。

Playground 不执行用户输入的代码，不访问私有对象，也不在浏览器存储中持久化 workbook。

## 数据流

```text
URL mode
  -> preset registry
  -> keyed React preset boundary
  -> TegoSheet public package API
  -> public callbacks
  -> capped event buffer / read-only JSON inspector
```

每个 preset 接收自己的初始 fixture、展示说明和对应文档链接。`TegoSheet` 的公开回调被规范化为只读事件记录；JSON inspector 只能通过公开回调结果或 `TegoSheetHandle.getValue()` 获取数据。

## 错误处理

- 非法 URL mode 回退到默认 preset，不产生空白页面。
- Docusaurus 页面级 Error Boundary 捕获 Playground 渲染错误，并提供 `Reset` 与 `Reload`。
- `TegoSheet.onError` 的可恢复错误进入事件面板，不冒充成功回调。
- Reset 失败时保留错误信息并允许重新挂载默认 preset。
- TypeDoc 生成失败、TypeScript 错误、Docusaurus 失效链接或重复路由都会终止生产构建。
- Pages 部署仅消费成功构建生成的 artifact，不从工作树直接发布。

## 测试策略

### 静态与架构测试

- 验证 TypeDoc 只有 `src/index.ts` 一个 entry point。
- 验证 Playground 只从公共 package exports 导入。
- 验证站点 `baseUrl` 为 `/tego-sheet/`，所有内部链接兼容项目子路径。
- 验证生成 API 目录不被 Git 跟踪。
- 验证未文档化的公共 API 会使文档构建失败。

### 组件测试

- 覆盖五个 preset 的 props 与初始 fixture。
- 覆盖 URL mode 解析、非法值回退和历史导航。
- 覆盖模式切换的卸载隔离与 Reset。
- 覆盖受控值接受、事件上限和 JSON inspector。
- 覆盖 Error Boundary 与 `onError` 展示。

### 浏览器测试

- 从 `/tego-sheet/` 打开首页，并进入 Quick Start、API 和 Playground。
- 在 Playground 切换五种模式并验证 URL。
- 实际编辑单元格，验证事件与 JSON 更新。
- 验证 Reset、刷新、前进与后退。
- 验证 GitHub Pages 子路径下没有资源、链接或路由 404。

### 视觉测试

- 首页 desktop。
- 典型文档页 desktop。
- Playground desktop 与窄屏布局。

视觉测试使用固定字体、固定 viewport 和现有 macOS 基线策略。API 生成页不逐页截图，结构与链接由构建及浏览器测试覆盖。

## CI 与 GitHub Pages

现有 CI 增加文档构建、文档浏览器测试和文档视觉测试。Pull Request 运行全部验证但不部署。

`main` push 的部署流程：

```text
library build
  -> TypeDoc generation
  -> Docusaurus build
  -> docs browser and visual verification
  -> upload-pages-artifact
  -> deploy-pages
```

部署 job 必须等待现有 commit policy、quality、Vitest、package contract、browser、visual、parity release 和新增 docs jobs 全部成功。部署 job 单独拥有 `pages: write` 与 `id-token: write`，其他 job 保持最小只读权限。GitHub Pages environment 使用 `github-pages`，部署 URL 由 `deploy-pages` 输出。

Actions 继续固定到不可变 commit SHA。站点首版使用 `url: https://sealday.github.io` 与 `baseUrl: /tego-sheet/`，不配置自定义域名。

## 验收标准

1. `https://sealday.github.io/tego-sheet/` 可以访问首页、文档、API Reference 和 Playground。
2. API Reference 由当前 `src/index.ts` 自动生成，且不包含内部模块。
3. 未文档化或存在无效链接的公共 API 无法通过 CI。
4. 使用文档覆盖安装、核心 React 状态模型、ref、回调、slots、locale、主要功能和迁移。
5. Playground 的五个 preset 均可直接操作，只使用公共 package API。
6. 模式 URL 可分享；切换、Reset、错误和事件面板行为符合设计。
7. GitHub Pages 项目子路径中的脚本、样式、图片、文档链接和路由均正常。
8. Pull Request 不部署；只有 `main` 的全部质量门禁通过后才发布 Pages。
9. 现有 library build、package exports、demo 和行为测试保持兼容。
