# tego-sheet Roadmap

状态基线：2026-07-24

本 Roadmap 面向可嵌入业务系统的 React 电子表格组件与 TypeScript SDK。优先级表达能力依赖和建议实施顺序，不代表发布日期。

## Product direction

核心方向是电子表格模板打印与文档生成：业务用户在表格中设计模板，应用传入结构化数据，SDK 生成确定性的预览、指定区域浏览器打印、PDF、XLSX 和图片输出。

## Planned phases

| Phase | Capability                        | Status  | Design                                    |
| ----- | --------------------------------- | ------- | ----------------------------------------- |
| 2     | XLSX template output              | planned | [Template printing](template-printing.md) |
| 3     | CSV/TSV, XLSX and ODS interchange | planned | [Formulas and data](formulas-data.md)     |

已完成能力见 [Shipped Roadmap capabilities](shipped.md)。

## Maintenance rules

- 只将尚未交付的能力标为 `planned`。
- 能力完成并达到 Mini-RFC 验收标准后，移动到发布记录，不继续作为待办展示。
- Roadmap 项目必须链接到设计文档；没有技术和产品定义的想法不能进入主 Roadmap。
- Host integrations 只定义 SDK 协议和组件 UI 接入点，不承诺内建 SaaS 服务。
- 任何阶段调整必须同步更新总设计、Roadmap 索引和受影响 Mini-RFC。
