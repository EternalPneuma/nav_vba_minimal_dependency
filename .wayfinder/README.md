# 本地 Wayfinder 票据运行说明

本目录是在仓库未配置远程 issue tracker 时使用的本地 Markdown tracker。

## 如何发起

开启一个新的 agent 会话，使用以下任一提示：

```text
使用 wayfinder 处理 `.wayfinder/交接文档编写规格地图.md` 的下一个前沿票据。
```

也可以按名称指定票据：

```text
使用 wayfinder 处理 `.wayfinder/交接文档编写规格地图.md` 中的“核实标准运行流程与停止条件”。
```

未指定票据时，agent 应选择一个状态为 `open`、`assignee` 为空且所有 `blocked_by` 均已关闭的票据。每个会话通常只解决一个票据。

## 认领与关闭约定

1. 开始工作前，将票据的 `assignee` 改为 `pi:<PI_SESSION_ID>`，避免并发会话重复处理。
2. 将解决过程产生的答案写入 `.wayfinder/resolutions/<票据名称>.md`，不要写进票据的 `## Question`。
3. 完成后把票据 `status` 改为 `closed`，并在 front matter 中添加指向答案的 `resolution` 链接。
4. 在地图的 `Decisions so far` 中追加一条以票据名称为链接的一行摘要。
5. 根据答案新增或调整票据、依赖关系和 `Not yet specified`；不要提前执行地图范围之外的文档编写工作。

## 票据类型

- `wayfinder:task`：agent 可独立核实仓库内事实，适合离开后运行。
- `wayfinder:prototype`：需要 agent 给出具体结构草案，再由交接负责人现场反馈。
- `wayfinder:grilling`：需要交接负责人现场回答和确认决策。

## 当前并行方式

当前两条路线可以并行：

- 运行路线从“核实标准运行流程与停止条件”开始。
- 维护路线从“核实系统架构与维护入口”开始。

两条路线完成后在“确定三份文档边界与编写规格验收标准”汇合。
