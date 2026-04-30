# autogen-core

## 描述
根据 specs 生成 prompts/skills/rules/agents 的最小工作流。

## 触发
当用户提到“生成 prompts/skills/rules/agents”或“省 token / VIBE coding”时使用。

## 工作流
- 先确认缺失信息（选择题/填空题）
- 只改 specs/templates 与生成器，不直接手改生成物
- 生成后给出可执行命令与文件路径
