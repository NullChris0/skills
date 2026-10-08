# Personal Skills

供团队内部使用的 Agent Skills。仓库为私有仓库；安装者需要具有仓库访问权限，并已配置可用于 GitHub 克隆的 Git 凭据。

## Skills

- `estimate-with-nesma`：审查证据并生成可复核的 NESMA 概要或详细 FPA 审阅包。
- `write-unicom-as-built-requirements`：从竣工代码证据和已确认的 NESMA 结论生成联通需求说明书与功能清单。

## 安装

```bash
npx skills@latest add NullChris0/skills
```

## 本地维护

运行以下脚本，将 `skills/*/` 单向链接到 `~/.agents/skills/`：

```bash
scripts/link-skills.sh
```

本机的 `~/.claude/skills` 已链接到 `~/.agents/skills`，脚本无需重复写入。领域语境和 ADR 位于 `docs/<domain>/`，仅供维护，不随 Skill 安装。
