# Personal Skills

供团队内部使用的 Agent Skills。仓库为私有仓库；安装者需要具有仓库访问权限，并已配置可用于 GitHub 克隆的 Git 凭据。

## Skills

- `nesma-count`：从已确认证据生成四表 NESMA 计数包和计数模型。
- `unicom-deliverable`：从阶段一 JSON 生成联通交付 Word、功能清单和接口表。

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

## 测试

```bash
python -m unittest \
  skills/nesma-count/tests/test_generate_count_package.py \
  skills/unicom-deliverable/tests/test_build_delivery.py \
  skills/unicom-deliverable/tests/test_render_cards.py \
  skills/unicom-deliverable/tests/test_validate_delivery.py
bash tests/test_link_skills.sh
```

员工考核 golden fixture 位于 `tests/fixtures/nesma-employee-assessment-v3/`。
