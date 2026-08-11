---
name: estimate-with-nesma
description: Review evidence and produce a NESMA overview or detailed FPA review package.
disable-model-invocation: true
---

# NESMA 功能点估算

Run one evidence-review and counting session. Produce an unadjusted functional-size review package, never effort, cost, schedule, SNAP, adjustment factors, certification, or a customer-ready deliverable.

## Workflow

1. Read [知识来源.md](references/知识来源.md), [证据审查.md](references/证据审查.md), and [计数对象与方法选择.md](references/计数对象与方法选择.md).
2. Resolve the user, counting scope, application boundary, counting subject, version point, count type, requested method, evidence roles, and output directory. Recommend the input profile from `证据审查.md`; accept other formats by semantic role.
3. Inspect inputs with existing document and repository capabilities. For every Git evidence repository and submodule, require a clean worktree and record HEAD. Exercise production collection only when explicitly authorized and demonstrably state-neutral. Write only inside the output directory.
4. Apply [通用计数规则.md](references/通用计数规则.md). Read [概要FPA.md](references/概要FPA.md) for overview FPA. For detailed FPA, read [详细FPA-数据功能.md](references/详细FPA-数据功能.md) and [详细FPA-事务功能.md](references/详细FPA-事务功能.md). Read [详细FPA-增强项目.md](references/详细FPA-增强项目.md) for enhancement counts.
5. Stop formal counting when the selected method is blocked. A detailed-only request remains blocked rather than becoming overview FPA. Preserve candidates, evidence, conflicts, assumptions, and non-counted requirements.
6. Read [复核与输出.md](references/复核与输出.md), create a transient JSON model, and run:

   ```bash
   python scripts/generate_review_package.py /tmp/nesma-input.json <new-output-directory>
   ```

7. Verify exactly three artifacts exist and report status as only `计数阻断` or `待复核`. Delete the transient JSON after successful generation. Treat workbook review edits as feedback for a new version; preserve the old package and rerun.

Never overwrite a non-empty output directory. Follow the target repository's documentation rules; otherwise use `./NESMA估算/<计数对象名称>/`.
