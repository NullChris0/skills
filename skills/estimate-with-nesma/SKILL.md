---
name: estimate-with-nesma
description: Review evidence and produce a NESMA overview or detailed FPA review package.
disable-model-invocation: true
---

# NESMA 功能点估算

Run one evidence-review and counting session. Produce an unadjusted functional-size review package, never effort, cost, schedule, SNAP, adjustment factors, certification, or a customer-ready deliverable.

## Workflow

1. Read and complete [启动访谈.md](references/启动访谈.md). Stop until the user explicitly confirms its complete summary in a later turn.
2. After confirmation, read [知识来源.md](references/知识来源.md), [证据审查.md](references/证据审查.md), and [计数对象与方法选择.md](references/计数对象与方法选择.md).
3. Inspect the confirmed inputs with existing document and repository capabilities. Exercise production collection only when explicitly authorized and demonstrably state-neutral. Write only inside the output directory.
4. Apply [通用计数规则.md](references/通用计数规则.md) to identify candidates and remaining forks. Read [概要FPA.md](references/概要FPA.md) for overview FPA. For detailed FPA, read [详细FPA-数据功能.md](references/详细FPA-数据功能.md) and [详细FPA-事务功能.md](references/详细FPA-事务功能.md). Read [详细FPA-增强项目.md](references/详细FPA-增强项目.md) for enhancement counts. Do not silently pick a side when more than one legal reading remains.
5. Read and complete [计数判别访谈.md](references/计数判别访谈.md). Stop until the user explicitly confirms its discrimination summary in a later turn.
6. Stop formal counting when the selected method is blocked. A detailed-only request remains blocked rather than becoming overview FPA. Preserve candidates, evidence, conflicts, assumptions, confirmed discriminations, and non-counted requirements.
7. Read [复核与输出.md](references/复核与输出.md), create a transient JSON model, resolve this Skill's directory from the loaded `SKILL.md`, and run:

   ```bash
   python <skill-directory>/scripts/generate_review_package.py /tmp/nesma-input.json <new-output-directory>
   ```

8. Verify exactly three artifacts exist and report status as only `计数阻断` or `待复核`. Delete the transient JSON after successful generation. Treat workbook review edits as feedback for a new version; preserve the old package and rerun.

Never overwrite a non-empty output directory. Follow the target repository's documentation rules; otherwise use `./NESMA估算/<计数对象名称>/`.
