# Use JSON as the stage-one authority and the workbook as projection

Supersedes ADR 0002.

The first stage now emits exactly `计数模型.json` and `功能点计数证据.xlsx`. The JSON is the authoritative count model consumed by the delivery skill; the four-sheet workbook is the human-readable projection. There is no Markdown report and no independent detail workbook.

This prevents stage two from silently choosing between a workbook, a report, and a JSON copy. Any correction is made in the source evidence and a new version directory is generated; the workbook is not an editable second source of truth.
