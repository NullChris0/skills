# Require named DET/RET/FTR inventories as first-class count-model fields

Detailed FPA already required integer RET/DET/FTR counts, but those integers could not be reviewed or compared across reruns. Persist named inventories in the count model and project them as `数据功能详表` and `事务功能详表`, one element per row; block when list length disagrees with the integer. Transaction DET rows carry a side (`输入` / `输出` / `启动触发` / `消息`); FTR rows reference data-function IDs; evidence stays on the function. Keep old packages as historical evidence and rerun rather than migrating them.
