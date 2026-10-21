# Resolve count forks in batches and exclude unresolved items

Supersedes ADR 0006.

The earlier one-question-per-turn confirmation loop made large legacy counts unusably slow. After evidence review, the agent now presents all high-impact forks in one batch, each with evidence, legal readings, FP impact, and a proposed default. The user confirms or overrides the batch.

Confirmed answers are recorded as user-confirmation evidence. The final JSON contains no `assumptions` or `unresolved` fields. An item the user cannot confirm is excluded with a reason and does not enter UFP; the skill must not invent a conservative answer or emit a blocked total.
