---
name: data-cutover
description: Plan, implement, iterate, and verify data cutovers for legacy-system rewrites when authoritative data moves to a confirmed target model. Use for immutable source snapshots, target-model feedback, legacy concept retirement, quarantine, reconciliation, cutover workspaces, and adapting tested artifacts toward production. Do not use for ordinary schema migrations, cleanup, reversible backfills, or one-off imports that do not transfer system-of-record authority.
---

# Data Cutover

Use a locked source batch to force the rewrite toward one explicit target model. Legacy data supplies evidence and mapping pressure; it does not earn compatibility fields, parallel runtime paths, or authority over the new model.

## Establish The Workspace

Inspect the target repository, its PRD and domain records, source exports, existing cutover artifacts, and known test/production database constraints before designing transformations. Resolve facts from available material; ask the user only for decisions or unavailable inputs.

Require a user-designated clean Git worktree. Never initialize Git inside a raw export directory. Keep source data, credentials, sensitive quarantine details, large SQL, and other protected artifacts outside ordinary Git.

Start the worktree with only:

- `CUTOVER.md`, the living cutover contract;
- `source-manifest.json`, the machine-verifiable source batch identity.

Record each logical source's `migrate` or `reference` role, filename, byte size, record count, format-specific structure fingerprint, and SHA-256. Derive one batch fingerprint from the sorted logical source names and file hashes. Verify the manifest before every run. Use standard or repository-native parsers; do not add a generic manifest framework.

One source batch owns one branch and worktree. If any source snapshot changes, stop that worktree, commit a minimal `HANDOFF.md`, and create a new branch/worktree from the handoff commit. The handoff names the changed source, reusable implementation and decisions, open model findings, and evidence that must be rerun. It does not make old evidence valid for the new batch.

Warn when the old system may keep changing after the snapshot or when snapshots from several systems need time alignment. This skill does not invent incremental capture or cross-system consistency rules.

## Maintain The Cutover Contract

Keep `CUTOVER.md` current with:

- scope, retired concepts, and the domain-level authority cutover;
- source batch fingerprint;
- target-model repository, commit, and governing PRD/domain/decision references;
- known test and production database product, version, compatibility mode, and execution-method differences;
- target-field mappings and legacy business-concept dispositions;
- identity, merge, split, deduplication, quarantine, and retirement decisions;
- destructive scope and a usable recovery point or method;
- generated artifacts and their source/model/generator bindings;
- `proved` and `not_proved` evidence.

For each business decision, record the decision, evidence, scope, accountable role, target-model version, and source batch. Use the repository's existing Issue, PRD, ADR, or decision mechanism. Do not require customer signatures, fixed approvers, approval forms, or external reference numbers.

## Feed Evidence Back Into The Target Model

Treat the target model as the current confirmed baseline, not an eternal input. When profiling or mapping shows that it cannot express a confirmed business need or invariant:

1. Record a model finding with aggregate evidence and impact in `CUTOVER.md`.
2. Pause only affected mappings and evidence. Reopen the whole batch when identity, authority cutover, source conservation, or recovery is affected.
3. Route the decision to the PRD/domain authority. The cutover workspace may propose; it does not decide the product model.
4. After the target repository accepts a new baseline, record `from`, `to`, finding, decision reference, and invalidated mappings/scripts/evidence.
5. Rerun affected work and update `proved`/`not_proved`. Keep the same worktree while the source batch is unchanged.

Do not copy the PRD into the cutover workspace. Reference its immutable version. Do not let a physical legacy table or dirty-data workaround become a target concept without a product decision.

Historical capability is a target-model choice. If the target preserves occurrence-time truth, validate against occurrence-time evidence and never silently substitute current master data. Do not prescribe zipper tables, snapshot tables, redundant fact fields, locked periods, or audit triggers when the target model does not require them.

## Transform Toward The New Model

Account for every target field as source, derivation, or new-system default. Classify every legacy business concept as migrated, retired, or evidence-only; group technical/log/cache columns when individual treatment adds no decision value.

Every readable row from a `migrate` source has exactly one outcome:

- **Migrated:** satisfies the confirmed mapping;
- **Quarantined:** belongs to a retained concept but cannot yet map safely;
- **Retired:** belongs to a deliberately retired legacy concept.

Require `read = migrated + quarantined + retired` for each `migrate` source. Validate reference sources and account for how mappings use them, but do not force their rows into migration outcomes. Corrupt structure, unreliable source identity, or manifest mismatch blocks the batch instead of becoming a fourth row outcome.

Prefer stable identity keys. A decision may use mutable or display fields only with measured null, duplicate, conflict, and coverage evidence; migrate only deterministic unambiguous matches. Add a per-row migration ledger only for rekeying, merge, split, deduplication, multi-source convergence, or another real traceability need. Direct one-to-one mappings can trace by source key.

Default to deterministic file-to-target conversion. Add database staging only when database-side transformation, restartable bulk processing, or controlled detail review actually needs it. Generating staging SQL does not prove staging was loaded.

Never add runtime fallback to the old system for quarantined or retired data. After authority cutover, use controlled remediation or a separate read-only archive when historical access is required.

## Prove Only What Happened

Always leave runnable checks for:

- source-manifest verification;
- source outcome conservation;
- bidirectional target reconciliation under the declared one-to-one, merge, or split rules;
- repeatability for the same source batch, target-model baseline, and generator commit.

Add format, mapping, database, aggregate, constraint, and recovery checks only when the actual inputs, target invariants, or destructive scope require them. Do not import a fixed checklist from a previous cutover.

Test success proves only the named environment and execution method. Check generated SQL against known production database and repository constraints so it is not obviously incompatible, but do not claim that MySQL client success proves replay through a production Web SQL executor. Continue iterating existing artifacts when the user asks for production adaptation.

Generation and execution are separate. Unresolved items may coexist with generated artifacts for testing and review; execution remains blocked where those items affect safety or correctness. Before a database write, resolve the exact target, authorization, destructive scope, and recovery path.

Bind every artifact to the source batch, target-model commit, and clean generator commit. Commit safe, reasonably sized artifacts normally. For external artifacts, record only type, controlled-location reference, SHA-256, source batch, target-model commit, and generator commit in `artifact-manifest.json`.

Report evidence as `proved` and `not_proved`; never collapse it into `ready`, `final`, or a signed completion claim. A generated SQL file is evidence of generation, not execution. A test load is evidence for that test environment, not production.
