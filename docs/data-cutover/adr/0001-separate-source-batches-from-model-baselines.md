# Separate source batches from target-model baselines

Each cutover campaign uses one Git repository and gives every immutable source batch its own branch and worktree; a changed batch starts from a committed handoff instead of overwriting or copying the old workspace. The target-model baseline may advance within the same worktree only after its canonical PRD or domain authority accepts the change, after which affected evidence is invalidated and rerun. This keeps source evidence reproducible without preventing cutover findings from improving the product model.
