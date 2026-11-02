# Plan 001: Restore `master` to the `v1.11.0` release line

> **Trusted Publishing update (2026-11-02)**: This plan's fast-forward-only
> instructions below describe the original `9c02f26` baseline. They are no
> longer executable unchanged: `origin/master` now points to `a3c7d28`, which
> includes PR #2's OIDC publishing migration and diverges from `v1.11.0`.
> Integrate the current `origin/master` into the existing reviewed release
> branch using a normal merge, preserving the released feature commits and
> Trusted Publishing configuration. Do not reset `master` to `v1.11.0`,
> force-push, or restore the old API key workflow. After integration, verify
> both `v1.11.0` and the fetched `origin/master` are ancestors of the result,
> run the build/test/format checks below, and run `actionlint`. The result
> will be a merge commit rather than exactly `3da15f1`; the old exact-tip and
> three-commit-only acceptance criteria are superseded by these ancestry and
> OIDC-preservation checks. Keep operator authorization requirements for
> remote writes. This note does not mark the integration complete.

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the STOP conditions occurs, stop and report; do
> not improvise. When done, update this plan's row in `plans/README.md` unless
> a reviewer told you they own the index.
>
> **Drift check (run first)**: `git diff --stat 9c02f26..HEAD -- .`
> Expected before execution: no committed changes after `9c02f26` on the
> stale `master`. Untracked `.agents/`, `skills-lock.json`, and `plans/` do not
> justify staging, deleting, or resetting user files.

## Status

- **Priority**: P1
- **Effort**: S
- **Risk**: LOW
- **Depends on**: none
- **Category**: migration
- **Planned at**: commit `9c02f26`, 2026-07-12

## Why this matters

The default branch and `origin/master` point at `v1.8.1`, but the repository
has already tagged and published three later commits: `v1.9.0`, `v1.10.0`, and
`v1.11.0`. Those commits are a direct linear continuation of `master`, not a
conflicting fork. Leaving the default branch behind makes contributors develop
against obsolete code and makes the next tag capable of regressing already
released features. This plan restores one authoritative development line
without recreating or squashing released commits.

## Current state

- `master` / `origin/master` / `v1.8.1`: `9c02f26`.
- `v1.9.0`: `0b89ac1`, adds `RenderOptions`, warnings, and `TemplateCore.cs`.
- `v1.10.0`: `3551d62`, adds DOCX header/footer/textbox rendering.
- `v1.11.0`: `3da15f1`, adds XLSX range maintenance, PPTX media, and CLI.
- `git merge-base --is-ancestor master v1.11.0` returned success during the
  audit; `git rev-list --count master..v1.11.0` returned `3`.
- `src/NDocxTemplater/NDocxTemplater.csproj:7` currently contains:

  ```xml
  <Version>1.8.1</Version>
  ```

  The same file at `v1.11.0` contains `<Version>1.11.0</Version>`.
- Commit messages follow Conventional Commits, for example
  `feat: finish xlsx pptx cli expansion` and
  `fix: upgrade github actions to node 24 compatible versions`.
- `AGENTS.md` requires feature tests/examples/README and SemVer tag-based
  publishing. Those requirements were already satisfied by the three tagged
  commits; this synchronization must not manufacture a new version or tag.

## Commands you will need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| Confirm lineage | `git merge-base --is-ancestor master v1.11.0` | exit 0 |
| Confirm delta | `git log --oneline master..v1.11.0` | exactly the three released commits listed above |
| Integrate | `git merge --ff-only v1.11.0` | fast-forward to `3da15f1`; no merge commit |
| Build | `dotnet build NDocxTemplater.sln -c Release --disable-build-servers -m:1` | exit 0 |
| Test | `dotnet test NDocxTemplater.sln -c Release --no-build --disable-build-servers -m:1` | exit 0, all post-sync tests pass |
| Format | `dotnet format NDocxTemplater.sln --verify-no-changes --no-restore --verbosity minimal` | exit 0, no changes |

## Scope

**In scope**:

- Moving the integration branch pointer from `9c02f26` to existing tagged
  commit `3da15f1` using fast-forward semantics.
- Verifying every file already contained in commits `0b89ac1`, `3551d62`, and
  `3da15f1` as an indivisible released sequence.
- Updating only the plan status in `plans/README.md` after verification.

**Out of scope**:

- Rebasing, squashing, cherry-picking, or rewriting any tagged commit.
- Creating a new merge commit, package version, tag, or release.
- Editing the implementation while synchronizing the branches.
- Deleting or staging `.agents/`, `skills-lock.json`, or unrelated user files.
- Pushing or changing the remote default branch without explicit operator
  authorization.

## Git workflow

- Work on an integration branch named `codex/001-restore-release-line` created
  at `master`, unless the operator explicitly authorizes updating `master`
  directly.
- Use `git merge --ff-only v1.11.0`; do not create a commit manually.
- Do not push or open a PR unless the operator instructed it. If a PR is
  required, use the existing `codex/m4-m6-finish`/`v1.11.0` tip as the source;
  do not duplicate its commits.

## Steps

### Step 1: Prove the release history is still a safe fast-forward

Run:

```bash
git status --short
git merge-base --is-ancestor master v1.11.0
git rev-list --count master..v1.11.0
git log --oneline --decorate master..v1.11.0
```

**Verify**: lineage exits 0, the count is `3`, and the log contains only
`0b89ac1`, `3551d62`, and `3da15f1` in order. The status may show the known
untracked advisor files, but no tracked source changes.

### Step 2: Fast-forward the integration branch

Create/switch to `codex/001-restore-release-line` at `master`, then run:

```bash
git merge --ff-only v1.11.0
git rev-parse --short HEAD
git describe --tags --exact-match HEAD
```

**Verify**: HEAD is `3da15f1`, the exact tag is `v1.11.0`, and Git reports a
fast-forward rather than a merge commit.

### Step 3: Verify the complete released state

Run the build, test, and format commands from the table. Also run:

```bash
git diff --check master..HEAD
git status --short
```

**Verify**: all .NET commands exit 0; `git diff --check` has no output; no
tracked files were modified beyond the already-tagged commits and the advisor
plan status update.

### Step 4: Hand off the default-branch update

Report the verified fast-forward SHA and the exact remote action required. Do
not push unless explicitly authorized. Because `v1.11.0` already exists, do
not change the package version or trigger another release for this step.

**Verify**: once the operator updates the default branch, both
`git rev-parse origin/master` and `git rev-parse v1.11.0` resolve to the same
commit. If remote mutation was not authorized, mark the plan `BLOCKED` with
the single remaining action instead of claiming `DONE`.

## Test plan

- No new tests are added: this plan integrates three already-tested, tagged
  feature commits without changing them.
- Run the full post-sync test suite because the solution adds CLI tests and
  new DOCX/XLSX/PPTX behavior.
- Verify examples 17–21 and `tools/NDocxTemplater.Cli` are present after the
  fast-forward; do not regenerate example binaries during this plan.

## Done criteria

- [ ] The integration tip is exactly `3da15f1` / `v1.11.0`.
- [ ] The integration used `--ff-only`; no new merge/squash commit exists.
- [ ] Full Release build, test, and format verification exit 0.
- [ ] `src/NDocxTemplater/NDocxTemplater.csproj` reports version `1.11.0`.
- [ ] No new tag or NuGet package was created.
- [ ] The operator either updated `origin/master`, or the plan is marked
      `BLOCKED` with that explicit authorization-dependent action.
- [ ] `plans/README.md` status is updated.

## STOP conditions

Stop and report if:

- `master` is no longer an ancestor of `v1.11.0`.
- The delta contains anything other than the three audited release commits.
- `v1.11.0` was moved, deleted, or no longer resolves to `3da15f1`.
- Fast-forward integration would overwrite tracked or uncommitted user work.
- Any post-sync build/test/format command fails twice.
- Completing the remote update requires force-push or history rewriting.

## Maintenance notes

After this lands, protect the default branch as the only source for future
release tags. A reviewer should confirm that future version tags are ancestors
of the default branch; plan 002 turns that expectation into automated gates.
