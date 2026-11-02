# Plan 002: Enforce release identity and NuGet audit gates

> **Executor instructions**: Follow this plan step by step and run every
> verification command. Stop rather than improvising if a STOP condition is
> met. Update this plan's status row in `plans/README.md` when done unless a
> reviewer owns the index.
>
> **Drift check (run first)**:
> `git diff --stat v1.11.0..HEAD -- .github/workflows/ci.yml .github/workflows/publish-nuget.yml Directory.Build.props src/NDocxTemplater/NDocxTemplater.csproj`
> Plan 001 must be complete. If HEAD does not contain `v1.11.0`, stop.

## Status

- **Priority**: P1
- **Effort**: S
- **Risk**: LOW
- **Depends on**: `plans/001-restore-release-line.md`
- **Category**: dx
- **Planned at**: commit `9c02f26`, 2026-07-12; implementation baseline `v1.11.0` (`3da15f1`)
- **Trusted Publishing update (2026-11-02)**: `origin/master` at `a3c7d28`
  already includes the OIDC migration from PR #2. Preserve that migration when
  integrating this plan's implementation; it is not a new migration task.

## Why this matters

The publish workflow accepts any `v*` tag and manual dispatch, packs whatever
version is in the project, then uses `--skip-duplicate`. A mismatched tag can
therefore publish a differently named package or silently skip an existing
package while the workflow looks successful. The repository also disables
NuGet audit globally, so future high/critical advisories do not block CI or a
release. This plan makes release identity fail-fast and restores online audit
gates without breaking offline local development.

## Current state

- `.github/workflows/publish-nuget.yml:3-7` permits `workflow_dispatch` and all
  tags matching `v*`.
- `.github/workflows/publish-nuget.yml` on `master` currently does:

  ```yaml
  - name: Pack
    run: dotnet pack src/NDocxTemplater/NDocxTemplater.csproj -c Release --no-build -o artifacts

  - name: NuGet login (OIDC)
    id: nuget-login
    uses: NuGet/login@8d196754b4036150537f80ac539e15c2f1028841 # v1
    with:
      user: whynpc

  - name: Publish to NuGet
    run: dotnet nuget push "artifacts/*.nupkg" --source "https://api.nuget.org/v3/index.json" --api-key "${{ steps.nuget-login.outputs.NUGET_API_KEY }}" --skip-duplicate
  ```

- The publish job has `contents: read` and `id-token: write` permissions.
  Keep both. NuGet policy matches owner `whynpc9`, repository
  `NDocxTemplater`, workflow file `publish-nuget.yml`, and no environment.
  Policy matching and token exchange must still be verified by an actual
  authorized GitHub Actions release; local lint does not verify them.
- `Directory.Build.props:3-4` contains:

  ```xml
  <RestoreIgnoreFailedSources>true</RestoreIgnoreFailedSources>
  <NuGetAudit>false</NuGetAudit>
  ```

- Keep that local/offline default if desired, but override it explicitly in
  CI and publish restore commands with `-p:NuGetAudit=true` and a documented
  failure threshold.
- `AGENTS.md:7-13` requires a SemVer project version, matching `vX.Y.Z` tag,
  tag-triggered NuGet workflow, and confirmed publication.
- Workflow actions are already on Node 24-compatible `actions/checkout@v5` and
  `actions/setup-dotnet@v5`; do not downgrade them.

## Commands you will need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| Baseline | `git merge-base --is-ancestor v1.11.0 HEAD` | exit 0 |
| Audited restore | `dotnet restore NDocxTemplater.sln -p:NuGetAudit=true -p:NuGetAuditLevel=high` | exit 0; no high/critical advisory |
| Build | `dotnet build NDocxTemplater.sln -c Release --no-restore --disable-build-servers -m:1` | exit 0 |
| Test | `dotnet test NDocxTemplater.sln -c Release --no-build --disable-build-servers -m:1` | all pass |
| Format | `dotnet format NDocxTemplater.sln --verify-no-changes --no-restore --verbosity minimal` | exit 0 |
| Workflow diff checks | `git diff --check` | no output |

## Scope

**In scope**:

- `.github/workflows/ci.yml`
- `.github/workflows/publish-nuget.yml`
- `src/NDocxTemplater/NDocxTemplater.csproj` for the required patch version
- `README.md` release documentation if manual-dispatch semantics change
- `plans/README.md` status only

**Out of scope**:

- Updating application dependencies merely because newer versions exist.
- Changing package ID, target frameworks, or replacing Trusted Publishing
  with long-lived API key authentication.
- Publishing GitHub Releases or signing packages.
- Removing offline-friendly local restore behavior unless the maintainer
  explicitly requests it.
- Pushing a tag or package without operator authorization.

## Git workflow

- Branch: `codex/002-enforce-release-gates`, based on the completed plan 001.
- Commit style: `fix: enforce nuget release version gates`.
- Do not push, tag, or open a PR unless instructed. The final authorized
  release must use a new patch version derived from the live version at
  execution time; do not assume `1.11.1` if another plan already released.

## Steps

### Step 1: Add explicit online audit gates

Update the restore step in both CI and publish workflows so hosted runners run
NuGet audit even though local defaults disable it. Use high severity as the
initial failure threshold and keep restore errors visible; do not combine the
audit with `RestoreIgnoreFailedSources=true` in a way that hides source failure.

Target command shape:

```yaml
run: dotnet restore NDocxTemplater.sln -p:NuGetAudit=true -p:NuGetAuditLevel=high
```

**Verify**: run the audited restore command locally with network access. It
must exit 0 and report no high/critical vulnerable packages. Grep both workflow
files and confirm each has `NuGetAudit=true`.

### Step 2: Make manual publication explicit

Give `workflow_dispatch` required inputs for the intended package version and
source ref. Configure checkout to use the requested ref for manual dispatch;
tag events continue to check out the tag. Do not allow a manual run to infer a
version silently from whichever default-branch commit happens to be current.

**Verify**: inspect the YAML and confirm both inputs are required, the checkout
ref is deterministic, and tag-triggered behavior remains supported.

### Step 3: Compare expected, project, and packed versions before push

Before packing/publishing:

1. Derive the expected version from `GITHUB_REF_NAME` for tag events after
   removing exactly one leading `v`; use the required manual version input for
   workflow dispatch.
2. Read the evaluated project version with:
   `dotnet msbuild src/NDocxTemplater/NDocxTemplater.csproj -getProperty:Version`.
3. Fail with a clear message if expected and project versions differ.
4. Pack to a clean `artifacts` directory.
5. Require exactly one primary package named
   `artifacts/NDocxTemplater.<expected>.nupkg`; exclude `.snupkg` from the
   primary package count.

Do not rely only on the tag name or only on the filename. Keep the checks in a
dedicated named workflow step so failures are visible.

**Verify**: locally run the shell logic once with the matching live project
version and once with a deliberately wrong environment value. Matching exits
0 and locates exactly one package; mismatch exits nonzero before push. Do not
request an OIDC token or use a real NuGet API key in local verification.

### Step 4: Stop masking duplicate release attempts

Remove `--skip-duplicate` from the official release push. A duplicate version
is a release-process error and must fail. Keep symbol package generation and
the existing NuGet source.

**Verify**: `rg -n "skip-duplicate" .github/workflows/publish-nuget.yml`
returns no matches. The push command still points at nuget.org and uses only
`steps.nuget-login.outputs.NUGET_API_KEY`, never a repository API key secret.
Keep the pinned login action after all build/test/package identity checks and
immediately before push so the short-lived key is acquired only when needed.

### Step 5: Document and version the fix

Update the README release section to state the tag/project/package identity
rule and required manual inputs. Increase the package patch version from the
live value. Do not create a feature example: this is release-infrastructure
hardening, not a template feature.

**Verify**: README describes both tag and manual flows; the csproj version is a
new patch; full build/test/format pass.

### Step 6: Perform the authorized release closeout

After merge and only with explicit operator authorization, create the matching
`vX.Y.Z` tag, push it, watch the `Publish NuGet` workflow, and verify the exact
package version is visible on NuGet. Confirm the OIDC login step succeeded;
do not accept a successful duplicate skip as proof of publication. If operator
authorization or a matching NuGet Trusted Publishing policy is unavailable,
mark the plan `BLOCKED` with the remaining closeout action. Do not request a
long-lived `NUGET_API_KEY` as a prerequisite or fallback.

**Verify**: tag text, project version, workflow run, and NuGet package version
all match exactly; the workflow did not pass through a duplicate skip.

## Test plan

- Matching tag/project/package version path succeeds.
- Mismatched tag and project version fails before pack/push.
- Manual dispatch without version/ref cannot start.
- Manual dispatch with mismatched version fails.
- Multiple or missing primary `.nupkg` files fail.
- Audited restore fails on high/critical advisories and succeeds on the current
  audited dependency graph.
- Use workflow-level checks and local shell simulation; do not add application
  unit tests for YAML.

## Done criteria

- [ ] CI and publish restore explicitly enable high-severity NuGet audit.
- [ ] Manual publish requires explicit version and source ref.
- [ ] Tag/manual expected version equals evaluated csproj version.
- [ ] Exactly one matching primary nupkg is required before push.
- [ ] `--skip-duplicate` is absent.
- [ ] OIDC job permissions, pinned login action, and temporary-key output are
      preserved; no `secrets.NUGET_API_KEY` reference remains in the workflow.
- [ ] README release instructions match the workflow.
- [ ] Package patch version is incremented.
- [ ] Full build/test/format pass.
- [ ] Matching tag is pushed and NuGet publication confirmed, or the plan is
      explicitly `BLOCKED` awaiting operator-authorized release closeout.
- [ ] `plans/README.md` status is updated.

## STOP conditions

Stop and report if:

- Plan 001 is not complete or HEAD does not contain `v1.11.0`.
- NuGet audit reports a high/critical advisory; do not suppress it in this plan.
- GitHub Actions expression behavior cannot distinguish tag/manual events
  deterministically.
- A correct version check would require printing or exposing any secret.
- The maintainer wants duplicate releases silently accepted.
- Release mutation is required but not explicitly authorized.

## Maintenance notes

Reviewers should scrutinize shell quoting and GitHub event differences. Future
changes to package ID or multi-package output must update the exact package
identity check rather than weakening it to a wildcard.
