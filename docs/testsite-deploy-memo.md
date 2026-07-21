# TestSite deploy memo

TestSite deploy pattern for this workspace.

## Absolute branch-name rule

- **Run the deployment procedure one command at a time (One-by-One).**
- **After every command, inspect its complete output and confirm that it matches the expected state before running the next command.**
- **Do not paste or run the full deployment sequence as one batch.**
- When Codex guides this deployment, Codex must provide only one command, wait for the user's result, evaluate it, and then provide the next command.
- If any output is unexpected, stop at that command. Do not continue to fetch, worktree creation, cherry-pick, or push until the cause is understood.
- **Never interpret, simplify, normalize, rename, or "correct" the TestSite branch/ref notation.**
- **Never assume the repeated `testsite/testsite/...` text is a typo. It is intentional.**
- Copy the branch and ref strings below exactly as written. Do not remove either `testsite` segment.
- Remote name: `testsite`
- Branch name on that remote: `testsite/fix-toggle-selfscope-20260323`
- Remote-tracking ref: `refs/remotes/testsite/testsite/fix-toggle-selfscope-20260323`
- Push refspec: `HEAD:testsite/fix-toggle-selfscope-20260323`
- Before cherry-pick, `HEAD` must equal `refs/remotes/testsite/testsite/fix-toggle-selfscope-20260323`.
- If Git reports an ambiguous ref, a different hash, or a non-fast-forward push, stop. Do not reinterpret the branch name, pull, rebase, or force-push.

- Run the deployment commands in **Git Bash**, not PowerShell.
- In Git Bash, start from the workspace with `cd ~/azurechat-office-work`.
- Do not use `git rebase` for TestSite deploys.
- Do not use `git clone` for this normal TestSite deploy flow.
- Use `git fetch testsite`, then create a separate `git worktree` from the current remote TestSite branch.
- Use `git cherry-pick` to put only the localOK commit(s) onto that TestSite worktree.
- The TestSite remote is `testsite`.
- The branch name on the `testsite` remote is `testsite/fix-toggle-selfscope-20260323`.
- Its remote-tracking ref is `refs/remotes/testsite/testsite/fix-toggle-selfscope-20260323`.
- The doubled `testsite/testsite/...` in the remote-tracking ref is intentional: the first `testsite` is the remote name and the second is part of the branch name.
- If direct push from the main working branch is rejected as non-fast-forward, it means the remote TestSite branch has newer commits. Do not rebase. Use the worktree + cherry-pick flow below.
- Trap to remember: after cherry-pick, the TestSite commit has a different hash from the local main working branch commit even if the content is identical. Example: local `6c528af` became TestSite `67b6e91`. A later local commit on top of `6c528af` is not automatically on top of `67b6e91`, so direct push from the main working branch can still be non-fast-forward. In that case, go to the TestSite worktree at `67b6e91` and cherry-pick only the later local commit.

Typical flow:

```bash
cd ~/azurechat-office-work

# After creating the Local OK commit, record its exact commit SHA.
git rev-parse HEAD

git fetch testsite

git worktree add --detach ../testsite-pick-YYYYMMDD refs/remotes/testsite/testsite/fix-toggle-selfscope-20260323
cd ../testsite-pick-YYYYMMDD

git status
git rev-parse HEAD refs/remotes/testsite/testsite/fix-toggle-selfscope-20260323
git log --oneline -5

git cherry-pick <local-ok-commit-sha>
git status
git log --oneline -3

git push testsite HEAD:testsite/fix-toggle-selfscope-20260323
```

Important:

- The deploy target is not `testsite/main`.
- The deploy target is `testsite/fix-toggle-selfscope-20260323`.
- The core rule is: rebase no, clone no, cherry-pick yes.
- This keeps the original working directory clean and applies only the intended localOK commit onto the latest TestSite branch.
- Git rule: same content does not mean same commit. If the hash differs, Git treats it as a different history.
