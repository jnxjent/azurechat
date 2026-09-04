# TestSite deploy memo

TestSite deploy pattern for this workspace.

## TestSiteから本番へPRで反映する手順

TestSiteで動作確認できた変更を、本番リポジトリの`main`へPRで反映する場合の手順。

### 絶対ルール

- **Git Bashで1コマンドずつ実行し、毎回出力を確認してから次へ進む。**
- 本番へ直接pushしない。`origin/main`をbaseにしたPRを作成し、そのPRをmergeする。
- `git rebase`、force-push、`git add .`、`git add -A`は使用しない。
- 通常作業ツリーに未コミット変更がある場合、cherry-pickは**cleanな別worktree**で実行する。未コミット変更がある作業ツリーでcherry-pickすると、必要な差分が空として扱われるおそれがある。

### 手順

1. 通常作業ツリーで現状を確認する。

   ```bash
   cd ~/azurechat-office-work
   git status -sb
   git fetch origin
   ```

2. TestSite確認済みコミットがあるローカルブランチを基準に、`main`と内容が重複するコミットを除外する。`+`だけが本番PRへ必要、`-`は既に`main`に同等内容があるためcherry-pickしない。

   ```bash
   git cherry -v origin/main <testsite-verified-local-branch>
   ```

3. cleanな本番PR用worktreeを作成する。

   ```bash
   git worktree add --detach ../prod-pick-YYYYMMDD origin/main
   cd ../prod-pick-YYYYMMDD
   git status
   git switch -c <production-pr-branch>
   ```

4. `git cherry`で`+`だったコミットだけを、**古い順**にcherry-pickする。競合・空コミット・予期しない出力が出たら、その場で止めて確認する。

   ```bash
   git cherry-pick <oldest-required-commit>
   ```

5. PR内容を確認してpushする。

   ```bash
   git status
   git log --oneline --reverse origin/main..HEAD
   git diff --check origin/main...HEAD
   git push -u origin HEAD:<production-pr-branch>
   ```

6. GitHubでPRを作成する。

   - **base**: `main`
   - **compare**: `<production-pr-branch>`
   - コミット数・変更ファイル・差分を確認し、競合がないことを確認する。
   - PRを作成後、必要な確認・レビューを終えてから`Merge pull request`を実行する。

7. merge後、GitHub Actionsのbuild/deployが成功し、本番で動作確認できるまで完了としない。

### 本番PR前のビルド検証と、CIビルド失敗時の修正PR

- 本番PRを作る前に、**PR用のclean worktreeで**依存関係を入れ直し、`npm run build`を成功させる。別worktreeの`node_modules`は共有されないため、最初は`npm ci`が必要である。
- `npm ci`で表示されるdeprecated・脆弱性の警告と、`Browserslist`の更新通知は、ビルドが成功している限り今回のデプロイ可否を直接は妨げない。内容は別途計画して対応する。
- importした関数・型・新規ファイルの変更が、通常作業ツリーで未コミットのまま残っていないかを確認する。必要な実装がコミットに含まれないと、GitHub Actionsで`has no exported member`や`Module not found`になる。

```bash
# PR用worktreeのリポジトリ直下で、1コマンドずつ実行する。
cd src
npm ci
npm run build
cd ..
git status -sb
```

今回のPDF全文要約では、`sharepoint-summary-service.ts`が要求する
`SearchAllSharePointDocumentChunks` のexport追加コミットが、本番PRの5コミットに含まれていなかった。そのためActionsのコンパイルで停止し、**デプロイ処理は開始されなかった。本番環境には未反映**であった。

すでにmergeしたPRのActionsビルドが失敗した場合は、すぐRevertしない。失敗がデプロイ前かをログで確認し、原因の最小修正だけを新しいPRで反映する。

1. `git fetch origin`で、merge後の最新`origin/main`を取得する。
2. **リポジトリ直下から**最新`origin/main`を起点にclean worktreeを作り、修正PR用ブランチを切る。
3. 不足していた最小コミットだけをcherry-pickする。
4. `npm ci`、`npm run build`、`git diff --check origin/main...HEAD`で検証する。
5. 修正1コミットだけをpushし、`base: main`の新しい修正PRを作成・mergeする。
6. 修正PRのGitHub Actionsが成功し、本番の動作を確認する。

```bash
# リポジトリ直下（例: ~/azurechat-office-work）で実行する。
git fetch origin
git worktree add --detach ../prod-fix-YYYYMMDD origin/main
cd ../prod-fix-YYYYMMDD
git switch -c <production-fix-branch>
git cherry-pick <missing-fix-commit>
cd src
npm ci
npm run build
cd ..
git diff --check origin/main...HEAD
git push -u origin HEAD:<production-fix-branch>
```

### 本番障害時のrevert

- 一時的なActions失敗やビルド失敗は、すぐrevertせずログを確認する。
- 本番反映後の重大な機能不良では、merge済みPR画面の**Revert**を使う。
- `Revert`を押すとGitHubがrevert用ブランチを作成し、**Create pull request**画面を開く。
- revert PRのbaseが`main`であることを確認してPRを作成し、`Merge pull request`を実行する。
- revert PRをmergeした後も、GitHub Actionsのデプロイ成功と本番復旧を確認する。**Revertを押しただけでは本番は戻らない。**

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

## 通常作業ツリーがdirtyな場合の重複なし手順

通常作業ツリーに、すでにTestSiteへ反映済みの変更と今回の未反映変更が混在している場合は、通常作業ツリーでそのままcommitしない。次の2段階worktree方式を使う。

1. `git fetch testsite`後のTestSite先端を基準に、Local OK用worktreeを作成する。
2. 通常作業ツリーとTestSite先端の**ネット差分だけ**を、対象ファイルを明示したpatchとして作る。
3. Local OK用worktreeへpatchを適用し、テスト・build後にcommitする。
4. push直前に、最新TestSite先端から別のdetached worktreeを作る。
5. Local OK commitだけをcherry-pickし、差分と履歴を確認してpushする。

この方法では、TestSite先端をpatchの比較元とするため、すでにTestSiteへ存在する同一内容はpatchへ入らない。`git add .`、`git add -A`、通常作業ツリー全体のコピーは行わない。

```bash
cd ~/azurechat-office-work
git fetch testsite
git rev-parse refs/remotes/testsite/testsite/fix-toggle-selfscope-20260323

git worktree add -b localok/<feature>-YYYYMMDD ../testsite-localok-<feature>-YYYYMMDD refs/remotes/testsite/testsite/fix-toggle-selfscope-20260323

# 対象ファイルは今回変更したものだけを列挙する。
git diff --binary refs/remotes/testsite/testsite/fix-toggle-selfscope-20260323 -- <file-1> <file-2> > /tmp/testsite-<feature>-YYYYMMDD.patch

git -C ../testsite-localok-<feature>-YYYYMMDD apply --3way /tmp/testsite-<feature>-YYYYMMDD.patch
git -C ../testsite-localok-<feature>-YYYYMMDD status -sb
git -C ../testsite-localok-<feature>-YYYYMMDD diff --check

# テストとbuildを完了後、対象ファイルだけaddする。
git -C ../testsite-localok-<feature>-YYYYMMDD add -- <file-1> <file-2>
git -C ../testsite-localok-<feature>-YYYYMMDD diff --cached --name-status
git -C ../testsite-localok-<feature>-YYYYMMDD commit -m "<commit-message>"

# push用worktreeは、その時点の最新TestSite先端から新規作成する。
git fetch testsite
git worktree add --detach ../testsite-pick-<feature>-YYYYMMDD refs/remotes/testsite/testsite/fix-toggle-selfscope-20260323
git -C ../testsite-pick-<feature>-YYYYMMDD rev-parse HEAD refs/remotes/testsite/testsite/fix-toggle-selfscope-20260323
git -C ../testsite-pick-<feature>-YYYYMMDD cherry-pick <local-ok-commit-sha>
git -C ../testsite-pick-<feature>-YYYYMMDD status -sb
git -C ../testsite-pick-<feature>-YYYYMMDD log --oneline refs/remotes/testsite/testsite/fix-toggle-selfscope-20260323..HEAD
git -C ../testsite-pick-<feature>-YYYYMMDD push testsite HEAD:testsite/fix-toggle-selfscope-20260323
```

注意:

- patch作成前とpush用worktree作成前の両方で`git fetch testsite`する。
- `git apply --3way`で競合した場合は、その場で止める。競合ファイルを一括でours/theirsにしない。
- `git rev-parse HEAD refs/remotes/testsite/testsite/fix-toggle-selfscope-20260323`の2行が一致しない場合、cherry-pickしない。
- cherry-pick後の`git log ...remote..HEAD`には、今回追加するcommitだけが表示されることを確認する。
- pushがnon-fast-forwardになった場合、force-pushしない。最新remoteからpush用worktreeを作り直して、Local OK commitだけをcherry-pickする。

Important:

- **The push does not start the TestSite deployment automatically.**
- **After a successful push, manually start the GitHub Actions workflow described below.**
- Do not report the TestSite deployment as complete immediately after `git push`.
- The deploy target is not `testsite/main`.
- The deploy target is `testsite/fix-toggle-selfscope-20260323`.
- The core rule is: rebase no, clone no, cherry-pick yes.
- This keeps the original working directory clean and applies only the intended localOK commit onto the latest TestSite branch.
- Git rule: same content does not mean same commit. If the hash differs, Git treats it as a different history.

## Push後の手動Workflow起動（必須）

TestSiteへのpush成功後、GitHub ActionsのWorkflowを手動で起動して初めてAzure TestSiteへのデプロイが始まる。pushだけではデプロイされない。

1. 次のGitHub Actions画面を開く。
   `https://github.com/jnxjent/azurechat-gpt5-test/actions/workflows/azure-dev-validate.yml`
2. Workflow **`Build & deploy Next.js app to Azure Web App`** を選ぶ。
3. **Run workflow** を押す。
4. Branchは必ず **`testsite/fix-toggle-selfscope-20260323`** を選ぶ。`main`を選ばない。
5. もう一度 **Run workflow** を押して実行を開始する。
6. 起動されたRunが、直前にpushしたTestSiteコミットを対象としていることを確認する。
7. `build` と `test`（Azure deploy）の両Jobが成功するまで確認する。
8. Workflow成功後にTestSiteを開き、対象機能を確認する。

GitHub CLIを使う場合も、次のコマンドを単独で実行する。

```bash
gh workflow run azure-dev-validate.yml --repo jnxjent/azurechat-gpt5-test --ref testsite/fix-toggle-selfscope-20260323
```

Codexが案内する場合は、push成功後の次の案内を「自動デプロイが走る」や「デプロイ完了」として終わらせず、必ずこの手動Workflow起動へ進める。

## TestSite用Teams Bot設定チェック

TestSite用のAzure Botを新規作成・切り替えするときは、次の順番で確認する。

1. TestSiteの環境変数 `TEAMS_BOT_ID` と `TEAMS_BOT_SECRET` を、新しいBotの値へ変更して保存する。
2. `TEAMS_ENABLED=true` と `TEAMS_TENANT_ID` が正しいことを確認する。
3. Azure Botのメッセージングエンドポイントを次のURLに設定する。
   `https://azurechat-gpt5-test.azurewebsites.net/api/teams/messages`
4. 「ストリーミング エンドポイントを有効にする」はオフのままにする。
5. 構成画面の「適用」を押す。
6. Azure Botの「チャネル」でMicrosoft Teamsを選び、利用条件に同意する。
7. **チャネル画面の最後にある「適用」を必ず押す。選択しただけでは有効化されない。**
8. Microsoft Teamsチャネルが「実行中」になったことを確認してから、Teamsアプリを追加する。

再発防止メモ：

- Teamsで「無効なボット」「ボットが登録されており、チームのチャネルが有効になっていることを確認してください」と表示された場合、最初にAzure Botのチャネル画面で最後の「適用」を押したか確認する。
- 「適用」の押し忘れは、ZIPの競合や単なる反映待ちのように見える。先にZIPを再作成・再アップロードしない。
- Teamsチャネルが「実行中」になっていなければ、Teams側で待ち続けても解消しない。
