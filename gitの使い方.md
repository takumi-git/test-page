
# Git環境構築・VSCode連携手順まとめ

ローカルにある既存のフォルダをGitHubのリポジトリに紐付け、VSCodeでブランチ操作を行うまでの手順。

## 1. ローカルフォルダとGitHubの紐付け（Git Bash）

まだフォルダがGit管理されていない（`.git`がない）状態からのスタート。

### ① Gitの初期化
まず、対象のフォルダをGitリポジトリとして初期化する。

```bash
git init
```
② リモート（GitHub）の登録
GitHubのリポジトリURLを origin という名前で登録する。

```Bash
git remote add origin https://github.com/takumi-git/test-page.git
```
③ 登録確認
正しく登録されたか確認する。URLが表示されればOK。
```Bash
git remote -v
```
2. リモート情報の同期（Git Bash）
リモートにあるブランチ情報（main 以外のブランチなど）をパソコンに取り込む。

```Bash
git fetch
```
※ これを行わないと、VSCode側でリモートのブランチが表示されない場合がある。<br>

3. ブランチの切り替え（VSCode）
Git Bashでの操作が反映されているため、マウス操作だけで切り替えが可能。
方法A：ステータスバーから（推奨）
VSCodeの画面左下隅にあるブランチ名（master や main）をクリック。
画面上部にブランチリストが表示される。
切り替えたいブランチ名を選択する。
方法B：コマンドパレットから
Ctrl + Shift + P を押す。
「Git checkout」と入力して選択。
リストからブランチを選ぶ。
#### よく使うトラブルシューティング
##### Q. fatal: not a git repository と出た
##### A. git init を忘れている。
そのフォルダがまだGit管理下になっていないため、手順1の git init を実行する。
##### Q. fatal: a branch named 'xxx' already exists と出た
##### A. すでにそのブランチがローカルに存在する。
新しく作る必要はないので、単に切り替えるだけで良い。

```Bash
git checkout xxx
```
##### Q. VSCodeのリストにブランチが出てこない
##### A. 情報が古い可能性がある。
Git Bashで git fetch を実行してから、もう一度VSCodeで確認する。

```bash
git fetch
```