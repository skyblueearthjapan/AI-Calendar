# AI Calendar - プロジェクトルール

## デプロイ手順（コード編集後は必ず実行）

コード編集が完了したら、以下の手順を必ず実行する：

1. **GitHub にコミット・プッシュ**
   - 変更内容を確認し、適切なコミットメッセージで commit
   - `git push` でリモートに反映

2. **差分確認後、clasp push**
   - `clasp push` 前に差分を確認（`clasp status` や `git diff` で）
   - 確認後 `clasp push` を実行して GAS プロジェクトに反映

3. **デプロイはユーザーが手動で行う**
   - clasp push 後のデプロイ作業はユーザーが実施する

## プロジェクト構成

- GAS プロジェクト ID: `1ahzovj4hocJlOplIwggf_T1skaBTA3uxNdTic50-ewR4qszPOXvARqsR`
- GitHub: `https://github.com/skyblueearthjapan/AI-Calendar.git`
- ソースコード: `src/` ディレクトリ（clasp rootDir）
- 参照元リポジトリ（編集禁止）: `https://github.com/skyblueearthjapan/AIYOTEI.git`
