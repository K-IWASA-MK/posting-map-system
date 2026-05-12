# 開発理念 (Core Principle)

このプロジェクトの目的は、特定の地区やデータに依存しない**「どんなデータを入れても完璧に動作する、ポスティング管理システムの『究極のテンプレート（基盤）』」**を構築することである。
- すべてのロジックは汎用性を最優先する。
- 特定のサンプルデータ（例：三重第2区）に依存する数値や名称をハードコーディングしない。
- データ構造の変化（エリア数の増減など）に対して、動的に対応する設計を維持する。

# AGENTS.md

## AI Editing Rules

### 絶対ルール

- ユーザー確認なしで編集しない
- 削除前に必ず確認する
- ファイルを勝手に削除しない
- Accept All を前提に修正しない
- AIは提案のみ行う

---

## GAS Rules

### GASテンプレート構文

- <?!= include(...) ?> は正常構文
- GAS include構文を修正対象にしない
- HTMLエラーとして扱わない

### 重要ファイル

- doGet() を勝手に変更しない
- HtmlService を慎重に扱う
- index.html の構造を壊さない

---

## Workflow

AI = 提案
User = 承認

修正前に必ず確認する。

---

## Dangerous Actions

以下は必ず事前確認：

- ファイル削除
- include修正
- doGet変更
- 大量置換
- 自動修正
- CSS全体変更

---

## UI Rules

- モバイル優先
- シンプルUI
- 数字を大きく表示
- リアルタイム感を重視

---

## Development Policy

- GAS + HTML 構成を維持
- include構成を維持
- 小さく修正する
- 一度に大量変更しない

---

## Protected Files

以下のファイル・構文は重要領域。
事前確認なしで編集・削除しない。

### Protected Files

- index.html
- gas_v2.gs
- master_gas.gs
- doGet()
- HtmlService
- <?!= include(...) ?> 構文

### Important Rules

- Accept All を前提にしない
- 修正差分を必ず明示する
- ファイル削除前は必ず確認する
- scripts フォルダは慎重に扱う

## Workflow

AGENTS.md
↓
AI proposes changes
↓
Review Changes
↓
Manual Accept only
↓
Test application
↓
Create backup after work

## 30秒ルール (Performance Rule)

- 1つの修正・思考プロセスは**最大30秒**までとする。
- 30秒を超えて「Loading」が続く場合はタイムアウトとし、一度中断してユーザーに指示・確認を仰ぐ。
- 複雑な修正は細かく分割し、ステップごとに承認を得てから進める。

## 開発フロー

AGENTS.md を参照
↓
AIは提案のみ
↓
Review Changes で差分確認
↓
必要部分のみ手動Accept
↓
動作確認
↓
作業終了時バックアップ作成

## アプリ公開情報 (Branding)
- **入り口URL**: https://area-management.github.io/posting-map-system/
- **アイコン画像**: icon180-v2.png (180x180 PNG / 白背景オレンジP)
- **キャッシュバスター**: v=70 (アイコン変更時はここを上げる)
- **運用ルール**: 配布員には上記GitHubのURLを伝え、「ホーム画面に追加」してから起動するよう案内する。
