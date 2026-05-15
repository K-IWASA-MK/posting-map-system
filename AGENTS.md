# 開発理念 (Core Principle)
- **究極のテンプレート**: 特定の地区に依存しない、汎用的なポスティング管理システムの基盤を構築する。
- **動的設計**: データ構造の変化に自動対応し、ハードコーディングを徹底排除する。

# AI行動指針 (Action Policy)
- **🚨 承認なき実行の絶対禁止**: AIは提案のみを行い、岩佐さんの明確な「承認(Yes/OK)」なしにいかなるファイル操作（編集・削除・適用）も実行しない。
- **Accept All の禁止**: 大量修正時も差分(Diff)を明示し、小分けにして承認を得ること。勝手な一括適用は「暴走」とみなす。
- **30秒ルール**: 思考プロセスが30秒を超えたら一度中断し、確認を仰ぐ。

# テクニカル・ガードレール (Technical Guardrails)
- **GAS構文の保護**: `<?!= include(...) ?>` は正常構文として扱い、修正対象にしない。
- **重要領域の死守**: `doGet()`、`HtmlService`、`index.html` の構造変更は「二重の確認」を必須とする。
- **開発フロー**: 修正前に必ず提案し、作業終了時はバックアップ（Git/clasp）を作成する。

# UI・ブランド定義 (UI & Branding)
- **設計思想**: モバイル優先。高齢者でも迷わない巨大な数字とシンプルUI。
- **運用ルール**: URLはGitHub Pagesを使用し、アイコン変更時はキャッシュバスター（v=70）を更新する。

# Frontend / Backend Architecture

- Frontend:
  GitHub Pages (PWA)

- Backend:
  Google Apps Script API

- Communication:
  fetch(JSON)

- GAS role:
  API only

- GitHub role:
  UI only

- GAS and GitHub are independent systems.

- Synchronization required:
  API names
  JSON structure
  data schema

- Synchronization NOT required:
  HTML
  CSS
  UI code
  GAS internal logic

# Forbidden

- HtmlService
- GAS UI rendering
- script.google.com redirects
- meta refresh redirect
- window.location.href to GAS

# GAS Response Rule

Always return JSON only.

Example:

return ContentService
  .createTextOutput(JSON.stringify(data))
  .setMimeType(ContentService.MimeType.JSON);

# Deployment Rule

Frontend changes:
git push origin main

GAS changes:
clasp push
clasp deploy -i <deployment_id> (IDを固定して更新することを推奨)

GitHub and GAS deploy separately.

AI may prepare commits locally (git add, git commit).

Human executes:
- git push
- production deploy
- release operations

Never push (Strictly enforced for AI):
- backup files
- old files
- experimental files
- temporary files


# Stability Rules

Do NOT refactor working API logic unless explicitly instructed.

Preserve existing:
- fetch structure
- API response format
- JSON keys
- doGet/doPost behavior

Do NOT optimize or simplify working GAS logic automatically.

Priority:
1. Stability
2. Compatibility
3. Existing behavior
4. Optimization

Avoid breaking existing frontend communication.

# UI DESIGN SYSTEM — POSTING MAP

## Core Style
* Background: Pure black (#000000)
* UI Direction: Minimal, Operational, Industrial, High-end PWA, Tesla-like, Apple-like, Professional terminal UI
* Design Priority: 1. Readability, 2. Simplicity, 3. Spacing, 4. Operational clarity, 5. Minimal colors

## Color Rules
### Main Colors
* Black: #000000
* Primary Blue: #2563eb
* White: #ffffff
* Secondary Text: rgba(255,255,255,0.72)
* Border: rgba(255,255,255,0.08)

## Spacing Rules
* Large vertical spacing
* Avoid crowded layouts
* Minimal information density
* Wide padding preferred

## Card Style
Use this as the default base card style:
```css
background: #0b0b0b;
border-radius: 32px;
border: 1px solid rgba(255,255,255,0.08);
box-shadow: 0 0 20px rgba(37,99,235,0.06);
```

## Button Style
* Large buttons
* Minimal decoration
* Rounded corners
* Blue solid fill
* White bold text
* Avoid: gradients, excessive glow, flashy animation

## Typography
### Main Title
* Bold
* White
* Large size
### Secondary Text
* Smaller
* Softer white
* line-height: 160%

## UI Philosophy
The system should feel like:
* Operational OS
* Logistics terminal
* Field management device
* Professional industrial application

Avoid:
* playful UI
* colorful UI
* template-like design
* crowded information
* excessive icons

Less information = stronger design.