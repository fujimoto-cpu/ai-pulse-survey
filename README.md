# AI Pulse Survey — 月次アンケート

KONNEKT INTERNATIONAL 社員向けのAI活用 月次アンケートフォーム。

## 公開URL

https://fujimoto-cpu.github.io/ai-pulse-survey/

（社内限定共有・noindex設定）

## 構成

```
ブラウザ（index.html / Vanilla JS）
    ↓ 隠しiframe経由のGETリクエスト（CORS回避）
Google Apps Script Web App
    ↓
Google Sheets（KONNEKT_AIPulse_データ）
```

## 質問項目

| # | カテゴリ | 内容 |
|---|---------|------|
| 0 | 回答者 | 名前選択（社員17名） |
| 1 | 時間削減 | 削減時間（時間/月） |
| 2 | 時間削減 | 1日あたりの業務時間変化 |
| 3 | 新規創出 | 新規成果物件数 |
| 4 | 新規創出 | 主な業務・用途（複数選択） |
| 5 | AI浸透度 | 利用頻度 |
| 6 | AI浸透度 | 使いやすさ（1-5） |
| 7 | 困りごと | 困っていること |
| 8 | 困りごと | 自由記述 |
| 9 | 振り返り | 活用事例（任意） |
| 10 | 振り返り | 改善要望（任意） |

## デプロイ

GitHub Pages 自動デプロイ。`main` ブランチ push で即反映。

## 関連プロジェクト

- [ai-pulse](https://github.com/fujimoto-cpu/ai-pulse) — 全社ダッシュボード本体
