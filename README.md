# 年収UP研究所 転職サービスガイド

X・公式LINEから案内する公開ページ。A8.netの広告リンクはこのページにだけ置く（LINEのトークやDMには貼らない：A8.netの禁止事項）。リンクは index.html の SERVICES[].link に入れる。

## 診断ページ（shindan/）の公開先
- 本番：https://nenshuup-lab.pages.dev/shindan/ （Cloudflare Pages、プロジェクト nenshuup-lab。2026-11-12 作成。Git連携なしの直接アップロード）
- 修正したら、index.html と shindan/ だけを入れたフォルダを `npx wrangler pages deploy <フォルダ> --project-name nenshuup-lab --branch main` で載せ直す（push だけでは Cloudflare 側は更新されない）。
- GitHub Pages（https://tomoya09225.github.io/nenshuup-guide/shindan/）も push で更新され、予備として残っている。
