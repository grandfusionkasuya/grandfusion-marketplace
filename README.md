# GrandFusion Marketplace

GrandFusion の人格AIスキルを各端末へ配布・更新するためのプラグインマーケットプレイスです。

## 中身

- `grandfusion` プラグイン … 統括ありさ（/arisa）＋各PM（/beaute・/aru・/maker・/monet）＋フリーランス共通テンプレ（/freelance）

## 各端末でのインストール手順（初回のみ）

1. このリポジトリを GitHub に公開する（例: `https://github.com/<アカウント>/grandfusion-marketplace`）。
2. Cowork（または Claude Code）で、このリポジトリをマーケットプレイスとして追加する。
3. 一覧から `grandfusion` プラグインをインストールする。これで `/arisa` `/beaute` `/aru` `/maker` `/monet` `/freelance` が使えるようになる。

## 更新のしかた

1. 各スキルの中身（`plugins/grandfusion/skills/<名前>/SKILL.md`）と Notion の人格ページを更新する。
2. GitHub に push する。
3. 各端末はマーケットプレイスを更新（再同期）すれば最新版が反映される。`.skill` を手で入れ直す必要はない。

マスター（人間が読む正）は Notion、配布の正はこの GitHub リポジトリ、という役割分担です。
