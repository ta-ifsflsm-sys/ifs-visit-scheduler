# IFS Japan Visit Scheduler — セットアップガイド

## 構成
- **フロントエンド**: React + Vite
- **DB / Realtime**: Supabase (PostgreSQL)
- **ホスティング**: Vercel
- **カレンダー連携**: .ics ファイルダウンロード（Outlook / Teams に取り込み可能）

---

## 手順 1 — Supabase プロジェクトを作成

1. https://supabase.com にアクセスしてサインアップ（無料）
2. 「New Project」を作成
3. ダッシュボードの **SQL Editor** を開く
4. `schema.sql` の内容を全てコピーして実行

---

## 手順 2 — Supabase の API キーを取得

Supabase ダッシュボード → **Settings** → **API**

| 項目 | 使う値 |
|------|--------|
| Project URL | `VITE_SUPABASE_URL` に設定 |
| `anon` public key | `VITE_SUPABASE_ANON_KEY` に設定 |

---

## 手順 3 — ローカルで動かす

```bash
# リポジトリをクローン後
npm install

# .env.local を作成
cp .env.example .env.local
# → .env.local を開いて Supabase の値を入力

npm run dev
# http://localhost:5173 で起動
```

---

## 手順 4 — Vercel にデプロイ

```bash
# Vercel CLI を使う場合
npm i -g vercel
vercel

# または GitHub にプッシュして Vercel ダッシュボードで Import
```

**Vercel の Environment Variables に追加：**

| Key | Value |
|-----|-------|
| `VITE_SUPABASE_URL` | Supabase の Project URL |
| `VITE_SUPABASE_ANON_KEY` | Supabase の anon key |

---

## 手順 5 — チームに URL を共有

デプロイ後の URL（例: `https://ifs-visit-scheduler.vercel.app`）を
チームメンバーに送るだけで全員が使えます。

週の URL は `?week=2025-07-07` 形式で保存されるため、
特定の週のリンクを直接共有することも可能です。

---

## カレンダー招待（Teams / Outlook）の使い方

1. スケジューラーでミーティングをクリック
2. 「📅 カレンダーに追加 (.ics)」ボタンをクリック
3. ダウンロードされた `.ics` ファイルをダブルクリック
4. Outlook / Teams が自動で開いて招待が作成される
5. 参加者を確認して「送信」

> Microsoft 管理者権限は不要です。各自のアカウントから送信できます。

---

## Realtime 同期について

Supabase の WebSocket を使っているため、
複数人が同時に操作しても **数秒以内に全員の画面に反映**されます。
ポーリングや手動リフレッシュは不要です。

---

## ファイル構成

```
ifs-visit-scheduler/
├── schema.sql          ← Supabase で実行する SQL
├── .env.example        ← 環境変数テンプレート
├── package.json
├── vite.config.js
├── index.html
└── src/
    ├── main.jsx        ← エントリポイント
    ├── index.css       ← グローバルスタイル
    ├── App.jsx         ← メインコンポーネント（全ロジック）
    ├── lib/
    │   └── supabase.js ← Supabase クライアント
    └── utils/
        └── ics.js      ← .ics ファイル生成
```
