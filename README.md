# word2ppt

高校日本史の授業プリント作成支援ツール。
Claude.ai Project が Word 原稿から生成した「構成Markdown」を PPTX に変換する。

Streamlit Web アプリ + CLI スクリプトの2つの形で利用可能。

## 妻の運用フロー（Web版）

```
1. Claude.ai の「日本史授業プリント」プロジェクトで Word 添付 →「変換お願い」
2. Claudeが出力した構成Markdownをコピー
3. Streamlit Cloud にデプロイしたアプリ（URL）を開く
4. MDを貼り付け → 「PPTXに変換」→ ダウンロード
5. PowerPoint で開く（**黄色字のクリックアニメは自動付与済み**）
6. 必要に応じて赤字や画像を貼り付け、授業で使用
```

## ディレクトリ構成

```
word2ppt/
├── app.py                 ← Streamlit Web UI
├── convert.py             ← MD→PPTX 変換ロジック（CLI も兼ねる）
├── requirements.txt       ← Streamlit Cloud 用の依存定義
├── template/
│   └── base_template.pptx ← 妻の自作テンプレ
├── samples/               ← ローカル開発用のサンプル（.gitignore で除外）
└── patterns/              ← ローカル分析データ（.gitignore で除外）
```

## ローカル実行（井上のMac）

### CLI として使う

```bash
cd "個人開発PJ/word2ppt"
python3 convert.py samples/NO1_構成.md samples/NO1_output.pptx
```

### Streamlit を起動

```bash
pip install -r requirements.txt
streamlit run app.py
# → http://localhost:8501 が自動で開く
```

## Streamlit Cloud デプロイ手順（初回のみ）

### 事前準備
- GitHubアカウント
- [Streamlit Community Cloud](https://streamlit.io/cloud) にGitHubでサインアップ（無料）

### 手順

1. **GitHubに公開リポジトリを作成**
   ```bash
   cd "個人開発PJ/word2ppt"
   git init
   git add .
   git commit -m "initial commit"
   # GitHubでブラウザから新規リポジトリ作成（例: word2ppt）
   git remote add origin https://github.com/<YOUR_GH_USER>/word2ppt.git
   git branch -M main
   git push -u origin main
   ```
   
   もしくは `gh` CLI が入っていれば：
   ```bash
   brew install gh
   gh auth login
   gh repo create word2ppt --public --source . --push
   ```

2. **Streamlit Cloud でデプロイ**
   - https://share.streamlit.io/ にアクセス
   - 「New app」→ 自分のGitHubリポジトリ `word2ppt` を選択
   - Main file path: `app.py`
   - Deploy ボタン
   - 数分でビルド完了 → `https://<your-app>.streamlit.app/` にアクセス可能

3. **妻にURLを共有**
   - ブックマーク化しておくと、次回からすぐ開ける
   - 数分アクセスがないとスリープするため、初回起動は10-30秒待つ

### 更新時

`convert.py` や `app.py` を変更したら：

```bash
git add -u
git commit -m "update converter"
git push
# Streamlit Cloud は push を検知して自動で再デプロイ（1-2分）
```

## 仕様

### スライドレイアウト
- 4:3 スライド（33.87 × 19.05 cm）
- 背景: 黒
- フォント: 游ゴシック・太字

### テキスト
- タイトル: 60pt 白
- 本文: **38pt 白固定**（長文は自動で折り返し・auto_size=NONEで縮小無効化）
- **黄色字（穴埋めの答え）: 44pt 黄 #FFFF00 固定**
  - 右サイドに縦並び（本文とオーバーラップ許容）
  - **クリックで上から順次出現アニメ自動付与**（spec: presetID=1/presetClass=entr/Appear）

### 自動スライド分割
物理上限ベースの厳格な閾値で1スライドに収まらない場合は自動分割：
- 本文の折り返し行数 > 10
- 本文の総文字数 > 200
- 1行が42字超
- 黄色字が7個超
- 意味区切り（◎・※・・・括弧）で切り、続きは「（続き）」と見出しを自動追加

## 仕様メモ

1. **赤字（強調）はツール側で生成しない** – 妻が PowerPoint で必要に応じて手動追加
2. **黄色字の位置は右サイド固定** – 完成形の「本文内に重ね配置」は再現しない
3. **画像は挿入しない** – 妻の運用「画像は後で貼る」に合わせる
4. **ピンク字（#FF00FF / #FF66FF）は出力しない** – Word原稿の教師メモは除外

## 調整可能な設定

`convert.py` の先頭定数で調整可能：

| 定数 | デフォルト | 意味 |
|---|---|---|
| `BODY_FONT_PT_DEFAULT` | 38 | 本文 |
| `YELLOW_FONT_PT_DEFAULT` | 44 | 黄色字 |
| `TITLE_FONT_PT` | 60 | タイトル |
| `MAX_BODY_LINES_PER_SLIDE` | 10 | 本文の折返行数上限（超で分割） |
| `MAX_BODY_TOTAL_CHARS` | 200 | 本文総文字数上限（超で分割） |
| `MAX_LINE_HARD_CHARS` | 42 | 1行の文字数上限（超で分割） |
| `MAX_YELLOW_PER_SLIDE` | 7 | 黄色字個数上限（超で分割） |
| `CHARS_PER_LINE_38PT` | 22 | 38pt行折返し計算用（実物27字想定の安全側） |
| `FONT_NAME` | 游ゴシック | 全体フォント |
