"""Streamlit Web UI for word2ppt converter.

妻がブラウザから構成MDを貼り付け → PPTXをダウンロードする用のアプリ。
Streamlit Community Cloud にデプロイする前提。

ローカル実行:
    streamlit run app.py
"""
from io import BytesIO
from datetime import datetime

import streamlit as st

from convert import convert_md_to_pptx


st.set_page_config(
    page_title="日本史授業プリント変換",
    page_icon="📚",
    layout="centered",
)


st.title("📚 日本史授業プリント変換")
st.caption("Claude.ai で作った構成Markdownを、穴埋めアニメーション付きPPTXに変換します")


with st.expander("📖 使い方", expanded=False):
    st.markdown(
        """
        1. [Claude.ai](https://claude.ai/) の「日本史授業プリント」プロジェクトで Word原稿を添付し、「変換お願い」と入力
        2. Claudeが出力した **構成Markdown全体** をコピー（---で挟まれた部分すべて）
        3. このページの下のテキストエリアに貼り付け
        4. ファイル名を入力して「PPTXに変換」ボタンを押す
        5. 「ダウンロード」ボタンからファイルを保存
        6. PowerPointで開く
        7. **黄色字のテキストボックスをまとめて選択** → 「アニメーション」タブ → 「フェード」→「開始：クリック時」（30秒で完成）
        8. 画像を貼り付けて、授業で使用
        """
    )


st.divider()


md_input = st.text_area(
    "構成Markdown",
    height=400,
    placeholder=(
        "---\n"
        "## slide1: タイトル\n"
        "### 本文（黒・太字）\n"
        "高２　日本史探究NO.１　旧石器・縄文時代\n"
        "---\n"
        "## slide2: 人類の進化\n"
        "...\n"
    ),
    help="Claude.ai が出力した構成Markdownをそのまま貼り付けてください",
)


uploaded_md = st.file_uploader(
    "または .md / .txt ファイルをアップロード",
    type=["md", "txt"],
    help="ファイルをアップロードすると、上のテキストエリアの内容より優先されます",
)


if uploaded_md is not None:
    md_input = uploaded_md.read().decode("utf-8")
    st.info(f"✅ `{uploaded_md.name}` を読み込みました（{len(md_input)} 文字）")


default_name = f"授業プリント_{datetime.now().strftime('%Y%m%d')}"
filename = st.text_input(
    "出力ファイル名（拡張子 .pptx は自動で付きます）",
    value=default_name,
)


can_convert = bool(md_input and md_input.strip())

if st.button("🚀 PPTXに変換", type="primary", disabled=not can_convert, use_container_width=True):
    try:
        with st.spinner("変換中..."):
            buf = BytesIO()
            result = convert_md_to_pptx(md_input, buf)
            buf.seek(0)

        split_added = result["expanded"] - result["original"]
        msg = f"✅ 変換完了：{result['expanded']}枚生成"
        if split_added:
            msg += f"（うち{split_added}枚は自動でスライド分割）"
        st.success(msg)

        safe_name = filename.strip() or default_name
        if not safe_name.endswith(".pptx"):
            safe_name += ".pptx"

        st.download_button(
            label="📥 PPTXをダウンロード",
            data=buf,
            file_name=safe_name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )
    except ValueError as e:
        st.error(f"❌ 入力エラー: {e}")
    except Exception as e:
        st.error(f"❌ 変換中にエラーが発生しました: {e}")
        with st.expander("詳細"):
            st.exception(e)


st.divider()
st.caption(
    "変換仕様：本文=游ゴシック38pt・白、黄色字=44pt、強調=赤、背景=黒。"
    "スライドに収まらない場合は自動で「（続き）」に分割されます。"
)
