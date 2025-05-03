import streamlit as st

st.set_page_config(
    page_title="Home",   # ブラウザのタブに表示されるタイトル
    page_icon=":material/home:",            # タブの左に表示されるアイコン（絵文字やURLでもOK）
    layout="centered",       # レイアウト（"centered" または "wide"）
    initial_sidebar_state="collapsed"  # サイドバーの初期状態（"auto", "expanded", "collapsed"）
)
st.title("トップページ")
if st.button("出演数集計アプリへ"):
    st.switch_page("pages/name_list_processor.py")
