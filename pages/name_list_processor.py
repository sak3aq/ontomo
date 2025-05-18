import streamlit as st
import pandas as pd
import difflib
import jaconv
from collections import Counter

st.set_page_config(
    page_title="list_processor",   # ブラウザのタブに表示されるタイトル
    page_icon="",            # タブの左に表示されるアイコン（絵文字やURLでもOK）
    layout="centered",       # レイアウト（"centered" または "wide"）
    initial_sidebar_state="collapsed"  # サイドバーの初期状態（"auto", "expanded", "collapsed"）
)




st.title("出演回数集計アプリ")

# 公式名簿の読み込み（1列すべてに名前があると仮定）
up_file = st.file_uploader("名簿のExcelファイルをアップロードしてください", type=["xlsx"])
if up_file:
    try:
        meibo_df = pd.read_excel(up_file)
        official_names = []
        for col in meibo_df.columns:
            for val in meibo_df[col]:
                if pd.notna(val):
                    official_names.append(str(val).strip())
    
        official_names_hira = [jaconv.kata2hira(jaconv.z2h(name, kana=True)) for name in official_names]
    
        # 表記揺れ補正関数
        def find_best_match(raw_input, threshold=0.3):
            input_hira = jaconv.kata2hira(jaconv.z2h(raw_input, kana=True))
            scores = []
            for i, hira_name in enumerate(official_names_hira):
                score = difflib.SequenceMatcher(None, input_hira, hira_name).ratio()
                if score >= threshold:
                    scores.append((official_names[i], score))
            scores.sort(key=lambda x: x[1], reverse=True)
            return scores[0][0] if scores else raw_input
    
        # データアップロード
        uploaded_file = st.file_uploader("集計対象のExcelファイルをアップロードしてください", type=["xlsx"])
    
        if uploaded_file:
            df = pd.read_excel(uploaded_file)
    
            # E2:K200 の範囲から名前を抽出
            name_range = df.iloc[1:200, 4:11]  # 行番号: 1~199 (E2:K200), 列番号: 4~10 (E-K)
    
            # 名前リストの作成
            raw_name_list = []
            for row in name_range.values:
                # 各行の名前を補正
                corrected_names_in_row = [find_best_match(str(val).strip()) for val in row if pd.notna(val)]
    
                # 行ごとに重複を除外
                unique_names_in_row = list(set(corrected_names_in_row))
                raw_name_list.extend(unique_names_in_row)
    
    
            # 出現回数をカウント
            counts = Counter(raw_name_list)
    
            # 名前ごとに代（generation）を調べて、先にデータ構造を作る
            annotated_counts = []  # [代, 名前, 出現回数] のリスト
    
            for name, count in counts.items():
                found = False
                for col in meibo_df.columns:
                    values = meibo_df[col].dropna().astype(str).str.strip()
                    if name in values.values:
                        generation = col  # 代として列の1行目（col名）を使う前提
                        annotated_counts.append([generation, name, count])
                        found = True
                        break
                if not found:
                    annotated_counts.append(["error", name, count])
                    
            def parse_generation(gen):
                try:
                    return int(gen)
                except:
                    return 999  # "error"など数値変換できないものは後回しにする
    
            # 回数降順 → 代昇順でソート
            sorted_counts = sorted(
                annotated_counts,
                key=lambda x: (-x[2], parse_generation(x[0]))
            )
    
    
            # 集計結果をデータフレームに変換
            result_df = pd.DataFrame(sorted_counts, columns=["代","名前", "出現回数"])
    
            # データフレームを表示
            st.subheader("補正後の名前ごとの出現回数")
            st.write(result_df)
    except Exception as e:
        st.write(f"名簿ファイルの読み込み中にエラーが発生しました: {e}")
