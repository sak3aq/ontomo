import streamlit as st
import openpyxl
import re
from openpyxl.styles import PatternFill, Font
import io

st.set_page_config(
    page_title="Schedule_Sync",   # ブラウザのタブに表示されるタイトル
    page_icon="",            # タブの左に表示されるアイコン（絵文字やURLでもOK）
    layout="centered",       # レイアウト（"centered" または "wide"）
    initial_sidebar_state="collapsed"  # サイドバーの初期状態（"auto", "expanded", "collapsed"）
)



# --- 2. メインの処理を関数にまとめる ---
# Streamlitアプリでは、メインのロジックを関数化すると管理しやすくなります。
def process_excel(uploaded_file):
    try:
        # アップロードされたファイルをメモリ上で読み込む
        wb = openpyxl.load_workbook(uploaded_file)
        # "出演順"シートを選択する
        st_syutuen = wb["出演順"]
        # "名簿"シートを選択する
        st_meibo = wb["名簿"]
    except KeyError as e:
        st.error(f"エラー: シート '{e.args[0]}' がExcelファイル内に見つかりません。")
        st.info("ファイルには '出演順' と '名簿' の2つのシートが必要です。")
        return None
    except Exception as e:
        st.error(f"ファイルの読み込み中にエラーが発生しました: {e}")
        return None

    # --- データの読み込みとリスト作成 ---
    st.info("1/5: シートからデータを読み込んでいます...")
    list_of_lists = [list(row) for row in st_syutuen.iter_rows(values_only=True)]
    
    # 名簿リストの作成
    circle_members_data = [list(row) for row in st_meibo.iter_rows(values_only=True)]
    circle_member_list = []
    for member_i in circle_members_data:
        for member_j in member_i:
            if member_j is not None:
                member_j = re.sub(r'\s', '', str(member_j))
                circle_member_list.append(member_j)

    # 出演順からユニークなメンバーリストを作成
    list_all_menber = []
    for i in list_of_lists[1:]: # ヘッダー行をスキップ
        if i is None: continue
        for j in i:
            if j is None: continue
            j_cleaned = re.sub(r'\s', '', str(j))
            list_all_menber.append(j_cleaned)
    list_all_menber = sorted(list(set([s for s in list_all_menber if s])))

    # --- 新しいシートの準備 ---
    st.info("2/5: 結果を出力するための 'After' シートを準備しています...")
    st_name = "After"
    if st_name in wb.sheetnames:
        del wb[st_name]
    st_after = wb.create_sheet(title=st_name)

    # --- 元のデータをコピー ---
    for row_data in list_of_lists:
        st_after.append(row_data)

    # --- マトリクスの書き込み ---
    st.info("3/5: メンバーのヘッダーを作成し、名簿と照合しています...")
    # ヘッダー（メンバー名）の書き込みと名簿外メンバーのチェック
    red_font = Font(color='FFFF0000') # 赤色のフォントを定義
    for i, member_name in enumerate(list_all_menber):
        cell = st_after.cell(row=1, column=23 + i, value=member_name)
        if member_name not in circle_member_list:
            cell.font = red_font
    
    st.info("4/5: 出演・シフト状況を書き込み、ブッキングをチェックしています...")
    # 出演・シフト・楽器の状況を書き込み
    booking_font = Font(color='FFFF0000') # ブッキング用の赤色フォント
    for row_index, row_data in enumerate(list_of_lists):
        if row_data is None: continue

        cleaned_row_values_syutuen = {re.sub(r'\s', '', str(cell)) for cell in row_data[:7] if cell}
        cleaned_row_values_busyo = {re.sub(r'\s', '', str(cell)) for cell in row_data[7:15] if cell}
        cleaned_row_values_gakki = {re.sub(r'\s', '', str(cell)) for cell in row_data[15:] if cell}
        
        for col_index, member_name in enumerate(list_all_menber):
            target_cell = st_after.cell(row=row_index + 1, column=23 + col_index)
            
            is_booked = False
            # 出演チェック
            if member_name in cleaned_row_values_syutuen:
                target_cell.value = "出演"
                is_booked = True
            
            # 部署シフトチェック
            if member_name in cleaned_row_values_busyo:
                if is_booked:
                    target_cell.value = "ブッキング"
                    target_cell.font = booking_font
                else:
                    target_cell.value = "部署シフト"
                    is_booked = True

            # 楽器シフトチェック
            if member_name in cleaned_row_values_gakki:
                if is_booked:
                    target_cell.value = "ブッキング"
                    target_cell.font = booking_font
                else:
                    target_cell.value = "楽器シフト"

    st.info("5/5: 処理が完了しました。")
    return wb

# --- 1. StreamlitのUI部分 ---
st.title("出演・シフト管理表 作成ツール")
st.write("「出演順」と「名簿」の2つのシートを含むExcelファイルをアップロードしてください。")
st.write("スケジュールのブッキングチェックと、名簿にない人物の洗い出しを自動で行います。")

# ファイルアップローダー
uploaded_file = st.file_uploader(
    "Excelファイルをアップロード",
    type=['xlsx'],
    help="「出演順」と「名簿」という名前のシートが含まれている .xlsx ファイルを選択してください。"
)

if uploaded_file is not None:
    # ファイルがアップロードされたら、処理を実行
    with st.spinner('Excelファイルを処理中です...'):
        processed_wb = process_excel(uploaded_file)

    if processed_wb:
        # --- 3. 結果のダウンロード ---
        # 処理済みのExcelファイルをメモリ上のバイナリデータとして保存
        output = io.BytesIO()
        processed_wb.save(output)
        output.seek(0)
        
        st.success("✔ 処理が完了しました！")
        
        # ダウンロードボタン
        st.download_button(
            label="結果のExcelファイルをダウンロード",
            data=output,
            file_name=f"{uploaded_file.name.split('.')[0]}_After.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )