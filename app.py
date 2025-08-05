import streamlit as st
import pandas as pd
import io

st.title("PPAデータサマリー生成アプリ（CSV / Excel 対応）")

uploaded_files = st.file_uploader(
    "CSV または Excelファイルをアップロード（複数可）",
    type=["xlsx", "csv"],
    accept_multiple_files=True
)

def try_read_csv(file):
    encodings = ["utf-8", "cp932", "iso-8859-1"]
    for enc in encodings:
        try:
            file.seek(0)
            return pd.read_csv(file, encoding=enc)
        except Exception:
            continue
    raise ValueError("読み込み可能なエンコーディングではありません")

required_cols = ["year", "month", "date", "time", "買電電力量(kWh)", "売電電力量(kWh)", "発電電力量(kWh)", "消費電力量(kWh)"]

if uploaded_files:
    dfs = []
    for file in uploaded_files:
        filename = file.name.lower()
        try:
            if filename.endswith(".xlsx"):
                df = pd.read_excel(file)
            elif filename.endswith(".csv"):
                df = try_read_csv(file)
            else:
                st.warning(f"未対応ファイル形式: {filename}")
                continue
        except Exception as e:
            st.error(f"{filename} の読み込み中にエラーが発生しました: {e}")
            continue

        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            st.warning(f"{filename} に必要な列がありません: {missing}")
            continue

        # datetime列を作る
        df["datetime"] = pd.to_datetime(df["date"].astype(str) + " " + df["time"].astype(str), errors="coerce")
        df = df.dropna(subset=["datetime"])

        dfs.append(df)

    if not dfs:
        st.warning("有効なデータが読み込めませんでした。")
    else:
        value_cols = ["買電電力量(kWh)", "売電電力量(kWh)", "発電電力量(kWh)", "消費電力量(kWh)"]

        # datetimeをインデックスにして concat し、列名をファイル名などで一意にしながら平均を計算
        df_list_for_avg = []
        for i, df in enumerate(dfs):
            tmp = df.set_index("datetime")[value_cols].copy()
            # 列名をユニーク化
            tmp.columns = [f"{col}_file{i}" for col in value_cols]
            df_list_for_avg.append(tmp)

        combined = pd.concat(df_list_for_avg, axis=1)
        avg_df = pd.DataFrame(index=combined.index)
        for col in value_cols:
            cols = [c for c in combined.columns if c.startswith(col)]
            avg_df[col] = combined[cols].mean(axis=1)

        avg_df = avg_df.reset_index()

        # 30分値シートの元データは最初のファイルのデータをベースに平均値を結合
        base_df = dfs[0].copy()
        base_df = base_df.drop(columns=value_cols)
        final_30min_df = pd.merge(base_df, avg_df, on="datetime", how="left")

        # サマリーシート：すべてのファイルの値を縦に連結して年・月毎の合計を算出
        summary_df = pd.concat(dfs, ignore_index=True)
        summary_df["year"] = summary_df["year"].astype(int)
        summary_df["month"] = summary_df["month"].astype(int)
        summary_grouped = summary_df.groupby(["year", "month"])[value_cols].sum().reset_index()

        # Excel出力
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_30min_df.to_excel(writer, index=False, sheet_name="30分値")
            summary_grouped.to_excel(writer, index=False, sheet_name="サマリー")

        st.success("✅ 集計完了！下記からダウンロードしてください。")
        st.download_button(
            label="📥 PPAデータサマリー.xlsx をダウンロード",
            data=output.getvalue(),
            file_name="PPAデータサマリー.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
