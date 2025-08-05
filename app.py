import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.title("PPAデータサマリー生成アプリ（CSV / Excel 対応）")

uploaded_files = st.file_uploader(
    "CSV または Excelファイルをアップロード（複数可）",
    type=["xlsx", "csv"],
    accept_multiple_files=True
)

def try_read_csv(file):
    """自動で文字コード判定して読み込む"""
    encodings = ["utf-8", "cp932", "iso-8859-1"]
    for enc in encodings:
        try:
            return pd.read_csv(file, encoding=enc)
        except Exception:
            file.seek(0)
    raise ValueError("読み込み可能なエンコーディングではありません")

required_cols = ["year", "month", "date", "time", "買電電力量(kWh)", "売電電力量(kWh)", "発電電力量(kWh)", "消費電力量(kWh)"]

if uploaded_files:
    dfs = []
    for file in uploaded_files:
        file_name = file.name.lower()

        try:
            if file_name.endswith(".xlsx"):
                df = pd.read_excel(file)
            elif file_name.endswith(".csv"):
                df = try_read_csv(file)
            else:
                st.warning(f"未対応ファイル形式: {file_name}")
                continue
        except Exception as e:
            st.error(f"{file_name} の読み込み中にエラーが発生しました: {e}")
            continue

        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            st.warning(f"{file_name} に必要な列がありません: {missing}")
            continue

        try:
            df["datetime"] = pd.to_datetime(
                df["date"].astype(str) + " " + df["time"].astype(str),
                errors="coerce"
            )
            df = df.dropna(subset=["datetime"])
            dfs.append(df)
        except Exception as e:
            st.error(f"{file_name} の datetime 生成に失敗しました: {e}")

    if not dfs:
        st.warning("有効なデータが読み込めませんでした。")
    else:
        value_cols = ["買電電力量(kWh)", "売電電力量(kWh)", "発電電力量(kWh)", "消費電力量(kWh)"]

        # ✅ 平均化処理
        merged_df = None
        for df in dfs:
            df_small = df[["datetime"] + value_cols].copy()
            if merged_df is None:
                merged_df = df_small
            else:
                merged_df = pd.merge(
                    merged_df,
                    df_small,
                    on="datetime",
                    how="outer",
                    suffixes=("", "_dup")
                )

        if merged_df is None or merged_df.empty:
            st.error("データ結合に失敗しました。")
        else:
            avg_df = merged_df.copy()
            for col in value_cols:
                avg_cols = [c for c in merged_df.columns if c.startswith(col)]
                avg_df[col] = merged_df[avg_cols].mean(axis=1)

            # ✅ 30分値シート構成
            base_df = dfs[0].copy()
            base_df["datetime"] = pd.to_datetime(base_df["date"].astype(str) + " " + base_df["time"].astype(str), errors="coerce")
            output_cols = base_df.columns.tolist()

            final_30min_df = pd.merge(
                base_df.drop(columns=value_cols),
                avg_df[["datetime"] + value_cols],
                on="datetime",
                how="left"
            )
            final_30min_df = final_30min_df[output_cols]

            # ✅ サマリーシート構成
            summary_df = pd.concat(dfs)
            summary_df["year"] = summary_df["year"].astype(int)
            summary_df["month"] = summary_df["month"].astype(int)
            summary_grouped = summary_df.groupby(["year", "month"])[value_cols].sum().reset_index()

            # ✅ 出力ファイル作成
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                final_30min_df.to_excel(writer, index=False, sheet_name="30分値")
                summary_grouped.to_excel(writer, index=False, sheet_name="サマリー")

            st.success("✅ 集計完了！ダウンロードしてください。")
            st.download_button(
                label="📥 PPAデータサマリー.xlsx をダウンロード",
                data=output.getvalue(),
                file_name="PPAデータサマリー.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
