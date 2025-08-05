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

if uploaded_files:
    dfs = []

    for file in uploaded_files:
        file_name = file.name.lower()

        try:
            if file_name.endswith(".xlsx"):
                df = pd.read_excel(file)
            elif file_name.endswith(".csv"):
                df = pd.read_csv(file, encoding="utf-8")
            else:
                st.warning(f"サポートされていない形式: {file_name}")
                continue
        except Exception as e:
            st.error(f"{file_name} の読み込み中にエラーが発生しました: {e}")
            continue

        # タイムスタンプ列の作成（例外に備えた処理）
        try:
            df["datetime"] = pd.to_datetime(
                df["date"].astype(str) + " " + df["time"].astype(str),
                errors="coerce"
            )
            dfs.append(df.dropna(subset=["datetime"]))
        except Exception as e:
            st.error(f"{file_name} の datetime 生成に失敗しました: {e}")

    if not dfs:
        st.warning("有効なデータが読み込めませんでした。")
    else:
        # マージ処理
        merged_df = dfs[0][["datetime"] + list(dfs[0].columns)]
        for df in dfs[1:]:
            merged_df = pd.merge(
                merged_df,
                df[["datetime", "買電電力量(kWh)", "売電電力量(kWh)", "発電電力量(kWh)", "消費電力量(kWh)"]],
                on="datetime",
                how="outer",
                suffixes=("", "_dup"),
            )

        # 平均値算出
        value_cols = ["買電電力量(kWh)", "売電電力量(kWh)", "発電電力量(kWh)", "消費電力量(kWh)"]
        avg_df = merged_df.copy()
        for col in value_cols:
            avg_cols = [c for c in merged_df.columns if c.startswith(col)]
            avg_df[col] = merged_df[avg_cols].mean(axis=1)

        # 30分値シート構成
        base_df = dfs[0].copy()
        base_df["datetime"] = pd.to_datetime(base_df["date"].astype(str) + " " + base_df["time"].astype(str), errors="coerce")
        output_cols = base_df.columns.tolist()

        final_30min_df = pd.merge(
            base_df.drop(columns=value_cols),
            avg_df[["datetime"] + value_cols],
            on="datetime",
            how="left"
        )
        final_30min_df = final_30min_df[output_cols]  # 列順を元に戻す

        # サマリーシート構成
        summary_df = pd.concat(dfs)
        summary_df["year"] = summary_df["year"].astype(int)
        summary_df["month"] = summary_df["month"].astype(int)
        summary_grouped = summary_df.groupby(["year", "month"])[value_cols].sum().reset_index()

        # Excel出力
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_30min_df.to_excel(writer, index=False, sheet_name="30分値")
            summary_grouped.to_excel(writer, index=False, sheet_name="サマリー")

        st.success("集計完了！以下からダウンロードしてください。")
        st.download_button(
            label="📥 PPAデータサマリー.xlsx をダウンロード",
            data=output.getvalue(),
            file_name="PPAデータサマリー.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
