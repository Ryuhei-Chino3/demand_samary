import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.title("PPAデータサマリー生成アプリ")

uploaded_files = st.file_uploader("Excelファイルをアップロード（複数選択可）", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    dfs = []

    for file in uploaded_files:
        df = pd.read_excel(file)
        df["datetime"] = pd.to_datetime(df["date"].astype(str) + " " + df["time"].astype(str), errors="coerce")
        dfs.append(df)

    # 有効なデータのみに限定
    dfs = [df.dropna(subset=["datetime"]) for df in dfs]

    # 共通キーでマージ（datetime）
    merged_df = dfs[0][["datetime"] + list(dfs[0].columns)]
    for df in dfs[1:]:
        merged_df = pd.merge(
            merged_df,
            df[["datetime", "買電電力量(kWh)", "売電電力量(kWh)", "発電電力量(kWh)", "消費電力量(kWh)"]],
            on="datetime",
            how="outer",
            suffixes=("", "_dup"),
        )

    # 平均を計算
    value_cols = ["買電電力量(kWh)", "売電電力量(kWh)", "発電電力量(kWh)", "消費電力量(kWh)"]
    avg_df = merged_df.copy()
    for col in value_cols:
        avg_cols = [c for c in merged_df.columns if c.startswith(col)]
        avg_df[col] = merged_df[avg_cols].mean(axis=1)

    # 最終出力用の「30分値」シート
    output_cols = dfs[0].columns.tolist()
    avg_df = avg_df[["datetime"] + value_cols]
    base_df = dfs[0].copy()
    base_df["datetime"] = pd.to_datetime(base_df["date"].astype(str) + " " + base_df["time"].astype(str), errors="coerce")
    final_30min_df = pd.merge(base_df.drop(columns=value_cols), avg_df, on="datetime", how="left")
    final_30min_df = final_30min_df[output_cols]  # 列順を元に戻す

    # 「サマリー」シート作成
    summary_df = pd.concat(dfs)
    summary_df["year"] = summary_df["year"].astype(int)
    summary_df["month"] = summary_df["month"].astype(int)

    summary_grouped = summary_df.groupby(["year", "month"])[value_cols].sum().reset_index()

    # 出力Excel作成
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
