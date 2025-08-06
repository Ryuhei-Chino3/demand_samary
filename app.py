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
    """自動で文字コード判定して読み込む"""
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
        file_name = file.name.lower()

        st.write(f"--- ファイル名: {file.name} ---")

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

        st.write("アップロード直後のデータ（先頭5行）")
        st.dataframe(df.head())

        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            st.warning(f"{file_name} に必要な列がありません: {missing}")
            continue

        try:
            # year, month, date はint型に変換してゼロ埋め文字列作成
            df["year"] = df["year"].astype(int)
            df["month"] = df["month"].astype(int)
            df["date"] = df["date"].astype(int)

            month_str = df["month"].astype(str).str.zfill(2)
            date_str = df["date"].astype(str).str.zfill(2)

            # datetime生成（timeは文字列で想定）
            df["datetime"] = pd.to_datetime(
                df["year"].astype(str) + "-" + month_str + "-" + date_str + " " + df["time"].astype(str),
                format="%Y-%m-%d %H:%M:%S",
                errors="coerce"
            )

            st.write("datetime列の中身（先頭5行）")
            st.dataframe(df["datetime"].head())

            null_dt_count = df["datetime"].isna().sum()
            st.write(f"datetime生成失敗（NaT）の行数: {null_dt_count}")

            df = df.dropna(subset=["datetime"])  # datetime生成失敗行は除外

            dfs.append(df)
        except Exception as e:
            st.error(f"{file_name} の datetime 生成に失敗しました: {e}")

    if not dfs:
        st.warning("有効なデータが読み込めませんでした。")
    else:
        value_cols = ["買電電力量(kWh)", "売電電力量(kWh)", "発電電力量(kWh)", "消費電力量(kWh)"]

        # 複数データのdatetimeでouterマージして平均化
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
                # suffix付き含むすべての該当列の平均を計算
                avg_cols = [c for c in merged_df.columns if c.startswith(col)]
                avg_df[col] = merged_df[avg_cols].mean(axis=1)

            # 30分値シート用：1件目のファイル構造をベースに平均値を合成
            base_df = dfs[0].copy()
            output_cols = base_df.columns.tolist()

            final_30min_df = pd.merge(
                base_df.drop(columns=value_cols),
                avg_df[["datetime"] + value_cols],
                on="datetime",
                how="left"
            )
            final_30min_df = final_30min_df[output_cols]

            # サマリーシート用：全ファイルを結合して年・月別合計
            summary_df = pd.concat(dfs)
            summary_df["year"] = summary_df["year"].astype(int)
            summary_df["month"] = summary_df["month"].astype(int)
            summary_grouped = summary_df.groupby(["year", "month"])[value_cols].sum().reset_index()

            # Excel出力
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
else:
    st.info("ファイルをアップロードしてください。")
