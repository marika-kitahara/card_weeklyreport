import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date
from pandas.api.types import is_datetime64_any_dtype as is_dt

# ページ設定
st.set_page_config(layout="wide")
st.title("📊 ウィークリーレポート集計用　期間中CV・配信費集計")

# =====================
# ユーティリティ
# =====================
def _norm_text(x) -> str:
    if x is None:
        return ""
    return str(x).replace("\r", "").replace("\n", "").strip()


def _coerce_date_series(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series(dtype="datetime64[ns]")
    if is_dt(s):
        return s
    s2 = s.copy()
    num = pd.to_numeric(s2, errors="coerce")
    num_mask = num.notna()
    if num_mask.any():
        s2.loc[num_mask] = (
            pd.to_timedelta(num[num_mask], unit="D")
            + pd.Timestamp("1899-12-30")
        )
    return pd.to_datetime(s2, errors="coerce")


def _excel_col_to_idx(col: str) -> int:
    """Excel列記号(A, Z, AA, BC...)を0始まり列番号へ変換。"""
    col = _norm_text(col).upper()
    idx = 0
    for ch in col:
        if "A" <= ch <= "Z":
            idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def _sum_col_in_period(df: pd.DataFrame, date_col: str, value_col: str,
                       start_date, end_date) -> float:
    """指定日付列が選択期間内の行について、指定値列を合計。"""
    date_idx = _excel_col_to_idx(date_col)
    value_idx = _excel_col_to_idx(value_col)
    if date_idx < 0 or value_idx < 0:
        return 0.0
    if date_idx >= len(df.columns) or value_idx >= len(df.columns):
        return 0.0

    dt = _coerce_date_series(df.iloc[:, date_idx])
    mask = (
        (dt >= pd.to_datetime(start_date)) &
        (dt <= pd.to_datetime(end_date))
    )
    return float(
        pd.to_numeric(df.loc[mask, df.columns[value_idx]], errors="coerce")
        .fillna(0)
        .sum()
    )


# =====================
# アップロード
# =====================
st.header("📑 CV・配信費集計")
cost_file = st.file_uploader("コストレポート", type="xlsx", key="cost")

# =====================
# 期間決定
# =====================
def _safe_minmax_dates_from_cost(file):
    """Listing(B列) / Affiliate(A列) の日付から初期期間を決定。"""
    try:
        xls = pd.ExcelFile(file)
        dates = []
        for sheet in xls.sheet_names:
            sl = sheet.lower()
            if "affiliate" in sl:
                date_col = "A"
            elif "listing" in sl:
                date_col = "B"
            else:
                continue
            df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")
            idx = _excel_col_to_idx(date_col)
            if idx >= len(df.columns):
                continue
            dt = _coerce_date_series(df.iloc[:, idx]).dropna()
            if not dt.empty:
                dates.extend(dt.tolist())
        if dates:
            dts = pd.to_datetime(pd.Series(dates), errors="coerce").dropna()
            if not dts.empty:
                return dts.min().date(), dts.max().date()
    except Exception:
        pass
    return None


default_start = date.today()
default_end = date.today()
if cost_file:
    mm = _safe_minmax_dates_from_cost(cost_file)
    if mm:
        default_start, default_end = mm

start_date, end_date = st.date_input(
    "集計期間",
    value=(default_start, default_end),
)
if start_date > end_date:
    st.stop()

days = (pd.to_datetime(end_date) - pd.to_datetime(start_date)).days + 1

# =====================
# 申込件数：コストレポートだけで作成
# =====================
# 指定された列定義
LISTING_MEDIA_COLS = {
    "LS_Google単体":       {"cv": "BC", "cost": "BB"},
    "LS_Google単体以外":   {"cv": "CM", "cost": "CL"},
    "LS_Googleその他":     {"cv": "DW", "cost": "DV"},
    "LS_Yahoo単体":        {"cv": "FG", "cost": "FF"},
    "LS_Yahoo単体以外":    {"cv": "GQ", "cost": "GP"},
    "LS_Microsoft単体":    {"cv": "IA", "cost": "HZ"},
    "LS_Microsoft単体以外": {"cv": "JK", "cost": "JJ"},
}


def _build_application_summary_from_cost(xls: pd.ExcelFile) -> pd.DataFrame:
    values = {
        "Affiliate": {"cv": 0.0, "cost": 0.0},
        **{media: {"cv": 0.0, "cost": 0.0} for media in LISTING_MEDIA_COLS},
    }

    for sheet in xls.sheet_names:
        sl = sheet.lower()
        if "affiliate" in sl:
            df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")
            values["Affiliate"]["cv"] += _sum_col_in_period(
                df, "A", "D", start_date, end_date
            )
            values["Affiliate"]["cost"] += _sum_col_in_period(
                df, "A", "W", start_date, end_date
            )
        elif "listing" in sl:
            df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")
            for media, cols in LISTING_MEDIA_COLS.items():
                values[media]["cv"] += _sum_col_in_period(
                    df, "B", cols["cv"], start_date, end_date
                )
                values[media]["cost"] += _sum_col_in_period(
                    df, "B", cols["cost"], start_date, end_date
                )

    def total(names, key):
        return sum(values[name][key] for name in names)

    google = ["LS_Google単体", "LS_Google単体以外", "LS_Googleその他"]
    yahoo = ["LS_Yahoo単体", "LS_Yahoo単体以外"]
    microsoft = ["LS_Microsoft単体", "LS_Microsoft単体以外"]
    tandoku = ["LS_Google単体", "LS_Yahoo単体", "LS_Microsoft単体"]
    tandoku_igai = ["LS_Google単体以外", "LS_Yahoo単体以外", "LS_Microsoft単体以外"]

    aggregate = {
        "Google": {
            "cv": total(google, "cv"),
            "cost": total(google, "cost"),
        },
        "Yahoo": {
            "cv": total(yahoo, "cv"),
            "cost": total(yahoo, "cost"),
        },
        "Microsoft": {
            "cv": total(microsoft, "cv"),
            "cost": total(microsoft, "cost"),
        },
        "単体": {
            "cv": total(tandoku, "cv"),
            "cost": total(tandoku, "cost"),
        },
        "単体以外": {
            "cv": total(tandoku_igai, "cv"),
            "cost": total(tandoku_igai, "cost"),
        },
    }

    aggregate["SEM"] = {
        "cv": aggregate["Google"]["cv"] + aggregate["Yahoo"]["cv"] + aggregate["Microsoft"]["cv"],
        "cost": aggregate["Google"]["cost"] + aggregate["Yahoo"]["cost"] + aggregate["Microsoft"]["cost"],
    }
    aggregate["ALL"] = {
        "cv": values["Affiliate"]["cv"] + aggregate["SEM"]["cv"],
        "cost": values["Affiliate"]["cost"] + aggregate["SEM"]["cost"],
    }

    rows = []

    def add_row(category, media, cv, cost):
        rows.append({
            "分類": category,
            "媒体": media,
            "CV合計": round(float(cv), 0),
            "CV日割り": round(float(cv) / days, 2),
            "合計費用": round(float(cost), 0),
        })

    # 媒体別
    add_row("Affiliate", "Affiliate", values["Affiliate"]["cv"], values["Affiliate"]["cost"])
    for media in LISTING_MEDIA_COLS:
        add_row("Listing", media, values[media]["cv"], values[media]["cost"])

    # 集約行（「その他」は廃止）
    for category in ["ALL", "SEM", "Google", "Yahoo", "Microsoft", "単体", "単体以外"]:
        add_row(category, "", aggregate[category]["cv"], aggregate[category]["cost"])

    return pd.DataFrame(rows)


# =====================
# コストレポート日別 Forecast/実績（全期間）
# =====================
daily_cost_df = None
daily_cost_df_for_excel = None


def _build_daily_cost_report_all_range(xls: pd.ExcelFile):
    sheets = []
    for s in xls.sheet_names:
        sl = s.lower()
        if "affiliate" in sl:
            sheets.append((s, "Affiliate"))
        elif "listing" in sl:
            sheets.append((s, "Listing"))
        elif "display" in sl and "nonifrs" not in sl:
            sheets.append((s, "Display"))
    if not sheets:
        return None, None

    # Affiliate Forecast配信費は V列（0始まり21）へ修正
    col_idx = {
        "Affiliate": {
            "date": _excel_col_to_idx("A"),
            "actual_afcv": _excel_col_to_idx("D"),
            "actual_cost": _excel_col_to_idx("U"),
            "fc_afcv": _excel_col_to_idx("C"),
            "fc_cost": _excel_col_to_idx("V"),
        },
        "Listing": {
            "date": _excel_col_to_idx("B"),
            "actual_afcv": _excel_col_to_idx("S"),
            "actual_cost": _excel_col_to_idx("R"),
            "fc_afcv": _excel_col_to_idx("G"),
            "fc_cost": _excel_col_to_idx("D"),
        },
        "Display": {
            "date": _excel_col_to_idx("B"),
            "actual_afcv": _excel_col_to_idx("S"),
            "actual_cost": _excel_col_to_idx("R"),
            "fc_afcv": _excel_col_to_idx("G"),
            "fc_cost": _excel_col_to_idx("D"),
        },
    }

    all_dates_collect = []

    def _read_sheet_robust(sheet_name):
        try:
            return pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl"), False
        except Exception:
            pass
        try:
            return pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl", header=None), True
        except Exception:
            return None, False

    for sheet_name, typ in sheets:
        df0, _ = _read_sheet_robust(sheet_name)
        if df0 is None or df0.empty:
            continue
        idxs = col_idx[typ]
        if idxs["date"] >= len(df0.columns):
            continue
        s_date0 = _coerce_date_series(df0.iloc[:, idxs["date"]]).dropna()
        if not s_date0.empty:
            all_dates_collect.extend(list(pd.to_datetime(s_date0).dt.floor("D")))

    if not all_dates_collect:
        return None, None

    global_min = min(all_dates_collect)
    global_max = max(all_dates_collect)
    all_days = pd.date_range(global_min, global_max, freq="D")

    def zero_series():
        return pd.Series(0.0, index=all_days)

    series_map = {}
    for status in ["Forecast", "実績"]:
        for metric in ["AFCV", "配信費"]:
            for typ in ["Listing", "Display", "Affiliate"]:
                series_map[(status, metric, typ)] = zero_series()

    for sheet_name, typ in sheets:
        df, _ = _read_sheet_robust(sheet_name)
        if df is None or df.empty:
            continue
        idxs = col_idx[typ]
        if idxs["date"] >= len(df.columns):
            continue

        s_date = _coerce_date_series(df.iloc[:, idxs["date"]])
        if s_date.dropna().empty:
            continue

        def safe_num(col_i):
            if col_i < len(df.columns):
                return pd.to_numeric(df.iloc[:, col_i], errors="coerce").fillna(0.0)
            return pd.Series(0.0, index=df.index)

        s_fc_afcv = safe_num(idxs["fc_afcv"])
        s_fc_cost = safe_num(idxs["fc_cost"])
        s_ac_afcv = safe_num(idxs["actual_afcv"])
        s_ac_cost = safe_num(idxs["actual_cost"])

        # 既存仕様を維持：Affiliate実績AFCVは×0.9
        if typ == "Affiliate":
            s_ac_afcv = s_ac_afcv * 0.9

        g = pd.DataFrame({
            "_date": pd.to_datetime(s_date).dt.floor("D"),
            "_fc_afcv": s_fc_afcv.values,
            "_fc_cost": s_fc_cost.values,
            "_ac_afcv": s_ac_afcv.values,
            "_ac_cost": s_ac_cost.values,
        })
        g = g.dropna(subset=["_date"]).groupby("_date", as_index=True).sum()
        g = g.reindex(all_days, fill_value=0.0)

        series_map[("Forecast", "AFCV", typ)] += g["_fc_afcv"]
        series_map[("Forecast", "配信費", typ)] += g["_fc_cost"]
        series_map[("実績", "AFCV", typ)] += g["_ac_afcv"]
        series_map[("実績", "配信費", typ)] += g["_ac_cost"]

    order = [
        ("Forecast", "AFCV", "Listing"), ("Forecast", "AFCV", "Display"), ("Forecast", "AFCV", "Affiliate"),
        ("Forecast", "配信費", "Listing"), ("Forecast", "配信費", "Display"), ("Forecast", "配信費", "Affiliate"),
        ("実績", "AFCV", "Listing"), ("実績", "AFCV", "Display"), ("実績", "AFCV", "Affiliate"),
        ("実績", "配信費", "Listing"), ("実績", "配信費", "Display"), ("実績", "配信費", "Affiliate"),
    ]
    data_dict = {f"{k[0]}_{k[1]}_{k[2]}": series_map[k].astype(float) for k in order}
    df_flat = pd.DataFrame(
        data_dict,
        index=series_map[("Forecast", "AFCV", "Listing")].index,
    ).reset_index()
    df_flat.rename(columns={"index": "日付"}, inplace=True)
    df_flat["日付"] = pd.to_datetime(df_flat["日付"]).dt.strftime("%Y/%m/%d")
    return df_flat, df_flat.copy()


# =====================
# コストレポート日別：発行実績
# =====================
def _build_daily_issue_cost_from_cost_report(xls: pd.ExcelFile):
    records = []

    for sheet in xls.sheet_names:
        sl = sheet.lower()

        if "listing" in sl:
            typ = "Listing"
            date_col = _excel_col_to_idx("B")
            cost_col = _excel_col_to_idx("W")
            rate = 1.0
        elif "display" in sl and "nonifrs" not in sl:
            typ = "Display"
            date_col = _excel_col_to_idx("B")
            cost_col = _excel_col_to_idx("W")
            rate = 1.0
        elif "affiliate" in sl:
            typ = "Affiliate"
            date_col = _excel_col_to_idx("A")
            cost_col = _excel_col_to_idx("L")
            rate = 0.9
        else:
            continue

        df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")
        if date_col >= len(df.columns) or cost_col >= len(df.columns):
            continue

        tmp = pd.DataFrame({
            "日付": _coerce_date_series(df.iloc[:, date_col]),
            "値": pd.to_numeric(df.iloc[:, cost_col], errors="coerce").fillna(0) * rate,
        }).dropna(subset=["日付"])

        if tmp.empty:
            continue

        tmp["日付"] = pd.to_datetime(tmp["日付"]).dt.floor("D")
        g = tmp.groupby("日付", as_index=False)["値"].sum()
        g["区分"] = typ
        records.append(g)

    if not records:
        return None

    df_all = pd.concat(records, ignore_index=True)
    return (
        df_all
        .pivot_table(index="日付", columns="区分", values="値", aggfunc="sum")
        .fillna(0)
        .reset_index()
    )


# =====================
# 集計実行・プレビュー
# =====================
final_df = None
issue_cost_df = None

if cost_file:
    try:
        xls = pd.ExcelFile(cost_file)
        final_df = _build_application_summary_from_cost(xls)
        daily_cost_df, daily_cost_df_for_excel = _build_daily_cost_report_all_range(xls)
        issue_cost_df = _build_daily_issue_cost_from_cost_report(xls)
    except Exception as e:
        st.error(f"コストレポートの集計でエラーが発生しました: {e}")

if final_df is not None and not final_df.empty:
    st.subheader("📤 領域別コンディション集計用テーブル — 期間適用")
    st.dataframe(
        final_df[["分類", "媒体", "CV合計", "CV日割り", "合計費用"]],
        use_container_width=True,
    )

if daily_cost_df is not None and not daily_cost_df.empty:
    st.subheader("🗓️ コストレポート（日別・Forecast/実績）※Affの実績AFCV=*0.9、DisはnonIFRS除外")
    st.dataframe(daily_cost_df, use_container_width=True)


# =====================
# Excel出力（申込件数 / コストレポート日別）
# =====================
if (final_df is not None and not final_df.empty) or \
   (daily_cost_df_for_excel is not None and not daily_cost_df_for_excel.empty):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        # 1) 申込件数
        if final_df is not None and not final_df.empty:
            final_df.to_excel(writer, index=False, sheet_name="申込件数")
            ws = writer.sheets["申込件数"]
            ws.write(0, 6, "集計期間")
            ws.write(0, 7, f"{start_date} ～ {end_date}")
            ws.write(1, 6, "集計日数")
            ws.write(1, 7, days)
            ws.set_column(0, 0, 14)
            ws.set_column(1, 1, 26)
            ws.set_column(2, 4, 14)

        # 2) コストレポート日別（全期間）
        if daily_cost_df_for_excel is not None and not daily_cost_df_for_excel.empty:
            ws2 = workbook.add_worksheet("コストレポート日別")
            writer.sheets["コストレポート日別"] = ws2

            fmt_center = workbook.add_format({"align": "center", "valign": "vcenter", "border": 1})
            fmt_date = workbook.add_format({"num_format": "yyyy/m/d", "border": 1, "align": "center"})
            fmt_num = workbook.add_format({"num_format": "#,##0.00", "border": 1})

            # ヘッダー
            ws2.merge_range(0, 0, 2, 0, "日付", fmt_center)
            ws2.merge_range(0, 1, 0, 6, "Forecast", fmt_center)
            ws2.merge_range(0, 7, 0, 12, "実績", fmt_center)
            ws2.merge_range(0, 13, 0, 15, "発行", fmt_center)
            ws2.merge_range(1, 13, 1, 15, "実績", fmt_center)

            ws2.merge_range(1, 1, 1, 3, "AFCV", fmt_center)
            ws2.merge_range(1, 4, 1, 6, "配信費", fmt_center)
            ws2.merge_range(1, 7, 1, 9, "AFCV", fmt_center)
            ws2.merge_range(1, 10, 1, 12, "配信費", fmt_center)

            headers = ["Listing", "Display", "Affiliate"]
            for i, h in enumerate(headers):
                ws2.write(2, 1 + i, h, fmt_center)
                ws2.write(2, 4 + i, h, fmt_center)
                ws2.write(2, 7 + i, h, fmt_center)
                ws2.write(2, 10 + i, h, fmt_center)
                ws2.write(2, 13 + i, h, fmt_center)

            ws2.set_column(0, 0, 12)
            ws2.set_column(1, 15, 14)

            order_cols = [
                "Forecast_AFCV_Listing", "Forecast_AFCV_Display", "Forecast_AFCV_Affiliate",
                "Forecast_配信費_Listing", "Forecast_配信費_Display", "Forecast_配信費_Affiliate",
                "実績_AFCV_Listing", "実績_AFCV_Display", "実績_AFCV_Affiliate",
                "実績_配信費_Listing", "実績_配信費_Display", "実績_配信費_Affiliate",
            ]

            dfw = daily_cost_df_for_excel.copy()
            dfw["日付"] = pd.to_datetime(dfw["日付"], format="%Y/%m/%d")

            if issue_cost_df is not None:
                issue_cost_df["日付"] = pd.to_datetime(issue_cost_df["日付"]).dt.floor("D")

            start_row = 3
            for r, (_, row) in enumerate(dfw.iterrows(), start=start_row):
                ws2.write_datetime(r, 0, row["日付"], fmt_date)

                for c, col in enumerate(order_cols, start=1):
                    ws2.write_number(r, c, float(row.get(col, 0.0)), fmt_num)

                if issue_cost_df is not None:
                    hit = issue_cost_df[issue_cost_df["日付"] == row["日付"]]
                    if not hit.empty:
                        ws2.write_number(r, 13, float(hit["Listing"].iloc[0]) if "Listing" in hit.columns else 0.0, fmt_num)
                        ws2.write_number(r, 14, float(hit["Display"].iloc[0]) if "Display" in hit.columns else 0.0, fmt_num)
                        ws2.write_number(r, 15, float(hit["Affiliate"].iloc[0]) if "Affiliate" in hit.columns else 0.0, fmt_num)

    st.download_button(
        "📥 集計結果をダウンロード",
        data=output.getvalue(),
        file_name=f"集計結果_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("📌 コストレポートをアップロードすると集計結果が表示されます。")
