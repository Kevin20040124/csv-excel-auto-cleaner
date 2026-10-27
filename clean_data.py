"""
Excel / CSV 自動整理工具
用法：
    python clean_data.py              -> 自動產生範例髒資料並整理（用來測試）
    python clean_data.py 你的檔案.csv   -> 整理你自己的檔案（支援 .csv / .xlsx）
輸出：原檔名_cleaned.xlsx（含「清理後資料」與「整理報告」兩個分頁）
"""
import sys
from pathlib import Path

import numpy as np
import pandas as pd


def make_demo(path: Path) -> None:
    """產生一份故意弄髒的範例資料。"""
    df = pd.DataFrame({
        " 姓名 ": ["王小明", "王小明", "李小華 ", None, "陳大文", "李小華 "],
        "Email": ["a@x.com", "a@x.com", "B@X.COM ", None, "c@x.com", "B@X.COM "],
        "金額": ["1,200", "1,200", "850", None, "N/A", "850"],
        "日期": ["2026/10/01", "2026/10/01", "2026-10-03", None, "2026.10.05", "2026-10-03"],
    })
    df.to_csv(path, index=False, encoding="utf-8-sig")


def load(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(path)
    return pd.read_csv(path, encoding="utf-8-sig")


def clean(df: pd.DataFrame):
    report = {"原始列數": len(df)}
    df = df.copy()

    # 1. 欄位名稱去除前後空白
    df.columns = [str(c).strip() for c in df.columns]

    # 2. 儲存格去除前後空白，並把常見的「空值寫法」統一成空值
    for col in df.columns:
        df[col] = df[col].apply(lambda v: v.strip() if isinstance(v, str) else v)
    df = df.replace({"": np.nan, "N/A": np.nan, "n/a": np.nan, "NA": np.nan, "-": np.nan})

    # 3. Email 統一小寫
    for col in df.columns:
        if "email" in col.lower() or "信箱" in col:
            df[col] = df[col].str.lower()

    # 4. 日期欄位統一成 YYYY-MM-DD
    for col in df.columns:
        if "日期" in col or "date" in col.lower():
            parsed = df[col].apply(
                lambda v: pd.to_datetime(str(v).replace(".", "-").replace("/", "-"), errors="coerce")
                if pd.notna(v) else pd.NaT
            )
            df[col] = pd.to_datetime(parsed).dt.strftime("%Y-%m-%d")

    # 5. 看起來像數字的欄位（含千分位逗號）轉成真正的數字
    for col in df.columns:
        if not pd.api.types.is_numeric_dtype(df[col]):
            as_num = pd.to_numeric(df[col].astype(str).str.replace(",", "", regex=False), errors="coerce")
            non_empty = df[col].notna().sum()
            if non_empty and as_num.notna().sum() / non_empty >= 0.6:
                df[col] = as_num

    # 6. 刪除整列都是空的
    before = len(df)
    df = df.dropna(how="all")
    report["刪除空白列"] = before - len(df)

    # 7. 刪除重複列
    before = len(df)
    df = df.drop_duplicates()
    report["刪除重複列"] = before - len(df)

    report["整理後列數"] = len(df)
    for col in df.columns:
        report[f"缺值數量_{col}"] = int(df[col].isna().sum())

    return df.reset_index(drop=True), report


def main() -> None:
    if len(sys.argv) > 1:
        src = Path(sys.argv[1])
    else:
        src = Path("demo_messy.csv")
        make_demo(src)
        print(f"沒有指定檔案，已產生範例資料：{src}")

    cleaned, report = clean(load(src))
    out = src.with_name(src.stem + "_cleaned.xlsx")

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        cleaned.to_excel(writer, sheet_name="清理後資料", index=False)
        pd.DataFrame(list(report.items()), columns=["項目", "數值"]).to_excel(
            writer, sheet_name="整理報告", index=False
        )

    print(f"完成！已輸出：{out}")
    for k, v in report.items():
        print(f"  {k}: {v}")


if __name__ == "__main__":
    main()
