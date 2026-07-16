from __future__ import annotations

import gspread
import pandas as pd
from loguru import logger
from oauth2client.service_account import ServiceAccountCredentials


# ============================================================
# 基本設定
# ============================================================

SPREADSHEET_ID = "1a7EbUzvvnTgQzMetEDrWxMJ_xFlFGsCtoGS8E48TTuk"

CREDENTIALS_PATH = "credentials.json"

SHEET_LIST_NAME = "SheetList"
RESULT_SHEET_NAME = "小天使配對總表"
ROSTER_SHEET_NAME = "小天使總表"

# SheetList 中工作表名稱位於第幾欄
# 1 代表 A欄
SHEET_LIST_NAME_COLUMN = 1

# 小天使總表中，找不到的小天使輸出位置
MISSING_ANGEL_NAME_COLUMN = "W"
MISSING_ANGEL_COUNT_COLUMN = "X"


# ============================================================
# 共用函式
# ============================================================

def normalize_name(value: object) -> str:
    """統一姓名格式，避免前後空白影響比對。"""

    return str(value or "").strip().lower()


def connect_to_sheet() -> gspread.Client:
    """連線到 Google Sheets。"""

    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]

    credentials = (
        ServiceAccountCredentials.from_json_keyfile_name(
            CREDENTIALS_PATH,
            scope,
        )
    )

    client = gspread.authorize(credentials)

    logger.info("✅ Google Sheets 認證成功")

    return client


# ============================================================
# 讀取 SheetList
# ============================================================

def get_target_sheet_names(
    spreadsheet: gspread.Spreadsheet,
    list_sheet_name: str = SHEET_LIST_NAME,
    name_column: int = SHEET_LIST_NAME_COLUMN,
) -> list[str]:
    """
    從 SheetList 指定欄位讀取要處理的工作表名稱。

    會自動：
    1. 排除空白
    2. 排除表頭
    3. 排除結果工作表
    4. 排除不存在的工作表
    5. 移除重複名稱
    """

    try:
        list_ws = spreadsheet.worksheet(list_sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        logger.error(
            f"❌ 找不到工作表「{list_sheet_name}」"
        )
        return []

    values = list_ws.col_values(name_column)

    existing_sheet_names = {
        worksheet.title
        for worksheet in spreadsheet.worksheets()
    }

    excluded_names = {
        list_sheet_name,
        RESULT_SHEET_NAME,
        ROSTER_SHEET_NAME,
    }

    possible_headers = {
        "sheet name",
        "sheetname",
        "工作表名稱",
        "分頁名稱",
        "表單名稱",
    }

    target_names: list[str] = []
    missing_names: list[str] = []

    for value in values:
        sheet_name = str(value).strip()

        if not sheet_name:
            continue

        if sheet_name.lower() in possible_headers:
            continue

        if sheet_name in excluded_names:
            continue

        if sheet_name not in existing_sheet_names:
            missing_names.append(sheet_name)
            continue

        target_names.append(sheet_name)

    # 移除重複，但保留 SheetList 原本順序
    target_names = list(dict.fromkeys(target_names))
    missing_names = list(dict.fromkeys(missing_names))

    logger.info(
        f"✅ SheetList 找到 {len(target_names)} 張有效工作表："
        f"{target_names}"
    )

    if missing_names:
        logger.warning(
            f"⚠️ SheetList 中有 {len(missing_names)} 張工作表不存在："
            f"{missing_names}"
        )

    return target_names


# ============================================================
# 讀取各班的小天使配對資料
# ============================================================

def count_students_needing_angels(
    spreadsheet: gspread.Spreadsheet,
    sheet_names: list[str],
) -> pd.DataFrame:
    """
    讀取 SheetList 指定的工作表，整理需要小天使的學員。

    預期欄位：
    - ID+名字
    - 小天使關懷員
    """

    total_count = 0
    all_valid_rows: list[pd.DataFrame] = []

    required_columns = {
        "ID+名字",
        "小天使關懷員",
    }

    for sheet_name in sheet_names:
        logger.info(f"🔄 開始處理工作表：{sheet_name}")

        try:
            worksheet = spreadsheet.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            logger.warning(
                f"⚠️ 找不到工作表「{sheet_name}」，已略過"
            )
            continue

        records = worksheet.get_all_values()

        if not records:
            logger.warning(
                f"⚠️ 工作表「{sheet_name}」沒有資料，已略過"
            )
            continue

        # 目前只讀取 A～G 欄
        values = []

        for row in records:
            row_values = row[:7]

            # 不足七欄時補空白，避免 DataFrame 欄位數不同
            if len(row_values) < 7:
                row_values += [""] * (7 - len(row_values))

            values.append(row_values)

        if not values:
            continue

        headers = [
            str(value).strip()
            for value in values[0]
        ]

        data_rows = values[1:]

        if not data_rows:
            logger.warning(
                f"⚠️ 工作表「{sheet_name}」只有表頭，已略過"
            )
            continue

        df = pd.DataFrame(
            data_rows,
            columns=headers,
        )

        missing_columns = required_columns - set(df.columns)

        if missing_columns:
            logger.warning(
                f"⚠️ 工作表「{sheet_name}」缺少必要欄位："
                f"{sorted(missing_columns)}，已略過"
            )
            continue

        # 大笑體系格式：
        # 例如 1234_王小明 → 王小明
        df["學員姓名"] = (
            df["ID+名字"]
            .fillna("")
            .astype(str)
            .apply(extract_student_name)
        )

        df["小天使關懷員"] = (
            df["小天使關懷員"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        # 保留來源工作表，方便日後檢查
        df["來源工作表"] = sheet_name

        df_valid = df[
            (df["學員姓名"] != "") &
            (df["學員姓名"] != "學員姓名") &
            (df["小天使關懷員"] != "") &
            (df["小天使關懷員"] != "不需要") &
            (df["小天使關懷員"] != "小天使關懷員")
        ].copy()

        all_valid_rows.append(df_valid)

        count = len(df_valid)
        total_count += count

        logger.info(
            f"✅ {sheet_name}：需要小天使的學員共 {count} 位"
        )

    logger.info(
        f"🎯 所有工作表總共需要小天使的學員："
        f"{total_count} 位"
    )

    if not all_valid_rows:
        return pd.DataFrame()

    result = pd.concat(
        all_valid_rows,
        ignore_index=True,
    )

    return result


def extract_student_name(value: object) -> str:
    """
    從 ID+名字欄位取出姓名。

    例如：
    1234_王小明 → 王小明
    王小明 → 王小明
    """

    text = str(value or "").strip()

    if "_" in text:
        return text.split("_", 1)[1].strip()

    return text


# ============================================================
# 統整小天使配對
# ============================================================

def summarize_angel_assignments(
    df_filtered: pd.DataFrame,
) -> tuple[dict[str, list[str]], pd.DataFrame]:
    """依照小天使統整負責的學員名單。"""

    if df_filtered.empty:
        logger.warning("⚠️ 沒有資料可以統整")

        empty_summary = pd.DataFrame(
            columns=[
                "小天使",
                "關懷數量",
                "關懷學員列表",
            ]
        )

        return {}, empty_summary

    grouped: dict[str, list[str]] = (
        df_filtered
        .groupby("小天使關懷員")["學員姓名"]
        .apply(list)
        .to_dict()
    )

    summary_rows = []

    for angel, students in grouped.items():
        # 同一位學員如果重複出現，名單只保留一次
        unique_students = list(dict.fromkeys(students))

        # 同步修正 grouped，讓接人數也是去重複後的數量
        grouped[angel] = unique_students

        summary_rows.append(
            {
                "小天使": angel,
                "關懷數量": len(unique_students),
                "關懷學員列表": ", ".join(unique_students),
            }
        )

    summary_df = pd.DataFrame(summary_rows)

    if not summary_df.empty:
        summary_df = summary_df.sort_values(
            by=[
                "關懷數量",
                "小天使",
            ],
            ascending=[
                False,
                True,
            ],
        ).reset_index(drop=True)

    logger.info(
        f"✅ 共統整出 {len(summary_df)} 位小天使"
    )

    logger.info(
        "\n📋 小天使配對統整表：\n"
        f"{summary_df.to_string(index=False)}"
    )

    return grouped, summary_df


# ============================================================
# 更新小天使配對總表
# ============================================================

def update_assignment_summary_sheet(
    spreadsheet: gspread.Spreadsheet,
    summary_df: pd.DataFrame,
    result_sheet_name: str = RESULT_SHEET_NAME,
) -> None:
    """更新小天使配對總表。"""

    try:
        worksheet = spreadsheet.worksheet(result_sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=result_sheet_name,
            rows=100,
            cols=20,
        )

        logger.info(
            f"✅ 已建立工作表「{result_sheet_name}」"
        )

    headers = [
        "小天使",
        "關懷數量",
        "關懷學員列表",
    ]

    if summary_df.empty:
        data_to_write = [headers]
    else:
        data_to_write = [
            summary_df.columns.tolist(),
            *summary_df.fillna("").values.tolist(),
        ]

    required_rows = max(100, len(data_to_write) + 10)
    required_cols = max(20, len(headers))

    if (
        worksheet.row_count < required_rows or
        worksheet.col_count < required_cols
    ):
        worksheet.resize(
            rows=max(worksheet.row_count, required_rows),
            cols=max(worksheet.col_count, required_cols),
        )

    # 清除舊資料
    worksheet.clear()

    worksheet.update(
        range_name="A1",
        values=data_to_write,
        value_input_option="RAW",
    )

    logger.info(
        f"✅ 「{result_sheet_name}」已更新，"
        f"共 {len(summary_df)} 位小天使"
    )


# ============================================================
# 更新小天使總表
# ============================================================

def fill_angel_counts(
    worksheet: gspread.Worksheet,
    existing_dict: dict[str, list[str]],
) -> None:
    """
    更新小天使總表中的接人數。

    名冊格式假設：
    - B、E、H、K……是小天使姓名
    - C、F、I、L……是接人數
    - 每組間隔三欄

    找不到的小天使會寫入 W:X。
    """

    values = worksheet.get_all_values()

    if not values:
        logger.warning("⚠️ 小天使總表是空的，無法更新")
        return

    data_rows = values[1:]

    normalized_dict: dict[str, list[str]] = {}
    display_name_dict: dict[str, str] = {}

    for angel, students in existing_dict.items():
        normalized_name = normalize_name(angel)

        if not normalized_name:
            continue

        normalized_dict[normalized_name] = students
        display_name_dict[normalized_name] = str(angel).strip()

    angels_in_roster: set[str] = set()
    count_cells: list[gspread.Cell] = []

    # sheet_row 從 2 開始，因為第 1 列是表頭
    for sheet_row, row in enumerate(data_rows, start=2):

        # 0-based：
        # B欄索引1、E欄索引4、H欄索引7……
        for col_idx in range(1, len(row), 3):
            angel_name = normalize_name(row[col_idx])

            if not angel_name:
                continue

            angels_in_roster.add(angel_name)

            assigned_count = len(
                normalized_dict.get(angel_name, [])
            )

            # col_idx 是 0-based。
            # B欄索引1，接人數在 C欄，Google Sheets欄號為3。
            count_column = col_idx + 2

            count_cells.append(
                gspread.Cell(
                    row=sheet_row,
                    col=count_column,
                    value=assigned_count,
                )
            )

    # 只更新接人數，不覆蓋姓名或其他公式
    if count_cells:
        worksheet.update_cells(
            count_cells,
            value_input_option="RAW",
        )

    logger.info(
        f"✅ 小天使總表已更新 "
        f"{len(count_cells)} 個接人數欄位"
    )

    missing_angels = sorted(
        normalized_name
        for normalized_name in normalized_dict
        if normalized_name not in angels_in_roster
    )

    logger.info(
        f"🔍 配對表中有、名冊中沒有的小天使共 "
        f"{len(missing_angels)} 位："
        f"{[display_name_dict[name] for name in missing_angels]}"
    )

    # 清除上一次留在 W:X 的結果，避免舊資料殘留
    worksheet.batch_clear(
        [
            f"{MISSING_ANGEL_NAME_COLUMN}2:"
            f"{MISSING_ANGEL_COUNT_COLUMN}"
        ]
    )

    if not missing_angels:
        return

    missing_rows = []

    for normalized_name in missing_angels:
        display_name = display_name_dict[normalized_name]

        assigned_count = len(
            normalized_dict.get(normalized_name, [])
        )

        missing_rows.append(
            [
                display_name,
                assigned_count,
            ]
        )

    last_missing_row = len(missing_rows) + 1

    worksheet.update(
        range_name=(
            f"{MISSING_ANGEL_NAME_COLUMN}2:"
            f"{MISSING_ANGEL_COUNT_COLUMN}{last_missing_row}"
        ),
        values=missing_rows,
        value_input_option="RAW",
    )

    summary_row = last_missing_row + 1

    worksheet.update(
        range_name=f"{MISSING_ANGEL_COUNT_COLUMN}{summary_row}",
        values=[
            [
                f"未配對到名單的小天使共 "
                f"{len(missing_angels)} 位"
            ]
        ],
        value_input_option="RAW",
    )

    logger.info(
        f"✅ 未配對到名冊的小天使已寫入 "
        f"{MISSING_ANGEL_NAME_COLUMN}:"
        f"{MISSING_ANGEL_COUNT_COLUMN}"
    )


# ============================================================
# 主程式
# ============================================================

def main() -> None:
    """執行完整的小天使配對統計流程。"""

    logger.info("🚀 開始執行小天使配對統計")

    # 1. 認證與開啟試算表
    client = connect_to_sheet()
    spreadsheet = client.open_by_key(SPREADSHEET_ID)

    logger.info(
        f"✅ 已開啟 Google 試算表：{spreadsheet.title}"
    )

    # 2. 從 SheetList 取得需要處理的分頁
    sheet_names = get_target_sheet_names(
        spreadsheet=spreadsheet,
        list_sheet_name=SHEET_LIST_NAME,
        name_column=SHEET_LIST_NAME_COLUMN,
    )

    if not sheet_names:
        logger.warning(
            "⚠️ SheetList 沒有可處理的工作表，程式結束"
        )
        return

    # 3. 讀取各班配對資料
    df_filtered = count_students_needing_angels(
        spreadsheet=spreadsheet,
        sheet_names=sheet_names,
    )

    if df_filtered.empty:
        logger.warning(
            "⚠️ 沒有需要小天使的學員"
        )

        # 即使沒有資料，也把配對總表更新成只有表頭
        empty_summary = pd.DataFrame(
            columns=[
                "小天使",
                "關懷數量",
                "關懷學員列表",
            ]
        )

        update_assignment_summary_sheet(
            spreadsheet=spreadsheet,
            summary_df=empty_summary,
        )

        # 將小天使總表的接人數更新成 0
        roster_ws = spreadsheet.worksheet(
            ROSTER_SHEET_NAME
        )

        fill_angel_counts(
            worksheet=roster_ws,
            existing_dict={},
        )

        logger.info("✅ 程式執行完成")
        return

    # 4. 統整小天使配對
    grouped_dict, summary_df = summarize_angel_assignments(
        df_filtered
    )

    # 5. 更新小天使配對總表
    update_assignment_summary_sheet(
        spreadsheet=spreadsheet,
        summary_df=summary_df,
        result_sheet_name=RESULT_SHEET_NAME,
    )

    # 6. 更新小天使總表接人數
    try:
        roster_ws = spreadsheet.worksheet(
            ROSTER_SHEET_NAME
        )
    except gspread.exceptions.WorksheetNotFound:
        logger.error(
            f"❌ 找不到工作表「{ROSTER_SHEET_NAME}」，"
            "無法更新接人數"
        )
        return

    fill_angel_counts(
        worksheet=roster_ws,
        existing_dict=grouped_dict,
    )

    logger.info("🎉 小天使配對統計全部完成")


if __name__ == "__main__":
    main()