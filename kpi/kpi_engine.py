import pandas as pd
import json
import os


# ============================================================
# PATHS
# ============================================================

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

FILE_PATH = os.path.join(
    BASE_DIR,
    "data",
    "MONTH_DATA_FILE_2025_26_CLEANED.xlsx"
)

OUTPUT_DIR = os.path.join(
    BASE_DIR,
    "output"
)

OUTPUT_FILE = os.path.join(
    OUTPUT_DIR,
    "kpi_results.json"
)


# ============================================================
# LOAD DATA
# ============================================================

def load_data(file_path=FILE_PATH):

    df = pd.read_excel(
        file_path,
        sheet_name="CLEANED_CHILDREN_DATA"
    )

    print("Data loaded successfully!")

    return df


# ============================================================
# HELPER
# ============================================================

def find_column(df, possible_names):

    for column in df.columns:

        column_clean = str(column).strip().upper()

        for name in possible_names:

            if name.upper() in column_clean:
                return column

    return None


def safe_percentage(value, total):

    if total == 0:
        return 0

    return round((value / total) * 100, 2)


# ============================================================
# OVERALL KPIs
# ============================================================

def calculate_overall_kpis(df):

    total_records = len(df)

    software_column = find_column(
        df,
        ["SOFTWEAR ID", "SOFTWARE ID"]
    )

    if software_column:

        valid_ids = (
            df[software_column]
            .replace(
                ["", "0", 0, "NAN", "nan"],
                pd.NA
            )
            .dropna()
        )

        unique_valid_ids = valid_ids.nunique()

    else:

        unique_valid_ids = 0

    gender_column = find_column(
        df,
        ["GENDER"]
    )

    male = 0
    female = 0

    if gender_column:

        gender = (
            df[gender_column]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        male = int((gender == "MALE").sum())
        female = int((gender == "FEMALE").sum())

    return {

        "total_records": int(total_records),

        "unique_valid_software_ids": int(
            unique_valid_ids
        ),

        "male": int(male),

        "female": int(female)
    }


# ============================================================
# STATUS KPIs
# ============================================================

def calculate_status_kpis(df):

    status_columns = [

        column

        for column in df.columns

        if "STATUS" in str(column).upper()

    ]

    if not status_columns:

        return {}

    latest_status_column = status_columns[-1]

    status = (

        df[latest_status_column]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    continued = int(
        (status == "CONTINUED").sum()
    )

    dropout = int(
        (status == "DROPOUT").sum()
    )

    migrated = int(
        (status == "MIGRATE").sum()
    )

    graduate = int(
        (status == "GRADUATE").sum()
    )

    total_status = (
        continued +
        dropout +
        migrated +
        graduate
    )

    return {

        "latest_status_column":
            latest_status_column,

        "continued":
            continued,

        "dropout":
            dropout,

        "migrated":
            migrated,

        "graduate":
            graduate,

        "continuation_rate":
            safe_percentage(
                continued,
                total_status
            ),

        "dropout_rate":
            safe_percentage(
                dropout,
                total_status
            ),

        "migration_rate":
            safe_percentage(
                migrated,
                total_status
            ),

        "graduate_rate":
            safe_percentage(
                graduate,
                total_status
            )
    }


# ============================================================
# ACTIVITY KPIs
# ============================================================

def calculate_activity_kpis(df):

    activity_column = find_column(
        df,
        ["UPDATED CLASS TYPE", "CLASS TYPE"]
    )

    if not activity_column:

        return {}

    status_columns = [

        column

        for column in df.columns

        if "STATUS" in str(column).upper()

    ]

    if not status_columns:

        return {}

    latest_status_column = status_columns[-1]

    result = {}

    activities = (
        df[activity_column]
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
    )

    for activity in activities:

        activity_df = df[
            df[activity_column]
            .astype(str)
            .str.strip()
            == activity
        ]

        total = len(activity_df)

        status = (
            activity_df[latest_status_column]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        continued = int(
            (status == "CONTINUED").sum()
        )

        dropout = int(
            (status == "DROPOUT").sum()
        )

        migrated = int(
            (status == "MIGRATE").sum()
        )

        graduate = int(
            (status == "GRADUATE").sum()
        )

        result[activity] = {

            "total": total,

            "continued": continued,

            "dropout": dropout,

            "migrated": migrated,

            "graduate": graduate,

            "continuation_rate":
                safe_percentage(
                    continued,
                    total
                ),

            "dropout_rate":
                safe_percentage(
                    dropout,
                    total
                ),

            "migration_rate":
                safe_percentage(
                    migrated,
                    total
                ),

            "graduate_rate":
                safe_percentage(
                    graduate,
                    total
                )
        }

    return result


# ============================================================
# FUNDER KPIs
# ============================================================

def calculate_funder_kpis(df):

    funder_column = find_column(
        df,
        ["FUNDING WISE", "FUNDER"]
    )

    if not funder_column:

        return {}

    status_columns = [

        column

        for column in df.columns

        if "STATUS" in str(column).upper()

    ]

    if not status_columns:

        return {}

    latest_status_column = status_columns[-1]

    result = {}

    funders = (
        df[funder_column]
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
    )

    for funder in funders:

        funder_df = df[
            df[funder_column]
            .astype(str)
            .str.strip()
            == funder
        ]

        total = len(funder_df)

        status = (
            funder_df[latest_status_column]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        continued = int(
            (status == "CONTINUED").sum()
        )

        dropout = int(
            (status == "DROPOUT").sum()
        )

        migrated = int(
            (status == "MIGRATE").sum()
        )

        graduate = int(
            (status == "GRADUATE").sum()
        )

        result[funder] = {

            "total": total,

            "continued": continued,

            "dropout": dropout,

            "migrated": migrated,

            "graduate": graduate,

            "continuation_rate":
                safe_percentage(
                    continued,
                    total
                ),

            "dropout_rate":
                safe_percentage(
                    dropout,
                    total
                ),

            "migration_rate":
                safe_percentage(
                    migrated,
                    total
                ),

            "graduate_rate":
                safe_percentage(
                    graduate,
                    total
                )
        }

    return result


# ============================================================
# ATTENDANCE KPIs
# ============================================================

def calculate_attendance_kpis(df):

    attendance_column = find_column(
        df,
        ["ATTENDANCE %"]
    )

    if not attendance_column:

        return {}

    attendance = pd.to_numeric(
        df[attendance_column],
        errors="coerce"
    )

    attendance = attendance.dropna()

    total = len(attendance)

    high = int(
        ((attendance >= 80) &
         (attendance <= 100)).sum()
    )

    medium = int(
        ((attendance >= 51) &
         (attendance <= 79)).sum()
    )

    low = int(
        ((attendance >= 1) &
         (attendance <= 50)).sum()
    )

    return {

        "total": total,

        "high_attendance": high,

        "medium_attendance": medium,

        "low_attendance": low,

        "high_attendance_rate":
            safe_percentage(
                high,
                total
            ),

        "medium_attendance_rate":
            safe_percentage(
                medium,
                total
            ),

        "low_attendance_rate":
            safe_percentage(
                low,
                total
            )
    }


# ============================================================
# LANGUAGE KPIs
# ============================================================

def calculate_language_kpis(df):

    language_columns = [

        column

        for column in df.columns

        if "LANGUAGE LEVEL" in str(column).upper()

    ]

    if not language_columns:

        return {}

    latest_column = language_columns[-1]

    language = (

        df[latest_column]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    language = language.replace(
        "NAN",
        pd.NA
    )

    # Handle accidental lowercase values
    language = language.replace(
        {
            "L1": "L1",
            "L2": "L2",
            "L3": "L3",
            "L4": "L4"
        }
    )

    l1 = int(
        (language == "L1").sum()
    )

    l2 = int(
        (language == "L2").sum()
    )

    l3 = int(
        (language == "L3").sum()
    )

    l4 = int(
        (language == "L4").sum()
    )

    total_assessed = (
        l1 + l2 + l3 + l4
    )

    total_records = len(df)

    missing = (
        total_records -
        total_assessed
    )

    lower = l1 + l2
    higher = l3 + l4

    return {

        "latest_language_column":
            latest_column,

        "l1": l1,
        "l2": l2,
        "l3": l3,
        "l4": l4,

        "total_assessed":
            total_assessed,

        "missing":
            missing,

        "l1_l2":
            lower,

        "l3_l4":
            higher,

        "l1_rate":
            safe_percentage(
                l1,
                total_assessed
            ),

        "l2_rate":
            safe_percentage(
                l2,
                total_assessed
            ),

        "l3_rate":
            safe_percentage(
                l3,
                total_assessed
            ),

        "l4_rate":
            safe_percentage(
                l4,
                total_assessed
            ),

        "l1_l2_rate":
            safe_percentage(
                lower,
                total_assessed
            ),

        "l3_l4_rate":
            safe_percentage(
                higher,
                total_assessed
            )
    }


# ============================================================
# MATHEMATICS KPIs
# ============================================================

def calculate_mathematics_kpis(df):

    maths_columns = [

        column

        for column in df.columns

        if "MATHS LEVEL" in str(column).upper()

    ]

    if not maths_columns:

        return {}

    latest_column = maths_columns[-1]

    maths = (

        df[latest_column]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    maths = maths.replace(
        "NAN",
        pd.NA
    )

    result = {

        "latest_mathematics_column":
            latest_column
    }

    total_assessed = 0

    for level in [
        "M1",
        "M2",
        "M3",
        "M4",
        "M5",
        "M6",
        "M7"
    ]:

        count = int(
            (maths == level).sum()
        )

        result[
            level.lower()
        ] = count

        total_assessed += count

    total_records = len(df)

    missing = (
        total_records -
        total_assessed
    )

    low = (
        result["m1"] +
        result["m2"] +
        result["m3"]
    )

    middle = (
        result["m4"] +
        result["m5"]
    )

    higher = (
        result["m6"] +
        result["m7"]
    )

    result["total_assessed"] = (
        total_assessed
    )

    result["missing"] = missing

    result["m1_m2_m3"] = low

    result["m4_m5"] = middle

    result["m6_m7"] = higher

    result["m1_m2_m3_rate"] = safe_percentage(
        low,
        total_assessed
    )

    result["m4_m5_rate"] = safe_percentage(
        middle,
        total_assessed
    )

    result["m6_m7_rate"] = safe_percentage(
        higher,
        total_assessed
    )

    return result


# ============================================================
# LOCATION KPIs
# ============================================================

def calculate_location_kpis(df):

    location_column = find_column(
        df,
        ["LOCATION NAME"]
    )

    area_column = find_column(
        df,
        ["AREA"]
    )

    location_type_column = find_column(
        df,
        ["TYPE OF LOCATION"]
    )

    status_columns = [

        column

        for column in df.columns

        if "STATUS" in str(column).upper()

    ]

    result = {}

    # --------------------------------------------------------
    # LOCATION STATUS
    # --------------------------------------------------------

    if location_column and status_columns:

        latest_status_column = status_columns[-1]

        location_result = {}

        locations = (
            df[location_column]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
        )

        for location in locations:

            location_df = df[
                df[location_column]
                .astype(str)
                .str.strip()
                == location
            ]

            total = len(location_df)

            status = (
                location_df[
                    latest_status_column
                ]
                .astype(str)
                .str.strip()
                .str.upper()
            )

            continued = int(
                (status == "CONTINUED").sum()
            )

            dropout = int(
                (status == "DROPOUT").sum()
            )

            migrated = int(
                (status == "MIGRATE").sum()
            )

            graduate = int(
                (status == "GRADUATE").sum()
            )

            location_result[location] = {

                "total": total,

                "continued": continued,

                "dropout": dropout,

                "migrated": migrated,

                "graduate": graduate,

                "continuation_rate":
                    safe_percentage(
                        continued,
                        total
                    ),

                "dropout_rate":
                    safe_percentage(
                        dropout,
                        total
                    ),

                "migration_rate":
                    safe_percentage(
                        migrated,
                        total
                    ),

                "graduate_rate":
                    safe_percentage(
                        graduate,
                        total
                    )
            }

        result["locations"] = location_result

    # --------------------------------------------------------
    # AREA
    # --------------------------------------------------------

    if area_column:

        area_counts = (
            df[area_column]
            .fillna("Unknown")
            .astype(str)
            .str.strip()
            .value_counts()
        )

        result["areas"] = {
            str(key): int(value)
            for key, value
            in area_counts.items()
        }

    # --------------------------------------------------------
    # LOCATION TYPE
    # --------------------------------------------------------

    if location_type_column:

        type_counts = (
            df[location_type_column]
            .fillna("Unknown")
            .astype(str)
            .str.strip()
            .value_counts()
        )

        result["location_types"] = {
            str(key): int(value)
            for key, value
            in type_counts.items()
        }

    return result


# ============================================================
# MASTER KPI FUNCTION
# ============================================================

def calculate_kpis(df):

    return {

        "overall":
            calculate_overall_kpis(df),

        "status":
            calculate_status_kpis(df),

        "activity":
            calculate_activity_kpis(df),

        "funder":
            calculate_funder_kpis(df),

        "attendance":
            calculate_attendance_kpis(df),

        "language":
            calculate_language_kpis(df),

        "mathematics":
            calculate_mathematics_kpis(df),

        "location":
            calculate_location_kpis(df)
    }


# ============================================================
# SAVE KPIs
# ============================================================

def save_kpis(kpis, output_file=OUTPUT_FILE):

    os.makedirs(
        os.path.dirname(output_file),
        exist_ok=True
    )

    with open(
        output_file,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            kpis,
            file,
            indent=4,
            ensure_ascii=False
        )

    print()
    print("KPI results saved to:")
    print(output_file)


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":

    print()
    print("-----------------------------")
    print("KPI ENGINE")
    print("-----------------------------")

    df = load_data()

    print()
    print("Rows:", len(df))
    print("Columns:", len(df.columns))

    kpis = calculate_kpis(df)

    save_kpis(kpis)

    print()
    print("-----------------------------")
    print("KPI CALCULATION COMPLETED")
    print("-----------------------------")