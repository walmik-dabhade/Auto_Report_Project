import pandas as pd
import json
import os
import sys

# Allow importing kpi_engine.py from the kpi folder
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

sys.path.append(
    os.path.join(BASE_DIR, "kpi")
)

from kpi_engine import calculate_kpis


# ============================================================
# PATHS
# ============================================================

DATA_FILE = os.path.join(
    BASE_DIR,
    "data",
    "MONTH_DATA_FILE_2025_26_CLEANED.xlsx"
)

OUTPUT_DIR = os.path.join(
    BASE_DIR,
    "output",
    "funder_kpis"
)


# ============================================================
# LOAD DATA
# ============================================================

def load_data():

    print()
    print("-----------------------------")
    print("LOADING CLEANED DATA")
    print("-----------------------------")

    df = pd.read_excel(
        DATA_FILE,
        sheet_name="CLEANED_CHILDREN_DATA"
    )

    print(f"Total records loaded: {len(df)}")
    print(f"Total columns: {len(df.columns)}")

    return df


# ============================================================
# FIND FUNDER COLUMN
# ============================================================

def find_funder_column(df):

    for column in df.columns:

        column_name = str(column).strip().upper()

        if column_name == "FUNDING WISE":

            return column

    # Fallback in case spacing/capitalisation is different
    for column in df.columns:

        column_name = str(column).strip().upper()

        if "FUNDING" in column_name:

            return column

    return None


# ============================================================
# CLEAN FUNDER NAME
# ============================================================

def clean_funder_name(value):

    return (
        str(value)
        .strip()
        .upper()
    )


# ============================================================
# MAIN FUNDER PROCESS
# ============================================================

def generate_funder_kpis():

    df = load_data()

    funder_column = find_funder_column(df)

    if funder_column is None:

        print()
        print("ERROR: FUNDING WISE column was not found.")
        print()
        print("Available columns containing funding:")
        
        for column in df.columns:

            if "FUND" in str(column).upper():

                print(column)

        return

    print()
    print("-----------------------------")
    print("FUNDER COLUMN")
    print("-----------------------------")

    print(f"Using column: {funder_column}")


    # --------------------------------------------------------
    # CLEAN FUNDER VALUES
    # --------------------------------------------------------

    df["_FUNDER_CLEAN"] = (
        df[funder_column]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # Remove invalid values
    invalid_values = [
        "",
        "NAN",
        "NONE",
        "NULL"
    ]

    df = df[
        ~df["_FUNDER_CLEAN"].isin(
            invalid_values
        )
    ]


    # --------------------------------------------------------
    # GET UNIQUE FUNDERS
    # --------------------------------------------------------

    funders = sorted(
        df["_FUNDER_CLEAN"]
        .unique()
    )

    print()
    print("-----------------------------")
    print("FUNDERS FOUND")
    print("-----------------------------")

    for funder in funders:

        count = int(
            (
                df["_FUNDER_CLEAN"]
                == funder
            ).sum()
        )

        print(
            f"{funder}: {count} records"
        )


    # --------------------------------------------------------
    # CREATE OUTPUT DIRECTORY
    # --------------------------------------------------------

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True
    )


    # --------------------------------------------------------
    # GENERATE KPI JSON FOR EACH FUNDER
    # --------------------------------------------------------

    print()
    print("-----------------------------")
    print("GENERATING FUNDER KPIs")
    print("-----------------------------")


    summary = {}


    for funder in funders:

        print()
        print(f"Processing: {funder}")


        # Filter dataset
        funder_df = df[
            df["_FUNDER_CLEAN"]
            == funder
        ].copy()


        # Remove helper column
        funder_df = funder_df.drop(
            columns=["_FUNDER_CLEAN"]
        )


        # Calculate KPIs
        kpis = calculate_kpis(
            funder_df
        )


        # Add funder information
        kpis["funder_name"] = funder

        kpis["source_records"] = len(
            funder_df
        )


        # ----------------------------------------------------
        # SAFE FILE NAME
        # ----------------------------------------------------

        safe_name = (
            funder
            .replace("/", "_")
            .replace("\\", "_")
            .replace(" ", "_")
        )


        output_file = os.path.join(
            OUTPUT_DIR,
            f"{safe_name}_kpis.json"
        )


        # ----------------------------------------------------
        # SAVE JSON
        # ----------------------------------------------------

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


        # ----------------------------------------------------
        # SUMMARY
        # ----------------------------------------------------

        overall = kpis.get(
            "overall",
            {}
        )

        status = kpis.get(
            "status",
            {}
        )


        summary[funder] = {

            "records":
                overall.get(
                    "total_records",
                    len(funder_df)
                ),

            "unique_ids":
                overall.get(
                    "unique_valid_software_ids",
                    0
                ),

            "male":
                overall.get(
                    "male",
                    0
                ),

            "female":
                overall.get(
                    "female",
                    0
                ),

            "continued":
                status.get(
                    "continued",
                    0
                ),

            "dropout":
                status.get(
                    "dropout",
                    0
                ),

            "migrated":
                status.get(
                    "migrated",
                    0
                ),

            "graduate":
                status.get(
                    "graduate",
                    0
                ),

            "continuation_rate":
                status.get(
                    "continuation_rate",
                    0
                ),

            "dropout_rate":
                status.get(
                    "dropout_rate",
                    0
                ),

            "migration_rate":
                status.get(
                    "migration_rate",
                    0
                ),

            "graduate_rate":
                status.get(
                    "graduate_rate",
                    0
                ),

            "kpi_file":
                output_file
        }


        print(
            f"  Records: {len(funder_df)}"
        )

        print(
            f"  KPI file: {output_file}"
        )


    # --------------------------------------------------------
    # SAVE MASTER SUMMARY
    # --------------------------------------------------------

    summary_file = os.path.join(
        OUTPUT_DIR,
        "funder_summary.json"
    )


    with open(
        summary_file,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            summary,
            file,
            indent=4,
            ensure_ascii=False
        )


    # --------------------------------------------------------
    # COMPLETED
    # --------------------------------------------------------

    print()
    print("========================================")
    print("FUNDER KPI GENERATION COMPLETED")
    print("========================================")

    print()
    print(
        f"Total funders processed: {len(funders)}"
    )

    print()
    print("Output folder:")
    print(OUTPUT_DIR)

    print()
    print("Master summary:")
    print(summary_file)


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    generate_funder_kpis()