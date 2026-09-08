import json
import os
import sys


# ============================================================
# PROJECT PATH
# ============================================================

PROJECT_ROOT = os.path.dirname(
    os.path.dirname(
        os.path.abspath(__file__)
    )
)

OUTPUT_FOLDER = os.path.join(
    PROJECT_ROOT,
    "output"
)

KPI_FILE = os.path.join(
    OUTPUT_FOLDER,
    "kpi_results.json"
)

FINAL_INSIGHTS_FILE = os.path.join(
    OUTPUT_FOLDER,
    "final_insights.json"
)


# ============================================================
# IMPORT INSIGHT MODULES
# ============================================================

sys.path.insert(
    0,
    os.path.dirname(
        os.path.abspath(__file__)
    )
)

from overall import generate_overall_insights
from activity import generate_activity_insights
from funder import generate_funder_insights
from attendance import generate_attendance_insights
from language import generate_language_insights
from mathematics import generate_mathematics_insights
from location import generate_location_insights


# ============================================================
# LOAD KPI DATA
# ============================================================

def load_kpi_data():

    if not os.path.exists(KPI_FILE):

        raise FileNotFoundError(
            f"KPI file not found:\n{KPI_FILE}"
        )

    with open(
        KPI_FILE,
        "r",
        encoding="utf-8"
    ) as file:

        return json.load(file)


# ============================================================
# GENERATE ALL INSIGHTS
# ============================================================

def generate_all_insights(kpi_data):

    all_insights = {

        "overall": generate_overall_insights(
            kpi_data
        ),

        "activity": generate_activity_insights(
            kpi_data
        ),

        "funder": generate_funder_insights(
            kpi_data
        ),

        "attendance": generate_attendance_insights(
            kpi_data
        ),

        "language": generate_language_insights(
            kpi_data
        ),

        "mathematics": generate_mathematics_insights(
            kpi_data
        ),

        "location": generate_location_insights(
            kpi_data
        )
    }

    return all_insights


# ============================================================
# SAVE FINAL INSIGHTS
# ============================================================

def save_insights(insights):

    os.makedirs(
        OUTPUT_FOLDER,
        exist_ok=True
    )

    with open(
        FINAL_INSIGHTS_FILE,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            insights,
            file,
            indent=4,
            ensure_ascii=False
        )


# ============================================================
# PRINT INSIGHTS
# ============================================================

def print_insights(insights):

    print("\n")
    print("=" * 60)
    print("AUTOMATED REPORT INSIGHTS")
    print("=" * 60)


    section_titles = {

        "overall":
            "OVERALL INSIGHTS",

        "activity":
            "ACTIVITY INSIGHTS",

        "funder":
            "FUNDER INSIGHTS",

        "attendance":
            "ATTENDANCE INSIGHTS",

        "language":
            "LANGUAGE INSIGHTS",

        "mathematics":
            "MATHEMATICS INSIGHTS",

        "location":
            "LOCATION INSIGHTS"
    }


    for section, section_insights in insights.items():

        print("\n")
        print("-" * 60)
        print(
            section_titles.get(
                section,
                section.upper()
            )
        )
        print("-" * 60)


        for number, insight in enumerate(
            section_insights,
            start=1
        ):

            print(
                f"{number}. {insight}"
            )


    print("\n")
    print("=" * 60)
    print("INSIGHTS GENERATION COMPLETED")
    print("=" * 60)

    print(
        f"\nSaved to:\n{FINAL_INSIGHTS_FILE}"
    )


# ============================================================
# MAIN
# ============================================================

def main():

    print(
        "\nLoading KPI results..."
    )

    kpi_data = load_kpi_data()

    print(
        "KPI data loaded successfully!"
    )

    print(
        "\nGenerating insights..."
    )

    insights = generate_all_insights(
        kpi_data
    )

    save_insights(
        insights
    )

    print_insights(
        insights
    )


# ============================================================
# RUN PROGRAM
# ============================================================

if __name__ == "__main__":

    main()