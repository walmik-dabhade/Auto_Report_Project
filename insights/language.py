import json
import os


# ==================================================
# FILE PATH
# ==================================================

KPI_FILE = "output/kpi_results.json"


# ==================================================
# LOAD KPI DATA
# ==================================================

def load_kpi_data():

    if not os.path.exists(KPI_FILE):

        raise FileNotFoundError(
            f"KPI file not found: {KPI_FILE}"
        )

    with open(
        KPI_FILE,
        "r",
        encoding="utf-8"
    ) as file:

        return json.load(file)


# ==================================================
# GENERATE LANGUAGE INSIGHTS
# ==================================================

def generate_language_insights(kpi_data):

    insights = []

    language = kpi_data.get(
        "language",
        {}
    )

    if not language:

        insights.append(
            "No language KPI data is available."
        )

        return insights


    # ==================================================
    # GET VALUES
    # ==================================================

    assessed = language.get(
        "assessed",
        0
    )

    missing = language.get(
        "missing",
        0
    )

    l1 = language.get(
        "l1",
        0
    )

    l2 = language.get(
        "l2",
        0
    )

    l3 = language.get(
        "l3",
        0
    )

    l4 = language.get(
        "l4",
        0
    )

    lower_level = language.get(
        "lower_level",
        0
    )

    higher_level = language.get(
        "higher_level",
        0
    )

    l1_rate = language.get(
        "l1_rate",
        0
    )

    l2_rate = language.get(
        "l2_rate",
        0
    )

    l3_rate = language.get(
        "l3_rate",
        0
    )

    l4_rate = language.get(
        "l4_rate",
        0
    )

    lower_rate = language.get(
        "lower_level_rate",
        0
    )

    higher_rate = language.get(
        "higher_level_rate",
        0
    )


    # ==================================================
    # ASSESSMENT COVERAGE
    # ==================================================

    insights.append(
        f"Language assessment data is available for "
        f"{assessed:,} children."
    )


    # ==================================================
    # MISSING ASSESSMENTS
    # ==================================================

    total_children = (
        assessed + missing
    )

    if total_children > 0:

        missing_rate = (
            missing
            / total_children
            * 100
        )

    else:

        missing_rate = 0


    if missing_rate >= 30:

        insights.append(
            f"{missing:,} children ({missing_rate:.2f}%) "
            f"do not have a recorded language assessment, "
            f"indicating a significant assessment-data gap."
        )

    elif missing_rate >= 15:

        insights.append(
            f"{missing:,} children ({missing_rate:.2f}%) "
            f"do not have a recorded language assessment, "
            f"which should be monitored for data completeness."
        )

    else:

        insights.append(
            f"{missing:,} children ({missing_rate:.2f}%) "
            f"do not have a recorded language assessment."
        )


    # ==================================================
    # DOMINANT LANGUAGE LEVEL
    # ==================================================

    levels = {

        "L1": l1_rate,

        "L2": l2_rate,

        "L3": l3_rate,

        "L4": l4_rate
    }


    dominant_level = max(
        levels,
        key=levels.get
    )

    dominant_rate = levels[
        dominant_level
    ]


    insights.append(
        f"{dominant_level} is the dominant language "
        f"level among assessed children, representing "
        f"{dominant_rate:.2f}%."
    )


    # ==================================================
    # LOWER VS HIGHER LEVEL
    # ==================================================

    insights.append(
        f"Lower language levels (L1 + L2) account for "
        f"{lower_rate:.2f}% of assessed children, while "
        f"higher levels (L3 + L4) account for "
        f"{higher_rate:.2f}%."
    )


    # ==================================================
    # HIGHER LEVEL PERFORMANCE
    # ==================================================

    if higher_rate >= 70:

        insights.append(
            f"Higher language levels show strong "
            f"performance, with {higher_rate:.2f}% of "
            f"assessed children at L3 or L4."
        )

    elif higher_rate >= 50:

        insights.append(
            f"{higher_rate:.2f}% of assessed children "
            f"are at higher language levels (L3 or L4), "
            f"indicating that more than half have reached "
            f"the higher proficiency group."
        )

    else:

        insights.append(
            f"Only {higher_rate:.2f}% of assessed children "
            f"are at higher language levels (L3 or L4), "
            f"indicating an area that may require "
            f"additional learning support."
        )


    # ==================================================
    # LOWER LEVEL PERFORMANCE
    # ==================================================

    if lower_rate >= 40:

        insights.append(
            f"{lower_rate:.2f}% of assessed children "
            f"remain at lower language levels (L1 or L2), "
            f"suggesting a need for targeted foundational "
            f"language support."
        )

    elif lower_rate >= 25:

        insights.append(
            f"{lower_rate:.2f}% of assessed children "
            f"are at lower language levels (L1 or L2), "
            f"which should continue to be monitored."
        )

    else:

        insights.append(
            f"Lower language levels account for only "
            f"{lower_rate:.2f}% of assessed children."
        )


    # ==================================================
    # L4 PERFORMANCE
    # ==================================================

    if l4_rate >= 50:

        insights.append(
            f"L4 is particularly strong, with "
            f"{l4_rate:.2f}% of assessed children "
            f"reaching the highest recorded language level."
        )


    return insights


# ==================================================
# MAIN
# ==================================================

if __name__ == "__main__":

    print("\n-----------------------------")
    print("LANGUAGE INSIGHTS")
    print("-----------------------------")

    kpi_data = load_kpi_data()

    insights = generate_language_insights(
        kpi_data
    )

    for number, insight in enumerate(
        insights,
        start=1
    ):

        print(
            f"{number}. {insight}"
        )