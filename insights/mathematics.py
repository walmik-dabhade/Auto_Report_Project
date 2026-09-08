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
# GENERATE MATHEMATICS INSIGHTS
# ==================================================

def generate_mathematics_insights(kpi_data):

    insights = []

    mathematics = kpi_data.get(
        "mathematics",
        {}
    )

    if not mathematics:

        insights.append(
            "No mathematics KPI data is available."
        )

        return insights


    # ==================================================
    # GET VALUES
    # ==================================================

    assessed = mathematics.get(
        "assessed",
        0
    )

    missing = mathematics.get(
        "missing",
        0
    )

    lower_level = mathematics.get(
        "lower_level",
        0
    )

    middle_level = mathematics.get(
        "middle_level",
        0
    )

    higher_level = mathematics.get(
        "higher_level",
        0
    )

    lower_rate = mathematics.get(
        "lower_level_rate",
        0
    )

    middle_rate = mathematics.get(
        "middle_level_rate",
        0
    )

    higher_rate = mathematics.get(
        "higher_level_rate",
        0
    )

    levels = mathematics.get(
        "levels",
        {}
    )

    level_rates = mathematics.get(
        "level_rates",
        {}
    )


    # ==================================================
    # ASSESSMENT COVERAGE
    # ==================================================

    insights.append(
        f"Mathematics assessment data is available "
        f"for {assessed:,} children."
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
            f"do not have a recorded mathematics "
            f"assessment, indicating a significant "
            f"assessment-data gap."
        )

    elif missing_rate >= 15:

        insights.append(
            f"{missing:,} children ({missing_rate:.2f}%) "
            f"do not have a recorded mathematics "
            f"assessment, which should be monitored "
            f"for data completeness."
        )

    else:

        insights.append(
            f"{missing:,} children ({missing_rate:.2f}%) "
            f"do not have a recorded mathematics "
            f"assessment."
        )


    # ==================================================
    # DOMINANT LEVEL
    # ==================================================

    if level_rates:

        dominant_level = max(
            level_rates,
            key=level_rates.get
        )

        dominant_rate = level_rates[
            dominant_level
        ]

        insights.append(
            f"{dominant_level} is the dominant "
            f"mathematics level among assessed "
            f"children, representing "
            f"{dominant_rate:.2f}%."
        )


    # ==================================================
    # LEVEL GROUP DISTRIBUTION
    # ==================================================

    insights.append(
        f"Lower mathematics levels (M1 + M2 + M3) "
        f"account for {lower_rate:.2f}% of assessed "
        f"children, middle levels (M4 + M5) account "
        f"for {middle_rate:.2f}%, while higher levels "
        f"(M6 + M7) account for {higher_rate:.2f}%."
    )


    # ==================================================
    # HIGHER LEVEL PERFORMANCE
    # ==================================================

    if higher_rate >= 50:

        insights.append(
            f"Higher mathematics levels show strong "
            f"performance, with {higher_rate:.2f}% of "
            f"assessed children at M6 or M7."
        )

    elif higher_rate >= 30:

        insights.append(
            f"{higher_rate:.2f}% of assessed children "
            f"are at higher mathematics levels (M6 or M7), "
            f"indicating a substantial higher-level group."
        )

    else:

        insights.append(
            f"Only {higher_rate:.2f}% of assessed children "
            f"are at higher mathematics levels (M6 or M7), "
            f"indicating an area that may require "
            f"additional learning support."
        )


    # ==================================================
    # LOWER LEVEL PERFORMANCE
    # ==================================================

    if lower_rate >= 40:

        insights.append(
            f"{lower_rate:.2f}% of assessed children "
            f"remain at lower mathematics levels "
            f"(M1 to M3), suggesting a need for "
            f"targeted foundational mathematics support."
        )

    elif lower_rate >= 25:

        insights.append(
            f"{lower_rate:.2f}% of assessed children "
            f"are at lower mathematics levels (M1 to M3), "
            f"which should continue to be monitored."
        )

    else:

        insights.append(
            f"Lower mathematics levels account for "
            f"{lower_rate:.2f}% of assessed children."
        )


    # ==================================================
    # STRONGEST INDIVIDUAL LEVEL
    # ==================================================

    if level_rates:

        strongest_level = max(
            level_rates,
            key=level_rates.get
        )

        strongest_rate = level_rates[
            strongest_level
        ]

        insights.append(
            f"{strongest_level} has the highest "
            f"representation among mathematics levels "
            f"at {strongest_rate:.2f}%."
        )


    return insights


# ==================================================
# MAIN
# ==================================================

if __name__ == "__main__":

    print("\n-----------------------------")
    print("MATHEMATICS INSIGHTS")
    print("-----------------------------")

    kpi_data = load_kpi_data()

    insights = generate_mathematics_insights(
        kpi_data
    )

    for number, insight in enumerate(
        insights,
        start=1
    ):

        print(
            f"{number}. {insight}"
        )