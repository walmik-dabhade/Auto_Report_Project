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
# GENERATE OVERALL INSIGHTS
# ==================================================

def generate_overall_insights(kpi_data):

    insights = []

    # --------------------------------------------------
    # DATA INFORMATION
    # --------------------------------------------------

    data = kpi_data["data_information"]

    total_records = (
        data["total_records"]
    )

    unique_ids = (
        data["unique_valid_software_ids"]
    )

    duplicate_difference = (
        total_records - unique_ids
    )


    insights.append(
        f"The dataset contains {total_records:,} "
        f"records representing {unique_ids:,} "
        f"unique valid Software IDs."
    )


    # --------------------------------------------------
    # DATA QUALITY
    # --------------------------------------------------

    if duplicate_difference > 0:

        insights.append(
            f"There is a difference of "
            f"{duplicate_difference:,} between total "
            f"records and unique valid Software IDs, "
            f"indicating that some records may require "
            f"duplicate or data-quality review."
        )

    else:

        insights.append(
            "Total records and unique valid Software "
            "IDs are aligned, indicating no difference "
            "at this level."
        )


    # --------------------------------------------------
    # GENDER
    # --------------------------------------------------

    gender = kpi_data["gender"]

    male = gender["male"]

    female = gender["female"]

    gender_total = male + female


    if gender_total > 0:

        male_rate = (
            male / gender_total * 100
        )

        female_rate = (
            female / gender_total * 100
        )

        if female > male:

            insights.append(
                f"Female children represent the larger "
                f"share of the dataset at "
                f"{female_rate:.2f}%, compared with "
                f"{male_rate:.2f}% male children."
            )

        elif male > female:

            insights.append(
                f"Male children represent the larger "
                f"share of the dataset at "
                f"{male_rate:.2f}%, compared with "
                f"{female_rate:.2f}% female children."
            )

        else:

            insights.append(
                "Male and female children are equally "
                "represented in the dataset."
            )


    # --------------------------------------------------
    # STATUS
    # --------------------------------------------------

    status = kpi_data["status"]

    continued = status["continued"]

    dropout = status["dropout"]

    migrated = status["migrated"]

    graduate = status["graduate"]

    status_total = (
        continued
        + dropout
        + migrated
        + graduate
    )


    if status_total > 0:

        continuation_rate = (
            continued
            / status_total
            * 100
        )

        dropout_rate = (
            dropout
            / status_total
            * 100
        )

        migration_rate = (
            migrated
            / status_total
            * 100
        )

        graduate_rate = (
            graduate
            / status_total
            * 100
        )


        # --------------------------------------------------
        # CONTINUATION
        # --------------------------------------------------

        insights.append(
            f"The latest status shows a "
            f"{continuation_rate:.2f}% continuation rate, "
            f"with {continued:,} children continuing."
        )


        # --------------------------------------------------
        # DROPOUT
        # --------------------------------------------------

        if dropout_rate >= 15:

            insights.append(
                f"The dropout rate is "
                f"{dropout_rate:.2f}%, which is relatively "
                f"high and should be reviewed to identify "
                f"locations, activities, or funders "
                f"contributing to dropout."
            )

        elif dropout_rate >= 10:

            insights.append(
                f"The dropout rate is "
                f"{dropout_rate:.2f}%, indicating an area "
                f"that should be monitored."
            )

        else:

            insights.append(
                f"The dropout rate is "
                f"{dropout_rate:.2f}%."
            )


        # --------------------------------------------------
        # MIGRATION
        # --------------------------------------------------

        if migration_rate >= 15:

            insights.append(
                f"The migration rate is "
                f"{migration_rate:.2f}%, which is significant "
                f"and warrants further analysis."
            )

        else:

            insights.append(
                f"The migration rate is "
                f"{migration_rate:.2f}%."
            )


        # --------------------------------------------------
        # GRADUATION
        # --------------------------------------------------

        insights.append(
            f"{graduate:,} children are recorded as "
            f"graduates in the latest status data."
        )


    return insights


# ==================================================
# MAIN
# ==================================================

if __name__ == "__main__":

    print("\n-----------------------------")
    print("OVERALL INSIGHTS")
    print("-----------------------------")

    kpi_data = load_kpi_data()

    insights = generate_overall_insights(
        kpi_data
    )

    for number, insight in enumerate(
        insights,
        start=1
    ):

        print(
            f"{number}. {insight}"
        )