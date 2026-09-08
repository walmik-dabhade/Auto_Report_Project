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
# GENERATE ACTIVITY INSIGHTS
# ==================================================

def generate_activity_insights(kpi_data):

    insights = []

    activities = kpi_data.get(
        "activity",
        {}
    )

    if not activities:

        insights.append(
            "No activity KPI data is available."
        )

        return insights


    # ==================================================
    # CREATE DATA LIST
    # ==================================================

    activity_list = []

    for activity, values in activities.items():

        activity_list.append({

            "name": activity,

            "total":
                values.get("total", 0),

            "continued":
                values.get("continued", 0),

            "dropout":
                values.get("dropout", 0),

            "migrated":
                values.get("migrated", 0),

            "graduate":
                values.get("graduate", 0),

            "continuation_rate":
                values.get(
                    "continuation_rate",
                    0
                ),

            "dropout_rate":
                values.get(
                    "dropout_rate",
                    0
                ),

            "migration_rate":
                values.get(
                    "migration_rate",
                    0
                )

        })


    # ==================================================
    # HIGHEST ENROLMENT ACTIVITY
    # ==================================================

    highest_activity = max(
        activity_list,
        key=lambda x: x["total"]
    )

    insights.append(
        f"{highest_activity['name']} has the highest "
        f"number of children, with "
        f"{highest_activity['total']:,} children."
    )


    # ==================================================
    # LOWEST ENROLMENT ACTIVITY
    # ==================================================

    lowest_activity = min(
        activity_list,
        key=lambda x: x["total"]
    )

    insights.append(
        f"{lowest_activity['name']} has the lowest "
        f"number of children among the recorded "
        f"activities, with "
        f"{lowest_activity['total']:,} children."
    )


    # ==================================================
    # HIGHEST CONTINUATION RATE
    # ==================================================

    highest_continuation = max(
        activity_list,
        key=lambda x: x["continuation_rate"]
    )

    insights.append(
        f"{highest_continuation['name']} has the highest "
        f"continuation rate at "
        f"{highest_continuation['continuation_rate']:.2f}%."
    )


    # ==================================================
    # LOWEST CONTINUATION RATE
    # ==================================================

    lowest_continuation = min(
        activity_list,
        key=lambda x: x["continuation_rate"]
    )

    insights.append(
        f"{lowest_continuation['name']} has the lowest "
        f"continuation rate at "
        f"{lowest_continuation['continuation_rate']:.2f}%."
    )


    # ==================================================
    # HIGHEST DROPOUT RATE
    # ==================================================

    highest_dropout = max(
        activity_list,
        key=lambda x: x["dropout_rate"]
    )

    insights.append(
        f"{highest_dropout['name']} has the highest "
        f"dropout rate at "
        f"{highest_dropout['dropout_rate']:.2f}%."
    )


    # ==================================================
    # HIGHEST MIGRATION RATE
    # ==================================================

    highest_migration = max(
        activity_list,
        key=lambda x: x["migration_rate"]
    )

    insights.append(
        f"{highest_migration['name']} has the highest "
        f"migration rate at "
        f"{highest_migration['migration_rate']:.2f}%."
    )


    # ==================================================
    # DROPOUT RISK
    # ==================================================

    high_dropout_activities = [

        activity

        for activity in activity_list

        if activity["dropout_rate"] >= 15

    ]


    if high_dropout_activities:

        names = ", ".join(
            activity["name"]
            for activity
            in high_dropout_activities
        )

        insights.append(
            f"The following activities have dropout "
            f"rates of 15% or higher and may require "
            f"further review: {names}."
        )


    # ==================================================
    # MIGRATION RISK
    # ==================================================

    high_migration_activities = [

        activity

        for activity in activity_list

        if activity["migration_rate"] >= 15

    ]


    if high_migration_activities:

        names = ", ".join(
            activity["name"]
            for activity
            in high_migration_activities
        )

        insights.append(
            f"The following activities have migration "
            f"rates of 15% or higher and may require "
            f"further analysis: {names}."
        )


    # ==================================================
    # STRONG CONTINUATION
    # ==================================================

    strong_activities = [

        activity

        for activity in activity_list

        if activity["continuation_rate"] >= 90

    ]


    if strong_activities:

        names = ", ".join(
            activity["name"]
            for activity
            in strong_activities
        )

        insights.append(
            f"The following activities show continuation "
            f"rates of 90% or higher, indicating strong "
            f"continuation: {names}."
        )


    return insights


# ==================================================
# MAIN
# ==================================================

if __name__ == "__main__":

    print("\n-----------------------------")
    print("ACTIVITY INSIGHTS")
    print("-----------------------------")

    kpi_data = load_kpi_data()

    insights = generate_activity_insights(
        kpi_data
    )

    for number, insight in enumerate(
        insights,
        start=1
    ):

        print(
            f"{number}. {insight}"
        )