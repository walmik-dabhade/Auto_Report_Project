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
# GENERATE FUNDER INSIGHTS
# ==================================================

def generate_funder_insights(kpi_data):

    insights = []

    funders = kpi_data.get(
        "funder",
        {}
    )

    if not funders:

        insights.append(
            "No funder KPI data is available."
        )

        return insights


    # ==================================================
    # PREPARE DATA
    # ==================================================

    funder_list = []

    for funder, values in funders.items():

        funder_list.append({

            "name": funder,

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
    # LARGEST FUNDER
    # ==================================================

    largest = max(
        funder_list,
        key=lambda x: x["total"]
    )

    insights.append(
        f"{largest['name']} supports the highest "
        f"number of children, with "
        f"{largest['total']:,} children."
    )


    # ==================================================
    # SMALLEST FUNDER
    # ==================================================

    smallest = min(
        funder_list,
        key=lambda x: x["total"]
    )

    insights.append(
        f"{smallest['name']} has the smallest "
        f"number of children among the recorded "
        f"funders, with "
        f"{smallest['total']:,} children."
    )


    # ==================================================
    # HIGHEST CONTINUATION
    # ==================================================

    highest_continuation = max(
        funder_list,
        key=lambda x: x["continuation_rate"]
    )

    insights.append(
        f"{highest_continuation['name']} has the highest "
        f"continuation rate at "
        f"{highest_continuation['continuation_rate']:.2f}%."
    )


    # ==================================================
    # LOWEST CONTINUATION
    # ==================================================

    lowest_continuation = min(
        funder_list,
        key=lambda x: x["continuation_rate"]
    )

    insights.append(
        f"{lowest_continuation['name']} has the lowest "
        f"continuation rate at "
        f"{lowest_continuation['continuation_rate']:.2f}%."
    )


    # ==================================================
    # HIGHEST DROPOUT
    # ==================================================

    highest_dropout = max(
        funder_list,
        key=lambda x: x["dropout_rate"]
    )

    insights.append(
        f"{highest_dropout['name']} has the highest "
        f"dropout rate at "
        f"{highest_dropout['dropout_rate']:.2f}%."
    )


    # ==================================================
    # HIGHEST MIGRATION
    # ==================================================

    highest_migration = max(
        funder_list,
        key=lambda x: x["migration_rate"]
    )

    insights.append(
        f"{highest_migration['name']} has the highest "
        f"migration rate at "
        f"{highest_migration['migration_rate']:.2f}%."
    )


    # ==================================================
    # HIGH DROPOUT FUNDERS
    # ==================================================

    high_dropout = [

        funder

        for funder in funder_list

        if funder["dropout_rate"] >= 15

    ]


    if high_dropout:

        names = ", ".join(
            funder["name"]
            for funder in high_dropout
        )

        insights.append(
            f"The following funders have dropout rates "
            f"of 15% or higher and may require further "
            f"review: {names}."
        )


    # ==================================================
    # HIGH MIGRATION FUNDERS
    # ==================================================

    high_migration = [

        funder

        for funder in funder_list

        if funder["migration_rate"] >= 15

    ]


    if high_migration:

        names = ", ".join(
            funder["name"]
            for funder in high_migration
        )

        insights.append(
            f"The following funders have migration rates "
            f"of 15% or higher and may require further "
            f"analysis: {names}."
        )


    # ==================================================
    # STRONG FUNDERS
    # ==================================================

    strong_funders = [

        funder

        for funder in funder_list

        if funder["continuation_rate"] >= 85

    ]


    if strong_funders:

        names = ", ".join(
            funder["name"]
            for funder in strong_funders
        )

        insights.append(
            f"The following funders have continuation "
            f"rates of 85% or higher, indicating strong "
            f"continuation performance: {names}."
        )


    return insights


# ==================================================
# MAIN
# ==================================================

if __name__ == "__main__":

    print("\n-----------------------------")
    print("FUNDER INSIGHTS")
    print("-----------------------------")

    kpi_data = load_kpi_data()

    insights = generate_funder_insights(
        kpi_data
    )

    for number, insight in enumerate(
        insights,
        start=1
    ):

        print(
            f"{number}. {insight}"
        )