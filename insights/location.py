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
# GENERATE LOCATION INSIGHTS
# ==================================================

def generate_location_insights(kpi_data):

    insights = []

    location_data = kpi_data.get(
        "location",
        {}
    )

    locations = location_data.get(
        "location",
        {}
    )

    areas = location_data.get(
        "area",
        {}
    )

    location_types = location_data.get(
        "location_type",
        {}
    )

    location_status = location_data.get(
        "location_status",
        {}
    )


    # ==================================================
    # CHECK DATA
    # ==================================================

    if not locations:

        insights.append(
            "No location KPI data is available."
        )

        return insights


    # ==================================================
    # PREPARE LOCATION DATA
    # ==================================================

    location_list = []

    for location, values in location_status.items():

        location_list.append({

            "name": location,

            "total":
                values.get(
                    "total",
                    0
                ),

            "continued":
                values.get(
                    "continued",
                    0
                ),

            "dropout":
                values.get(
                    "dropout",
                    0
                ),

            "migrated":
                values.get(
                    "migrated",
                    0
                ),

            "graduate":
                values.get(
                    "graduate",
                    0
                ),

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
    # LARGEST LOCATION
    # ==================================================

    largest = max(
        location_list,
        key=lambda x: x["total"]
    )

    insights.append(
        f"{largest['name']} has the highest number "
        f"of children, with "
        f"{largest['total']:,} children."
    )


    # ==================================================
    # SMALLEST LOCATION
    # ==================================================

    smallest = min(
        location_list,
        key=lambda x: x["total"]
    )

    insights.append(
        f"{smallest['name']} has the lowest number "
        f"of children, with "
        f"{smallest['total']:,} children."
    )


    # ==================================================
    # HIGHEST CONTINUATION
    # ==================================================

    highest_continuation = max(
        location_list,
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
        location_list,
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
        location_list,
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
        location_list,
        key=lambda x: x["migration_rate"]
    )

    insights.append(
        f"{highest_migration['name']} has the highest "
        f"migration rate at "
        f"{highest_migration['migration_rate']:.2f}%."
    )


    # ==================================================
    # HIGH DROPOUT LOCATIONS
    # ==================================================

    high_dropout_locations = [

        location

        for location in location_list

        if location["dropout_rate"] >= 20

    ]


    if high_dropout_locations:

        names = ", ".join(
            location["name"]
            for location
            in high_dropout_locations
        )

        insights.append(
            f"The following locations have dropout "
            f"rates of 20% or higher and may require "
            f"priority review: {names}."
        )


    # ==================================================
    # HIGH MIGRATION LOCATIONS
    # ==================================================

    high_migration_locations = [

        location

        for location in location_list

        if location["migration_rate"] >= 15

    ]


    if high_migration_locations:

        names = ", ".join(
            location["name"]
            for location
            in high_migration_locations
        )

        insights.append(
            f"The following locations have migration "
            f"rates of 15% or higher and may require "
            f"further analysis: {names}."
        )


    # ==================================================
    # STRONG CONTINUATION LOCATIONS
    # ==================================================

    strong_locations = [

        location

        for location in location_list

        if location["continuation_rate"] >= 90

    ]


    if strong_locations:

        names = ", ".join(
            location["name"]
            for location
            in strong_locations
        )

        insights.append(
            f"The following locations show continuation "
            f"rates of 90% or higher, indicating strong "
            f"continuation performance: {names}."
        )


    # ==================================================
    # AREA DISTRIBUTION
    # ==================================================

    if areas:

        largest_area = max(
            areas.items(),
            key=lambda x: x[1].get(
                "total",
                0
            )
        )

        insights.append(
            f"{largest_area[0]} represents the largest "
            f"area group, with "
            f"{largest_area[1].get('total', 0):,} children "
            f"({largest_area[1].get('percentage', 0):.2f}%)."
        )


    # ==================================================
    # LOCATION TYPE
    # ==================================================

    if location_types:

        largest_type = max(
            location_types.items(),
            key=lambda x: x[1].get(
                "total",
                0
            )
        )

        insights.append(
            f"{largest_type[0]} accounts for "
            f"{largest_type[1].get('percentage', 0):.2f}% "
            f"of the recorded location types."
        )


    return insights


# ==================================================
# MAIN
# ==================================================

if __name__ == "__main__":

    print("\n-----------------------------")
    print("LOCATION INSIGHTS")
    print("-----------------------------")

    kpi_data = load_kpi_data()

    insights = generate_location_insights(
        kpi_data
    )

    for number, insight in enumerate(
        insights,
        start=1
    ):

        print(
            f"{number}. {insight}"
        )