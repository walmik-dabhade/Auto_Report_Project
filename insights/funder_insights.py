import json
import os


# ============================================================
# PATHS
# ============================================================

BASE_DIR = os.path.dirname(
    os.path.dirname(os.path.abspath(__file__))
)

INPUT_DIR = os.path.join(
    BASE_DIR,
    "output",
    "funder_kpis"
)

OUTPUT_DIR = os.path.join(
    BASE_DIR,
    "output",
    "funder_insights"
)


# ============================================================
# HELPERS
# ============================================================

def save_json(data, path):

    os.makedirs(
        os.path.dirname(path),
        exist_ok=True
    )

    with open(
        path,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            data,
            file,
            indent=4,
            ensure_ascii=False
        )


def percentage(value):

    try:
        return f"{float(value):.2f}%"
    except:
        return "0.00%"


def get_nested(data, *keys, default=0):

    current = data

    for key in keys:

        if not isinstance(current, dict):
            return default

        current = current.get(
            key,
            default
        )

    return current


# ============================================================
# OVERALL INSIGHTS
# ============================================================

def generate_overall_insights(kpi, funder):

    overall = kpi.get(
        "overall",
        {}
    )

    status = kpi.get(
        "status",
        {}
    )

    total = overall.get(
        "total_records",
        0
    )

    unique = overall.get(
        "unique_valid_software_ids",
        0
    )

    male = overall.get(
        "male",
        0
    )

    female = overall.get(
        "female",
        0
    )

    continued = status.get(
        "continued",
        0
    )

    dropout = status.get(
        "dropout",
        0
    )

    migrated = status.get(
        "migrated",
        0
    )

    graduate = status.get(
        "graduate",
        0
    )

    continuation_rate = status.get(
        "continuation_rate",
        0
    )

    dropout_rate = status.get(
        "dropout_rate",
        0
    )

    migration_rate = status.get(
        "migration_rate",
        0
    )

    insights = []

    insights.append(
        f"{funder} supports {total:,} child records "
        f"representing {unique:,} unique valid Software IDs."
    )

    if female > male:

        female_share = (
            female / total * 100
            if total else 0
        )

        male_share = (
            male / total * 100
            if total else 0
        )

        insights.append(
            f"Female children represent the larger share "
            f"of the {funder} supported population at "
            f"{female_share:.2f}%, compared with "
            f"{male_share:.2f}% male children."
        )

    elif male > female:

        male_share = (
            male / total * 100
            if total else 0
        )

        female_share = (
            female / total * 100
            if total else 0
        )

        insights.append(
            f"Male children represent the larger share "
            f"of the {funder} supported population at "
            f"{male_share:.2f}%, compared with "
            f"{female_share:.2f}% female children."
        )

    insights.append(
        f"The latest status data shows that "
        f"{continued:,} children ({percentage(continuation_rate)}) "
        f"continued with the programme."
    )

    if dropout > 0:

        insights.append(
            f"{dropout:,} children ({percentage(dropout_rate)}) "
            f"are recorded as dropouts, indicating a need "
            f"for continued monitoring of programme continuity."
        )

    if migrated > 0:

        insights.append(
            f"{migrated:,} children ({percentage(migration_rate)}) "
            f"are recorded as migrated. Tracking migration-related "
            f"movement can help distinguish relocation from "
            f"programme disengagement."
        )

    if graduate > 0:

        insights.append(
            f"{graduate:,} children are recorded as graduates "
            f"in the latest status data."
        )

    return insights


# ============================================================
# ACTIVITY INSIGHTS
# ============================================================

def generate_activity_insights(kpi):

    activities = kpi.get(
        "activity",
        {}
    )

    if not activities:
        return []

    insights = []

    activity_items = list(
        activities.items()
    )

    # Highest reach
    highest = max(
        activity_items,
        key=lambda x: x[1].get(
            "total",
            0
        )
    )

    # Lowest reach
    lowest = min(
        activity_items,
        key=lambda x: x[1].get(
            "total",
            0
        )
    )

    insights.append(
        f"{highest[0]} has the highest reach, "
        f"with {highest[1].get('total', 0):,} children."
    )

    insights.append(
        f"{lowest[0]} has the lowest reach among "
        f"the recorded activities, with "
        f"{lowest[1].get('total', 0):,} children."
    )

    # Highest continuation
    highest_cont = max(
        activity_items,
        key=lambda x: x[1].get(
            "continuation_rate",
            0
        )
    )

    # Lowest continuation
    lowest_cont = min(
        activity_items,
        key=lambda x: x[1].get(
            "continuation_rate",
            0
        )
    )

    insights.append(
        f"{highest_cont[0]} has the highest continuation "
        f"rate at {percentage(highest_cont[1].get('continuation_rate', 0))}."
    )

    insights.append(
        f"{lowest_cont[0]} has the lowest continuation "
        f"rate at {percentage(lowest_cont[1].get('continuation_rate', 0))}."
    )

    # Highest dropout
    highest_dropout = max(
        activity_items,
        key=lambda x: x[1].get(
            "dropout_rate",
            0
        )
    )

    insights.append(
        f"{highest_dropout[0]} has the highest dropout "
        f"rate at {percentage(highest_dropout[1].get('dropout_rate', 0))}."
    )

    # Highest migration
    highest_migration = max(
        activity_items,
        key=lambda x: x[1].get(
            "migration_rate",
            0
        )
    )

    insights.append(
        f"{highest_migration[0]} has the highest migration "
        f"rate at {percentage(highest_migration[1].get('migration_rate', 0))}."
    )

    # Priority dropout activities
    high_dropout = [

        name

        for name, values
        in activity_items

        if values.get(
            "dropout_rate",
            0
        ) >= 15
    ]

    if high_dropout:

        insights.append(
            "Activities with dropout rates of 15% or "
            "higher that may require review: "
            + ", ".join(high_dropout)
            + "."
        )

    # Priority migration activities
    high_migration = [

        name

        for name, values
        in activity_items

        if values.get(
            "migration_rate",
            0
        ) >= 15
    ]

    if high_migration:

        insights.append(
            "Activities with migration rates of 15% "
            "or higher that may require analysis: "
            + ", ".join(high_migration)
            + "."
        )

    # Strong continuation
    strong_continuation = [

        name

        for name, values
        in activity_items

        if values.get(
            "continuation_rate",
            0
        ) >= 90
    ]

    if strong_continuation:

        insights.append(
            "Activities showing continuation rates of "
            "90% or higher: "
            + ", ".join(strong_continuation)
            + "."
        )

    return insights


# ============================================================
# ATTENDANCE INSIGHTS
# ============================================================

def generate_attendance_insights(kpi):

    attendance = kpi.get(
        "attendance",
        {}
    )

    if not attendance:
        return []

    total = attendance.get(
        "total",
        0
    )

    high = attendance.get(
        "high_attendance",
        0
    )

    medium = attendance.get(
        "medium_attendance",
        0
    )

    low = attendance.get(
        "low_attendance",
        0
    )

    high_rate = attendance.get(
        "high_attendance_rate",
        0
    )

    medium_rate = attendance.get(
        "medium_attendance_rate",
        0
    )

    low_rate = attendance.get(
        "low_attendance_rate",
        0
    )

    insights = []

    insights.append(
        f"Attendance data is available for "
        f"{total:,} children."
    )

    insights.append(
        f"High attendance is the dominant category, "
        f"with {high:,} children ({percentage(high_rate)})."
    )

    insights.append(
        f"{medium:,} children ({percentage(medium_rate)}) "
        f"fall within the medium attendance range."
    )

    insights.append(
        f"{low:,} children ({percentage(low_rate)}) "
        f"fall within the low attendance range and "
        f"may require targeted follow-up."
    )

    below_high = medium + low

    below_high_rate = (
        below_high / total * 100
        if total else 0
    )

    insights.append(
        f"Overall, {below_high:,} children "
        f"({below_high_rate:.2f}%) are below the "
        f"high-attendance category."
    )

    return insights


# ============================================================
# LANGUAGE INSIGHTS
# ============================================================

def generate_language_insights(kpi):

    language = kpi.get(
        "language",
        {}
    )

    if not language:
        return []

    assessed = language.get(
        "total_assessed",
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

    lower = language.get(
        "l1_l2",
        l1 + l2
    )

    higher = language.get(
        "l3_l4",
        l3 + l4
    )

    lower_rate = language.get(
        "l1_l2_rate",
        0
    )

    higher_rate = language.get(
        "l3_l4_rate",
        0
    )

    insights = []

    insights.append(
        f"Language assessment data is available for "
        f"{assessed:,} children."
    )

    if missing > 0:

        insights.append(
            f"{missing:,} children do not have a recorded "
            f"language assessment, indicating an assessment "
            f"coverage gap."
        )

    if assessed > 0:

        level_counts = {
            "L1": l1,
            "L2": l2,
            "L3": l3,
            "L4": l4
        }

        dominant_level = max(
            level_counts,
            key=level_counts.get
        )

        dominant_count = level_counts[
            dominant_level
        ]

        dominant_rate = (
            dominant_count /
            assessed *
            100
        )

        insights.append(
            f"{dominant_level} is the dominant language "
            f"level among assessed children, representing "
            f"{dominant_rate:.2f}%."
        )

        insights.append(
            f"Lower language levels (L1 + L2) account "
            f"for {lower:,} children ({percentage(lower_rate)}), "
            f"while higher levels (L3 + L4) account for "
            f"{higher:,} children ({percentage(higher_rate)})."
        )

    return insights


# ============================================================
# MATHEMATICS INSIGHTS
# ============================================================

def generate_mathematics_insights(kpi):

    maths = kpi.get(
        "mathematics",
        {}
    )

    if not maths:

        maths = kpi.get(
            "maths",
            {}
        )

    if not maths:
        return []

    assessed = maths.get(
        "total_assessed",
        0
    )

    missing = maths.get(
        "missing",
        0
    )

    low = maths.get(
        "m1_m2_m3",
        0
    )

    middle = maths.get(
        "m4_m5",
        0
    )

    higher = maths.get(
        "m6_m7",
        0
    )

    low_rate = maths.get(
        "m1_m2_m3_rate",
        0
    )

    middle_rate = maths.get(
        "m4_m5_rate",
        0
    )

    higher_rate = maths.get(
        "m6_m7_rate",
        0
    )

    insights = []

    insights.append(
        f"Mathematics assessment data is available for "
        f"{assessed:,} children."
    )

    if missing > 0:

        insights.append(
            f"{missing:,} children do not have a recorded "
            f"mathematics assessment, indicating a significant "
            f"assessment-data gap."
        )

    if assessed > 0:

        level_counts = {

            "M1": maths.get("m1", 0),
            "M2": maths.get("m2", 0),
            "M3": maths.get("m3", 0),
            "M4": maths.get("m4", 0),
            "M5": maths.get("m5", 0),
            "M6": maths.get("m6", 0),
            "M7": maths.get("m7", 0)

        }

        dominant_level = max(
            level_counts,
            key=level_counts.get
        )

        dominant_count = level_counts[
            dominant_level
        ]

        dominant_rate = (
            dominant_count /
            assessed *
            100
        )

        insights.append(
            f"{dominant_level} is the dominant mathematics "
            f"level among assessed children, representing "
            f"{dominant_rate:.2f}%."
        )

        insights.append(
            f"Lower mathematics levels (M1–M3) account "
            f"for {low:,} children ({percentage(low_rate)}), "
            f"middle levels (M4–M5) account for "
            f"{middle:,} ({percentage(middle_rate)}), "
            f"while higher levels (M6–M7) account for "
            f"{higher:,} ({percentage(higher_rate)})."
        )

    return insights


# ============================================================
# LOCATION INSIGHTS
# ============================================================

def generate_location_insights(kpi):

    location_data = kpi.get(
        "location",
        {}
    )

    locations = location_data.get(
        "locations",
        {}
    )

    if not locations:
        return []

    location_items = list(
        locations.items()
    )

    insights = []

    highest = max(
        location_items,
        key=lambda x: x[1].get(
            "total",
            0
        )
    )

    lowest = min(
        location_items,
        key=lambda x: x[1].get(
            "total",
            0
        )
    )

    highest_cont = max(
        location_items,
        key=lambda x: x[1].get(
            "continuation_rate",
            0
        )
    )

    lowest_cont = min(
        location_items,
        key=lambda x: x[1].get(
            "continuation_rate",
            0
        )
    )

    highest_dropout = max(
        location_items,
        key=lambda x: x[1].get(
            "dropout_rate",
            0
        )
    )

    highest_migration = max(
        location_items,
        key=lambda x: x[1].get(
            "migration_rate",
            0
        )
    )

    insights.append(
        f"{highest[0]} has the highest number of "
        f"children, with {highest[1].get('total', 0):,}."
    )

    insights.append(
        f"{lowest[0]} has the lowest number of "
        f"children, with {lowest[1].get('total', 0):,}."
    )

    insights.append(
        f"{highest_cont[0]} has the highest continuation "
        f"rate at {percentage(highest_cont[1].get('continuation_rate', 0))}."
    )

    insights.append(
        f"{lowest_cont[0]} has the lowest continuation "
        f"rate at {percentage(lowest_cont[1].get('continuation_rate', 0))}."
    )

    insights.append(
        f"{highest_dropout[0]} has the highest dropout "
        f"rate at {percentage(highest_dropout[1].get('dropout_rate', 0))}."
    )

    insights.append(
        f"{highest_migration[0]} has the highest migration "
        f"rate at {percentage(highest_migration[1].get('migration_rate', 0))}."
    )

    priority_dropout = [

        name

        for name, values
        in location_items

        if values.get(
            "dropout_rate",
            0
        ) >= 20
    ]

    if priority_dropout:

        insights.append(
            "Locations with dropout rates of 20% or "
            "higher requiring priority review: "
            + ", ".join(priority_dropout)
            + "."
        )

    priority_migration = [

        name

        for name, values
        in location_items

        if values.get(
            "migration_rate",
            0
        ) >= 15
    ]

    if priority_migration:

        insights.append(
            "Locations with migration rates of 15% "
            "or higher requiring further analysis: "
            + ", ".join(priority_migration)
            + "."
        )

    strong_locations = [

        name

        for name, values
        in location_items

        if values.get(
            "continuation_rate",
            0
        ) >= 90
    ]

    if strong_locations:

        insights.append(
            "Locations showing continuation rates of "
            "90% or higher: "
            + ", ".join(strong_locations)
            + "."
        )

    # --------------------------------------------------------
    # AREA
    # --------------------------------------------------------

    areas = location_data.get(
        "areas",
        {}
    )

    if areas:

        largest_area = max(
            areas.items(),
            key=lambda x: x[1]
        )

        total_area_records = sum(
            areas.values()
        )

        area_rate = (
            largest_area[1] /
            total_area_records *
            100
            if total_area_records
            else 0
        )

        insights.append(
            f"{largest_area[0]} represents the largest "
            f"area group, with {largest_area[1]:,} children "
            f"({area_rate:.2f}%)."
        )

    # --------------------------------------------------------
    # LOCATION TYPE
    # --------------------------------------------------------

    location_types = location_data.get(
        "location_types",
        {}
    )

    if location_types:

        largest_type = max(
            location_types.items(),
            key=lambda x: x[1]
        )

        total_type_records = sum(
            location_types.values()
        )

        type_rate = (
            largest_type[1] /
            total_type_records *
            100
            if total_type_records
            else 0
        )

        insights.append(
            f"{largest_type[0]} accounts for "
            f"{type_rate:.2f}% of recorded location types."
        )

    return insights


# ============================================================
# MASTER INSIGHT GENERATOR
# ============================================================

def generate_insights(kpi, funder):

    return {

        "funder_name":
            funder,

        "overall":
            generate_overall_insights(
                kpi,
                funder
            ),

        "activity":
            generate_activity_insights(
                kpi
            ),

        "attendance":
            generate_attendance_insights(
                kpi
            ),

        "language":
            generate_language_insights(
                kpi
            ),

        "mathematics":
            generate_mathematics_insights(
                kpi
            ),

        "location":
            generate_location_insights(
                kpi
            )
    }


# ============================================================
# PROCESS ALL FUNDERS
# ============================================================

def main():

    print()
    print("========================================")
    print("FUNDER INSIGHT ENGINE")
    print("========================================")

    if not os.path.exists(INPUT_DIR):

        print()
        print(
            "ERROR: Funder KPI folder not found:"
        )

        print(INPUT_DIR)

        return

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True
    )

    files = sorted(
        file

        for file in os.listdir(
            INPUT_DIR
        )

        if file.endswith(
            "_kpis.json"
        )
    )

    print()
    print(
        f"Funder KPI files found: {len(files)}"
    )

    for filename in files:

        input_file = os.path.join(
            INPUT_DIR,
            filename
        )

        with open(
            input_file,
            "r",
            encoding="utf-8"
        ) as file:

            kpi = json.load(file)

        funder = kpi.get(
            "funder_name",
            filename.replace(
                "_kpis.json",
                ""
            )
        )

        print()
        print(
            f"Processing: {funder}"
        )

        insights = generate_insights(
            kpi,
            funder
        )

        output_file = os.path.join(
            OUTPUT_DIR,
            f"{funder}_insights.json"
        )

        save_json(
            insights,
            output_file
        )

        print(
            f"Saved: {output_file}"
        )

    print()
    print("========================================")
    print("FUNDER INSIGHT GENERATION COMPLETED")
    print("========================================")

    print()
    print("Output folder:")
    print(OUTPUT_DIR)


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    main()