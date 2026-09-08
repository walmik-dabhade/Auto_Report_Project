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
# GENERATE ATTENDANCE INSIGHTS
# ==================================================

def generate_attendance_insights(kpi_data):

    insights = []

    attendance = kpi_data.get(
        "attendance",
        {}
    )

    if not attendance:

        insights.append(
            "No attendance KPI data is available."
        )

        return insights


    # ==================================================
    # GET VALUES
    # ==================================================

    total = attendance.get(
        "total",
        0
    )

    high = attendance.get(
        "high",
        0
    )

    medium = attendance.get(
        "medium",
        0
    )

    low = attendance.get(
        "low",
        0
    )

    high_rate = attendance.get(
        "high_rate",
        0
    )

    medium_rate = attendance.get(
        "medium_rate",
        0
    )

    low_rate = attendance.get(
        "low_rate",
        0
    )


    # ==================================================
    # OVERALL ATTENDANCE
    # ==================================================

    insights.append(
        f"Attendance data is available for "
        f"{total:,} children."
    )


    # ==================================================
    # DOMINANT CATEGORY
    # ==================================================

    categories = {

        "High Attendance": high_rate,

        "Medium Attendance": medium_rate,

        "Low Attendance": low_rate
    }


    dominant_category = max(
        categories,
        key=categories.get
    )

    dominant_rate = categories[
        dominant_category
    ]


    insights.append(
        f"{dominant_category} is the dominant "
        f"attendance category, representing "
        f"{dominant_rate:.2f}% of children."
    )


    # ==================================================
    # HIGH ATTENDANCE
    # ==================================================

    if high_rate >= 70:

        insights.append(
            f"High attendance is strong at "
            f"{high_rate:.2f}%, with "
            f"{high:,} children achieving attendance "
            f"between 80% and 100%."
        )

    elif high_rate >= 50:

        insights.append(
            f"{high_rate:.2f}% of children have high "
            f"attendance between 80% and 100%, "
            f"indicating that more than half of the "
            f"children maintain strong attendance."
        )

    else:

        insights.append(
            f"Only {high_rate:.2f}% of children have "
            f"high attendance between 80% and 100%, "
            f"indicating a potential area for improvement."
        )


    # ==================================================
    # LOW ATTENDANCE
    # ==================================================

    if low_rate >= 20:

        insights.append(
            f"Low attendance is a significant concern, "
            f"with {low_rate:.2f}% of children having "
            f"attendance between 1% and 50%."
        )

    elif low_rate >= 10:

        insights.append(
            f"{low_rate:.2f}% of children fall into the "
            f"low attendance category, suggesting that "
            f"attendance should be monitored."
        )

    else:

        insights.append(
            f"Low attendance is relatively limited at "
            f"{low_rate:.2f}%."
        )


    # ==================================================
    # MEDIUM ATTENDANCE
    # ==================================================

    if medium_rate >= 30:

        insights.append(
            f"{medium_rate:.2f}% of children have medium "
            f"attendance between 51% and 79%, indicating "
            f"a substantial group that could potentially "
            f"be encouraged toward higher attendance."
        )

    else:

        insights.append(
            f"{medium_rate:.2f}% of children fall within "
            f"the medium attendance range."
        )


    # ==================================================
    # COMBINED BELOW-HIGH ATTENDANCE
    # ==================================================

    below_high_rate = (
        medium_rate
        + low_rate
    )

    below_high_count = (
        medium
        + low
    )


    insights.append(
        f"Overall, {below_high_count:,} children "
        f"({below_high_rate:.2f}%) are below the "
        f"high-attendance category."
    )


    return insights


# ==================================================
# MAIN
# ==================================================

if __name__ == "__main__":

    print("\n-----------------------------")
    print("ATTENDANCE INSIGHTS")
    print("-----------------------------")

    kpi_data = load_kpi_data()

    insights = generate_attendance_insights(
        kpi_data
    )

    for number, insight in enumerate(
        insights,
        start=1
    ):

        print(
            f"{number}. {insight}"
        )