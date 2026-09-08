import os
import json
import time
from pathlib import Path
from typing import List

from dotenv import load_dotenv
from docx import Document
from pydantic import BaseModel, Field
from google import genai
from google.genai import types


# ============================================================
# CONFIGURATION
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parent.parent

INPUT_FILE = PROJECT_ROOT / "ACS-HY Marathi Report.docx"
OUTPUT_FILE = PROJECT_ROOT / "output" / "translated_report.json"

MODEL_NAME = "gemini-3.1-flash-lite"

MAX_RETRIES = 5


# ============================================================
# JSON STRUCTURE
# ============================================================

class ProgrammeOverview(BaseModel):
    programme_name: str = ""
    objective: str = ""
    target_community: str = ""
    key_context: str = ""


class ActivityStory(BaseModel):
    description: str = ""
    observed_changes: List[str] = Field(default_factory=list)
    implementation_context: List[str] = Field(default_factory=list)
    challenges: List[str] = Field(default_factory=list)


class LearningOutcome(BaseModel):
    story: str = ""
    observed_progress: List[str] = Field(default_factory=list)


class ParentEngagement(BaseModel):
    key_observations: List[str] = Field(default_factory=list)
    examples: List[str] = Field(default_factory=list)
    changes_observed: List[str] = Field(default_factory=list)


class TeacherTraining(BaseModel):
    topics: List[str] = Field(default_factory=list)
    implementation_changes: List[str] = Field(default_factory=list)


class SuccessStory(BaseModel):
    title: str = ""
    story: str = ""
    outcome: str = ""
    anonymized: bool = True


class ActivityStories(BaseModel):
    balwadi: ActivityStory = Field(default_factory=ActivityStory)
    study_class: ActivityStory = Field(default_factory=ActivityStory)
    reference_class: ActivityStory = Field(default_factory=ActivityStory)
    home_lending: ActivityStory = Field(default_factory=ActivityStory)
    library_class: ActivityStory = Field(default_factory=ActivityStory)


class LearningOutcomes(BaseModel):
    language: LearningOutcome = Field(default_factory=LearningOutcome)
    mathematics: LearningOutcome = Field(default_factory=LearningOutcome)


class TranslatedReport(BaseModel):
    programme_overview: ProgrammeOverview = Field(
        default_factory=ProgrammeOverview
    )

    activity_stories: ActivityStories = Field(
        default_factory=ActivityStories
    )

    learning_outcomes: LearningOutcomes = Field(
        default_factory=LearningOutcomes
    )

    parent_engagement: ParentEngagement = Field(
        default_factory=ParentEngagement
    )

    teacher_training: TeacherTraining = Field(
        default_factory=TeacherTraining
    )

    challenges: List[str] = Field(default_factory=list)

    future_planning: List[str] = Field(default_factory=list)

    success_stories: List[SuccessStory] = Field(
        default_factory=list
    )


# ============================================================
# READ DOCX
# ============================================================

def extract_docx_text(file_path: Path) -> str:

    if not file_path.exists():
        raise FileNotFoundError(
            f"Input file not found:\n{file_path}"
        )

    document = Document(file_path)

    sections = []

    # --------------------------------------------------------
    # Paragraphs
    # --------------------------------------------------------

    for paragraph in document.paragraphs:

        text = paragraph.text.strip()

        if text:
            sections.append(text)

    # --------------------------------------------------------
    # Tables
    # --------------------------------------------------------

    for table_index, table in enumerate(
        document.tables,
        start=1
    ):

        sections.append(
            f"\n[TABLE {table_index}]"
        )

        for row in table.rows:

            cells = []

            for cell in row.cells:

                cell_text = (
                    cell.text
                    .strip()
                    .replace("\n", " ")
                )

                cells.append(cell_text)

            sections.append(
                " | ".join(cells)
            )

    return "\n".join(sections)


# ============================================================
# GEMINI PROMPT
# ============================================================

def build_prompt(marathi_text: str) -> str:

    return f"""
You are an expert CSR programme-report translator,
qualitative research analyst and NGO impact-report writer.

You are processing a Marathi Community Education Programme
report.

Your task is NOT literal word-for-word translation.

You must understand the Marathi report and convert its
meaning into professional, natural English suitable for a
funder-facing CSR impact report.

============================================================
SOURCE RULES
============================================================

1. The Marathi report is the authoritative source for:
   - qualitative stories
   - observations
   - activity descriptions
   - implementation context
   - challenges
   - examples
   - parent engagement
   - teacher training
   - future plans
   - beneficiary stories

2. The cleaned Excel file is the authoritative source for
   numerical KPIs in the final report.

3. Do NOT invent statistics.

4. Do NOT calculate new statistics.

5. Do NOT estimate missing numbers.

6. Do NOT change the meaning of numbers appearing in the
   Marathi source.

7. If the Marathi report contains numerical information,
   preserve the information faithfully, but understand that
   the final KPI numbers will come from the cleaned Excel.

8. If information is absent, return an empty string or empty
   list.

9. Do not add facts based on general knowledge.

10. Do not make unsupported claims about programme impact.

============================================================
MAIN PURPOSE
============================================================

The final automated reporting system will combine:

CLEANED EXCEL
    -> authoritative numerical KPIs

MARATHI REPORT
    -> story and qualitative context

Therefore, identify the STORY BEHIND THE NUMBERS.

Look particularly for:

- changes observed in children
- learning improvements
- reading development
- language development
- mathematics development
- attendance and regularity observations
- activity implementation
- Balwadi experiences
- Study Class experiences
- Reference Class experiences
- Library experiences
- Home Lending experiences
- parent involvement
- teacher/supervisor training
- projects
- experiments
- festivals
- events
- challenges
- future plans
- meaningful beneficiary stories

============================================================
TRANSLATION STYLE
============================================================

Use clear, natural and professional English.

The output should sound like a professional CSR/NGO report,
not like machine translation.

Do not exaggerate.

Do not add emotional claims that are not present in the
source.

Preserve the actual meaning and context of the Marathi report.

============================================================
SUCCESS STORIES
============================================================

Extract meaningful beneficiary success stories.

Protect beneficiary privacy.

If a child's full name is present, anonymize the person in
the English output while preserving the substance of the
story.

For example:

"Sonu Tapure" -> "a 12-year-old boy"

Do not invent demographic details.

Do not remove important learning or behavioural changes
described in the source.

============================================================
OUTPUT
============================================================

Return ONLY valid JSON matching the supplied schema.

Do not return Markdown.

Do not return ```json.

Do not include explanations outside the JSON.

============================================================
MARATHI REPORT
============================================================

{marathi_text}

============================================================
END MARATHI REPORT
============================================================
"""


# ============================================================
# GEMINI API CALL WITH RETRY
# ============================================================

def translate_report(
    marathi_text: str
) -> TranslatedReport:

    load_dotenv(PROJECT_ROOT / ".env")

    api_key = os.getenv("GEMINI_API_KEY")

    if not api_key:
        raise ValueError(
            "GEMINI_API_KEY was not found in .env"
        )

    client = genai.Client(
        api_key=api_key
    )

    prompt = build_prompt(
        marathi_text
    )

    print("\nSending Marathi report to Gemini...")
    print(f"Model: {MODEL_NAME}")

    response = None

    for attempt in range(
        1,
        MAX_RETRIES + 1
    ):

        try:

            print(
                f"\nGemini attempt "
                f"{attempt}/{MAX_RETRIES}"
            )

            response = client.models.generate_content(
                model=MODEL_NAME,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=TranslatedReport,
                    temperature=0.2,
                ),
            )

            print(
                "\nGemini response received."
            )

            break

        except Exception as error:

            error_text = str(error)

            print(
                f"\nGemini request failed:"
            )

            print(
                error_text
            )

            if attempt == MAX_RETRIES:

                raise RuntimeError(
                    "\nGemini failed after "
                    f"{MAX_RETRIES} attempts.\n"
                    "Please try again later. "
                    "The API key and document "
                    "processing are working."
                ) from error

            # Exponential backoff
            wait_seconds = 10 * (
                2 ** (attempt - 1)
            )

            print(
                f"\nTemporary Gemini failure."
            )

            print(
                f"Waiting {wait_seconds} seconds "
                "before retrying..."
            )

            time.sleep(
                wait_seconds
            )

    if response is None:
        raise RuntimeError(
            "No response received from Gemini."
        )

    if not response.text:

        raise RuntimeError(
            "Gemini returned an empty response."
        )

    # --------------------------------------------------------
    # Validate JSON
    # --------------------------------------------------------

    print(
        "\nValidating structured JSON..."
    )

    try:

        result = (
            TranslatedReport
            .model_validate_json(
                response.text
            )
        )

    except Exception as error:

        print(
            "\nGemini returned:"
        )

        print(
            response.text
        )

        raise RuntimeError(
            "\nGemini returned a response, "
            "but it could not be validated "
            "against our JSON structure."
        ) from error

    print(
        "JSON validation successful."
    )

    return result


# ============================================================
# SAVE JSON
# ============================================================

def save_json(
    result: TranslatedReport
):

    OUTPUT_FILE.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    with open(
        OUTPUT_FILE,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            result.model_dump(),
            file,
            ensure_ascii=False,
            indent=2
        )

    print(
        "\nTranslated JSON saved successfully:"
    )

    print(
        OUTPUT_FILE
    )


# ============================================================
# MAIN
# ============================================================

def main():

    print("=" * 60)
    print(
        "MARATHI → ENGLISH REPORT TRANSLATOR"
    )
    print("=" * 60)

    print(
        f"\nInput:\n{INPUT_FILE}"
    )

    # --------------------------------------------------------
    # Extract DOCX
    # --------------------------------------------------------

    marathi_text = extract_docx_text(
        INPUT_FILE
    )

    print(
        f"\nExtracted characters: "
        f"{len(marathi_text):,}"
    )

    if not marathi_text.strip():

        raise ValueError(
            "No text was extracted from "
            "the Marathi report."
        )

    # --------------------------------------------------------
    # Gemini translation + extraction
    # --------------------------------------------------------

    result = translate_report(
        marathi_text
    )

    # --------------------------------------------------------
    # Save
    # --------------------------------------------------------

    save_json(
        result
    )

    print("\n" + "=" * 60)
    print(
        "TRANSLATION COMPLETED SUCCESSFULLY"
    )
    print("=" * 60)

    print(
        "\nOutput file:"
    )

    print(
        "output\\translated_report.json"
    )


if __name__ == "__main__":
    main()