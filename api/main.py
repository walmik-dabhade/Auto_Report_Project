from pathlib import Path
import sys
import os
import shutil
import tempfile
import subprocess
from datetime import date
from typing import Literal

from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from pydantic import BaseModel
from dotenv import load_dotenv
import psycopg


# =========================================================
# PROJECT ROOT
# =========================================================

ROOT = Path(__file__).resolve().parents[1]

load_dotenv(ROOT / ".env")


# =========================================================
# PREVENT MATPLOTLIB GUI / TKINTER
# =========================================================

os.environ["MPLBACKEND"] = "Agg"


# =========================================================
# IMPORT GOOGLE DRIVE MODULE
# =========================================================

sys.path.insert(
    0,
    str(ROOT / "api")
)

import google_drive


# =========================================================
# IMPORT EXISTING V5 REPORT GENERATOR
# =========================================================

sys.path.insert(
    0,
    str(ROOT / "report")
)

import funder_report_generator_v5 as report_generator


# =========================================================
# FASTAPI
# =========================================================

app = FastAPI(
    title="CSR Report Generation API",
    description=(
        "API for generating funder-wise CSR progress reports "
        "using Google Drive inputs and V5 report generation."
    ),
    version="2.0.0"
)


# =========================================================
# REQUEST MODELS
# =========================================================

class Duration(BaseModel):

    from_date: date
    to_date: date


class ReportRequest(BaseModel):

    funder: str

    duration: Duration

    format: Literal[
        "pdf",
        "docx"
    ] = "pdf"


# =========================================================
# POSTGRESQL CONNECTION
# =========================================================

def get_db_connection():

    return psycopg.connect(
        host=os.getenv("DB_HOST"),
        port=os.getenv("DB_PORT"),
        dbname=os.getenv("DB_NAME"),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASSWORD")
    )


# =========================================================
# SAVE REPORT METADATA TO POSTGRESQL
# =========================================================

def save_report_record(
    funder: str,
    from_date: date,
    to_date: date,
    output_format: str,
    status: str,
    file_name: str | None = None
):

    conn = None

    try:

        conn = get_db_connection()

        with conn.cursor() as cur:

            cur.execute(
                """
                INSERT INTO reports
                (
                    funder,
                    from_date,
                    to_date,
                    format,
                    status,
                    file_name
                )
                VALUES
                (
                    %s,
                    %s,
                    %s,
                    %s,
                    %s,
                    %s
                )
                """,
                (
                    funder,
                    from_date,
                    to_date,
                    output_format,
                    status,
                    file_name
                )
            )

        conn.commit()

    except Exception as e:

        print(
            "WARNING: Could not save report metadata "
            f"to PostgreSQL: {e}"
        )

    finally:

        if conn is not None:

            conn.close()


# =========================================================
# HEALTH CHECK
# =========================================================

@app.get("/health")
def health_check():

    db_status = "unknown"

    try:

        conn = get_db_connection()

        with conn.cursor() as cur:

            cur.execute(
                "SELECT 1"
            )

            cur.fetchone()

        conn.close()

        db_status = "connected"

    except Exception as e:

        db_status = "disconnected"

        print(
            f"PostgreSQL health check failed: {e}"
        )

    return {
        "status": "ok",
        "service": "CSR Report Generation API",
        "database": db_status
    }


# =========================================================
# RUN EXISTING MARATHI -> ENGLISH TRANSLATION
# =========================================================

def run_translation():

    translation_script = (
        ROOT /
        "translation" /
        "marathi_to_english.py"
    )

    if not translation_script.exists():

        raise RuntimeError(
            "Translation script not found:\n"
            f"{translation_script}"
        )

    print(
        "\n" + "=" * 60
    )

    print(
        "STARTING MARATHI -> ENGLISH TRANSLATION"
    )

    print(
        "=" * 60
    )

    print(
        "\nTranslation script:"
    )

    print(
        translation_script
    )

    # -----------------------------------------------------
    # IMPORTANT WINDOWS UTF-8 FIX
    #
    # The existing translation script contains characters
    # such as:
    #
    # Marathi → English
    #
    # Windows may otherwise use cp1252 when the script is
    # executed as a subprocess.
    # -----------------------------------------------------

    translation_env = os.environ.copy()

    translation_env[
        "PYTHONIOENCODING"
    ] = "utf-8"

    translation_env[
        "PYTHONUTF8"
    ] = "1"

    # -----------------------------------------------------
    # Run existing translation script
    # -----------------------------------------------------

    result = subprocess.run(
        [
            sys.executable,
            str(translation_script)
        ],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=translation_env
    )

    # -----------------------------------------------------
    # Print standard output
    # -----------------------------------------------------

    if result.stdout:

        print(
            result.stdout
        )

    # -----------------------------------------------------
    # Print errors / warnings
    # -----------------------------------------------------

    if result.stderr:

        print(
            "Translation stderr:"
        )

        print(
            result.stderr
        )

    # -----------------------------------------------------
    # Check translation process
    # -----------------------------------------------------

    if result.returncode != 0:

        raise RuntimeError(
            "Marathi -> English translation failed.\n"
            f"{result.stderr or result.stdout}"
        )

    # -----------------------------------------------------
    # Verify translated JSON
    # -----------------------------------------------------

    translated_json = (
        ROOT /
        "output" /
        "translated_report.json"
    )

    if not translated_json.exists():

        raise RuntimeError(
            "Translation completed but "
            "translated_report.json was not created."
        )

    if translated_json.stat().st_size == 0:

        raise RuntimeError(
            "translated_report.json was created "
            "but is empty."
        )

    print(
        "\nTranslation completed successfully."
    )

    print(
        "Translation output:"
    )

    print(
        translated_json
    )

    return translated_json


# =========================================================
# CONFIGURE V5 EXCEL INPUT
# =========================================================

def configure_v5_inputs(
    excel_path: Path
):

    if not excel_path.exists():

        raise RuntimeError(
            "Google Drive Excel file does not exist:\n"
            f"{excel_path}"
        )

    # -----------------------------------------------------
    # Point existing V5 generator to Google Drive input.
    #
    # No V5 design/layout changes are made.
    # -----------------------------------------------------

    report_generator.INPUT_XLSX = (
        excel_path
    )

    print(
        "\nV5 Excel input configured:"
    )

    print(
        report_generator.INPUT_XLSX
    )


# =========================================================
# DOCX -> PDF USING MICROSOFT WORD
# =========================================================

def convert_docx_to_pdf(
    docx_path: Path,
    pdf_path: Path
):

    try:

        import pythoncom
        import win32com.client

    except ImportError:

        raise RuntimeError(
            "pywin32 is not installed. "
            "Run: pip install pywin32"
        )

    word = None

    document = None

    pythoncom.CoInitialize()

    try:

        word = (
            win32com.client.DispatchEx(
                "Word.Application"
            )
        )

        word.Visible = False

        word.DisplayAlerts = 0

        document = word.Documents.Open(
            str(
                docx_path.resolve()
            ),
            ReadOnly=True,
            AddToRecentFiles=False
        )

        document.ExportAsFixedFormat(
            OutputFileName=str(
                pdf_path.resolve()
            ),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            CreateBookmarks=0
        )

        document.Close(
            SaveChanges=False
        )

        document = None

        word.Quit()

        word = None

        if not pdf_path.exists():

            raise RuntimeError(
                "Microsoft Word completed conversion "
                "but PDF was not created."
            )

    except Exception as e:

        if document is not None:

            try:

                document.Close(
                    SaveChanges=False
                )

            except Exception:

                pass

        if word is not None:

            try:

                word.Quit()

            except Exception:

                pass

        raise RuntimeError(
            "Microsoft Word PDF conversion failed: "
            f"{e}"
        )

    finally:

        pythoncom.CoUninitialize()


# =========================================================
# GENERATE REPORT
# =========================================================

@app.post("/generate-report")
def generate_report(
    request: ReportRequest,
    background_tasks: BackgroundTasks
):

    # -----------------------------------------------------
    # Normalize request
    # -----------------------------------------------------

    funder = (
        request.funder
        .strip()
        .upper()
    )

    output_format = (
        request.format
        .lower()
    )

    from_date = (
        request.duration.from_date
    )

    to_date = (
        request.duration.to_date
    )

    # =====================================================
    # VALIDATE FUNDER
    # =====================================================

    if funder not in report_generator.FUNDERS:

        raise HTTPException(
            status_code=400,
            detail={
                "message": "Invalid funder",
                "allowed_funders": (
                    report_generator.FUNDERS
                )
            }
        )

    # =====================================================
    # VALIDATE DATE ORDER
    # =====================================================

    if from_date > to_date:

        raise HTTPException(
            status_code=400,
            detail={
                "message": (
                    "from_date cannot be later "
                    "than to_date."
                )
            }
        )

    # =====================================================
    # CURRENT V5 SUPPORTED PERIOD
    # =====================================================

    expected_from = date(
        2025,
        10,
        1
    )

    expected_to = date(
        2025,
        12,
        31
    )

    if (
        from_date != expected_from
        or
        to_date != expected_to
    ):

        raise HTTPException(
            status_code=400,
            detail={
                "message": (
                    "The current V5 report generator "
                    "supports only October-December 2025."
                ),
                "supported_duration": {
                    "from": str(
                        expected_from
                    ),
                    "to": str(
                        expected_to
                    )
                }
            }
        )

    # =====================================================
    # TEMPORARY PDF DIRECTORY
    # =====================================================

    temp_dir = None

    try:

        # =================================================
        # STEP 1
        # GOOGLE DRIVE
        # =================================================

        print(
            "\n" + "=" * 60
        )

        print(
            "STEP 1: DOWNLOADING INPUTS FROM GOOGLE DRIVE"
        )

        print(
            "=" * 60
        )

        drive_files = (
            google_drive.download_csr_files()
        )

        excel_path = Path(
            drive_files["excel"]
        )

        word_path = Path(
            drive_files["word"]
        )

        # -------------------------------------------------
        # Verify files
        # -------------------------------------------------

        if not excel_path.exists():

            raise RuntimeError(
                "Google Drive Excel file "
                "was not found."
            )

        if not word_path.exists():

            raise RuntimeError(
                "Google Drive Marathi report "
                "was not found."
            )

        print(
            "\nGoogle Drive inputs ready."
        )

        print(
            f"Excel: {excel_path}"
        )

        print(
            f"Word:  {word_path}"
        )

        # =================================================
        # STEP 2
        # GEMINI TRANSLATION
        # =================================================

        print(
            "\n" + "=" * 60
        )

        print(
            "STEP 2: TRANSLATING MARATHI REPORT"
        )

        print(
            "=" * 60
        )

        run_translation()

        # =================================================
        # STEP 3
        # CONFIGURE V5 INPUT
        # =================================================

        print(
            "\n" + "=" * 60
        )

        print(
            "STEP 3: CONFIGURING V5 REPORT INPUT"
        )

        print(
            "=" * 60
        )

        configure_v5_inputs(
            excel_path
        )

        # =================================================
        # STEP 4
        # LOAD V5 DATA
        # =================================================

        print(
            "\n" + "=" * 60
        )

        print(
            "STEP 4: LOADING V5 DATA"
        )

        print(
            "=" * 60
        )

        all_df = (
            report_generator.load_data()
        )

        translated = (
            report_generator.load_translation()
        )

        # =================================================
        # STEP 5
        # GENERATE V5 DOCX
        # =================================================

        print(
            "\n" + "=" * 60
        )

        print(
            f"STEP 5: GENERATING "
            f"{funder} V5 REPORT"
        )

        print(
            "=" * 60
        )

        docx_path = (
            report_generator.build_report(
                funder,
                all_df,
                translated
            )
        )

        docx_path = Path(
            docx_path
        )

        if not docx_path.exists():

            raise RuntimeError(
                "V5 generator completed but "
                "DOCX file was not found."
            )

        print(
            "\nV5 DOCX generated:"
        )

        print(
            docx_path
        )

        # =================================================
        # STEP 6A
        # DOCX REQUEST
        # =================================================

        if output_format == "docx":

            output_file_name = (
                f"{funder}_Progress_Report.docx"
            )

            save_report_record(
                funder=funder,
                from_date=from_date,
                to_date=to_date,
                output_format="docx",
                status="SUCCESS",
                file_name=output_file_name
            )

            return FileResponse(
                path=str(
                    docx_path
                ),
                media_type=(
                    "application/"
                    "vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                ),
                filename=output_file_name
            )

        # =================================================
        # STEP 6B
        # PDF REQUEST
        # =================================================

        temp_dir = Path(
            tempfile.mkdtemp(
                prefix="csr_report_"
            )
        )

        pdf_path = (
            temp_dir /
            f"{funder}_Progress_Report.pdf"
        )

        print(
            "\n" + "=" * 60
        )

        print(
            "STEP 6: CONVERTING DOCX -> PDF"
        )

        print(
            "=" * 60
        )

        convert_docx_to_pdf(
            docx_path,
            pdf_path
        )

        if not pdf_path.exists():

            raise RuntimeError(
                "PDF conversion completed but "
                "PDF file was not found."
            )

        output_file_name = (
            f"{funder}_Progress_Report.pdf"
        )

        # =================================================
        # SAVE SUCCESS METADATA
        # =================================================

        save_report_record(
            funder=funder,
            from_date=from_date,
            to_date=to_date,
            output_format="pdf",
            status="SUCCESS",
            file_name=output_file_name
        )

        # =================================================
        # CLEAN TEMP DIRECTORY AFTER RESPONSE
        # =================================================

        background_tasks.add_task(
            shutil.rmtree,
            str(temp_dir),
            ignore_errors=True
        )

        # =================================================
        # RETURN PDF
        # =================================================

        return FileResponse(
            path=str(
                pdf_path
            ),
            media_type="application/pdf",
            filename=output_file_name
        )

    # =====================================================
    # HTTP EXCEPTION
    # =====================================================

    except HTTPException:

        raise

    # =====================================================
    # GENERAL ERROR
    # =====================================================

    except Exception as e:

        # -------------------------------------------------
        # Save FAILED metadata
        # -------------------------------------------------

        save_report_record(
            funder=funder,
            from_date=from_date,
            to_date=to_date,
            output_format=output_format,
            status="FAILED",
            file_name=None
        )

        # -------------------------------------------------
        # Remove temporary directory
        # -------------------------------------------------

        if temp_dir is not None:

            shutil.rmtree(
                temp_dir,
                ignore_errors=True
            )

        # -------------------------------------------------
        # Return API error
        # -------------------------------------------------

        raise HTTPException(
            status_code=500,
            detail=(
                "Report generation failed: "
                f"{str(e)}"
            )
        )