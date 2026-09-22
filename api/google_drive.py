from pathlib import Path
import io

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload


# =========================================================
# PROJECT PATHS
# =========================================================

ROOT = Path(__file__).resolve().parents[1]

CREDENTIALS_FILE = ROOT / "credentials.json"
TOKEN_FILE = ROOT / "token.json"

DRIVE_INPUT_DIR = ROOT / "input" / "google_drive"

DRIVE_INPUT_DIR.mkdir(
    parents=True,
    exist_ok=True
)


# =========================================================
# GOOGLE DRIVE PERMISSIONS
# =========================================================

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly"
]


# =========================================================
# GOOGLE DRIVE CONNECTION
# =========================================================

def get_drive_service():

    creds = None

    # -----------------------------------------------------
    # Load existing OAuth token
    # -----------------------------------------------------

    if TOKEN_FILE.exists():

        creds = Credentials.from_authorized_user_file(
            str(TOKEN_FILE),
            SCOPES
        )

    # -----------------------------------------------------
    # Refresh token or authenticate
    # -----------------------------------------------------

    if not creds or not creds.valid:

        if (
            creds
            and creds.expired
            and creds.refresh_token
        ):

            print(
                "Refreshing Google Drive token..."
            )

            creds.refresh(
                Request()
            )

        else:

            print(
                "Google Drive authentication required..."
            )

            flow = (
                InstalledAppFlow
                .from_client_secrets_file(
                    str(CREDENTIALS_FILE),
                    SCOPES
                )
            )

            creds = flow.run_local_server(
                port=0
            )

        TOKEN_FILE.write_text(
            creds.to_json(),
            encoding="utf-8"
        )

    # -----------------------------------------------------
    # Build Google Drive API service
    # -----------------------------------------------------

    return build(
        "drive",
        "v3",
        credentials=creds,
        cache_discovery=False
    )


# =========================================================
# FIND FOLDER
# =========================================================

def find_folder(
    service,
    folder_name
):

    query = (
        f"name = '{folder_name}' "
        "and mimeType = "
        "'application/vnd.google-apps.folder' "
        "and trashed = false"
    )

    result = (
        service.files()
        .list(
            q=query,
            pageSize=100,
            fields="files(id,name,mimeType)"
        )
        .execute()
    )

    folders = result.get(
        "files",
        []
    )

    if not folders:

        raise RuntimeError(
            f"Google Drive folder not found: "
            f"{folder_name}"
        )

    if len(folders) > 1:

        print(
            f"WARNING: Multiple folders found "
            f"with name '{folder_name}'."
        )

        print(
            "Using the first matching folder."
        )

    return folders[0]


# =========================================================
# FIND FILE INSIDE FOLDER
# =========================================================

def find_file(
    service,
    folder_id,
    file_name
):

    query = (
        f"name = '{file_name}' "
        f"and '{folder_id}' in parents "
        "and trashed = false"
    )

    result = (
        service.files()
        .list(
            q=query,
            pageSize=100,
            fields=(
                "files("
                "id,"
                "name,"
                "mimeType,"
                "size,"
                "parents"
                ")"
            )
        )
        .execute()
    )

    files = result.get(
        "files",
        []
    )

    if not files:

        raise RuntimeError(
            "File not found in Google Drive folder: "
            f"{file_name}"
        )

    if len(files) > 1:

        print(
            f"WARNING: Multiple files found "
            f"with name '{file_name}'."
        )

        print(
            "Using the first matching file."
        )

    return files[0]


# =========================================================
# DOWNLOAD ACTUAL BINARY FILE
# =========================================================

def download_file(
    service,
    file_id,
    file_name,
    expected_mime_type=None
):

    output_path = (
        DRIVE_INPUT_DIR /
        file_name
    )

    print(
        "\n" + "-" * 60
    )

    print(
        f"Preparing download: {file_name}"
    )

    # -----------------------------------------------------
    # Get file metadata
    # -----------------------------------------------------

    metadata = (
        service.files()
        .get(
            fileId=file_id,
            fields=(
                "id,"
                "name,"
                "mimeType,"
                "size"
            )
        )
        .execute()
    )

    drive_name = metadata.get(
        "name",
        ""
    )

    drive_mime = metadata.get(
        "mimeType",
        ""
    )

    drive_size = metadata.get(
        "size"
    )

    print(
        f"Drive file: {drive_name}"
    )

    print(
        f"Drive MIME type: {drive_mime}"
    )

    print(
        "Drive size: "
        f"{drive_size if drive_size else 'unknown'} bytes"
    )

    # -----------------------------------------------------
    # Validate filename
    # -----------------------------------------------------

    if drive_name != file_name:

        raise RuntimeError(
            "Unexpected Google Drive file.\n"
            f"Expected: {file_name}\n"
            f"Received: {drive_name}"
        )

    # -----------------------------------------------------
    # Validate MIME type
    # -----------------------------------------------------

    if (
        expected_mime_type
        and drive_mime != expected_mime_type
    ):

        raise RuntimeError(
            "Unexpected MIME type.\n"
            f"Expected: {expected_mime_type}\n"
            f"Received: {drive_mime}"
        )

    # -----------------------------------------------------
    # IMPORTANT FIX
    #
    # get_media() requests the actual binary file.
    #
    # Do NOT use:
    #
    # service.files().get(
    #     fileId=file_id,
    #     alt="media"
    # )
    #
    # -----------------------------------------------------

    request = (
        service.files()
        .get_media(
            fileId=file_id
        )
    )

    # -----------------------------------------------------
    # Download into memory first
    # -----------------------------------------------------

    memory_file = io.BytesIO()

    downloader = MediaIoBaseDownload(
        memory_file,
        request,
        chunksize=1024 * 1024
    )

    done = False

    while not done:

        status, done = (
            downloader.next_chunk()
        )

        if status:

            progress = int(
                status.progress() * 100
            )

            print(
                f"Downloading {file_name}: "
                f"{progress}%"
            )

    # -----------------------------------------------------
    # Get downloaded bytes
    # -----------------------------------------------------

    file_bytes = (
        memory_file.getvalue()
    )

    downloaded_size = len(
        file_bytes
    )

    print(
        f"Downloaded size: "
        f"{downloaded_size:,} bytes"
    )

    # -----------------------------------------------------
    # Validate downloaded size
    # -----------------------------------------------------

    if downloaded_size < 1000:

        print(
            "\nERROR: Downloaded content "
            "is unexpectedly small."
        )

        try:

            print(
                "\nDownloaded content:"
            )

            print(
                file_bytes.decode(
                    "utf-8",
                    errors="replace"
                )
            )

        except Exception:

            pass

        raise RuntimeError(
            f"Google Drive returned invalid "
            f"content for {file_name}."
        )

    # -----------------------------------------------------
    # Validate Office file signature
    #
    # DOCX/XLSX are ZIP containers.
    # ZIP files normally begin with PK.
    # -----------------------------------------------------

    if file_name.lower().endswith(
        (".docx", ".xlsx")
    ):

        if not file_bytes.startswith(
            b"PK"
        ):

            print(
                "\nERROR: Downloaded content "
                "does not appear to be a valid "
                "DOCX/XLSX file."
            )

            try:

                preview = file_bytes[:500].decode(
                    "utf-8",
                    errors="replace"
                )

                print(
                    "\nDownloaded content preview:"
                )

                print(
                    preview
                )

            except Exception:

                pass

            raise RuntimeError(
                f"Downloaded {file_name}, but "
                "the content is not a valid "
                "Office file."
            )

    # -----------------------------------------------------
    # Save binary file
    # -----------------------------------------------------

    with open(
        output_path,
        "wb"
    ) as file_handle:

        file_handle.write(
            file_bytes
        )

    # -----------------------------------------------------
    # Verify saved file
    # -----------------------------------------------------

    if not output_path.exists():

        raise RuntimeError(
            f"Downloaded file was not created: "
            f"{output_path}"
        )

    final_size = (
        output_path.stat().st_size
    )

    print(
        f"Saved size: "
        f"{final_size:,} bytes"
    )

    if final_size != downloaded_size:

        raise RuntimeError(
            "Saved file size does not match "
            "downloaded file size."
        )

    print(
        f"Saved successfully:"
    )

    print(
        output_path
    )

    return output_path


# =========================================================
# DOWNLOAD CSR INPUT FILES
# =========================================================

def download_csr_files():

    print(
        "=" * 60
    )

    print(
        "GOOGLE DRIVE CSR INPUT DOWNLOADER"
    )

    print(
        "=" * 60
    )

    # -----------------------------------------------------
    # Connect to Google Drive
    # -----------------------------------------------------

    print(
        "\nConnecting to Google Drive..."
    )

    service = get_drive_service()

    print(
        "Google Drive connection successful."
    )

    # -----------------------------------------------------
    # Find CSR folder
    # -----------------------------------------------------

    folder = find_folder(
        service,
        "CSR_Report_Automation"
    )

    folder_id = folder["id"]

    print(
        f"\nFolder found: "
        f"{folder['name']}"
    )

    print(
        f"Folder ID: "
        f"{folder_id}"
    )

    # =====================================================
    # EXCEL FILE
    # =====================================================

    excel_name = (
        "MONTH DATA FILE-2025-26 - complete.xlsx"
    )

    excel_file = find_file(
        service,
        folder_id,
        excel_name
    )

    print(
        f"\nExcel found: "
        f"{excel_file['name']}"
    )

    excel_path = download_file(
        service,
        excel_file["id"],
        excel_name,
        expected_mime_type=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )

    # =====================================================
    # MARATHI WORD FILE
    # =====================================================

    word_name = (
        "ACS-HY Marathi Report.docx"
    )

    word_file = find_file(
        service,
        folder_id,
        word_name
    )

    print(
        f"\nWord report found: "
        f"{word_file['name']}"
    )

    word_path = download_file(
        service,
        word_file["id"],
        word_name,
        expected_mime_type=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )

    # =====================================================
    # FINAL RESULT
    # =====================================================

    print(
        "\n" + "=" * 60
    )

    print(
        "GOOGLE DRIVE DOWNLOAD COMPLETED"
    )

    print(
        "=" * 60
    )

    print(
        "\nExcel:"
    )

    print(
        excel_path
    )

    print(
        "\nWord:"
    )

    print(
        word_path
    )

    print(
        "\nBoth files passed validation."
    )

    return {
        "excel": excel_path,
        "word": word_path
    }


# =========================================================
# TEST
# =========================================================

if __name__ == "__main__":

    download_csr_files()