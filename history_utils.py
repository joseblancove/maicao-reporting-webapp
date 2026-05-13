from __future__ import annotations

import io
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional

import pandas as pd

HISTORY_SHEET = "99_Report_History"
HISTORY_HEADERS = [
    "report_id",
    "client",
    "month",
    "generated_at_utc",
    "version",
    "source",
    "filename",
    "status",
    "warnings_count",
    "ppt_drive_url",
    "ppt_file_id",
    "notes",
]


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")


def upload_report_to_drive(
    file_bytes: bytes,
    filename: str,
    service_account_info: Dict[str, Any],
    folder_id: str | None = None,
    mime_type: str = "application/vnd.openxmlformats-officedocument.presentationml.presentation",
) -> Dict[str, str]:
    """Upload a generated PPT to Google Drive.

    The Drive folder must be shared with the service account. Returns Google file
    metadata with id and webViewLink when successful.
    """
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseUpload

    scopes = ["https://www.googleapis.com/auth/drive.file"]
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    service = build("drive", "v3", credentials=creds, cache_discovery=False)

    metadata: Dict[str, Any] = {"name": filename, "mimeType": mime_type}
    if folder_id:
        metadata["parents"] = [folder_id]
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=False)
    created = service.files().create(body=metadata, media_body=media, fields="id,webViewLink,webContentLink").execute()
    return {"id": created.get("id", ""), "webViewLink": created.get("webViewLink", ""), "webContentLink": created.get("webContentLink", "")}


def _get_history_ws(spreadsheet_id: str, service_account_info: Dict[str, Any]):
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    client = gspread.authorize(creds)
    sh = client.open_by_key(spreadsheet_id)
    try:
        ws = sh.worksheet(HISTORY_SHEET)
    except Exception:
        ws = sh.add_worksheet(title=HISTORY_SHEET, rows=200, cols=len(HISTORY_HEADERS) + 2)
        ws.append_row(HISTORY_HEADERS, value_input_option="USER_ENTERED")
        return ws

    values = ws.get_all_values()
    if not values:
        ws.append_row(HISTORY_HEADERS, value_input_option="USER_ENTERED")
    elif values[0] != HISTORY_HEADERS:
        # Preserve existing content but ensure the first row has the expected headers.
        ws.update("A1", [HISTORY_HEADERS])
    return ws


def append_history_row(spreadsheet_id: str, service_account_info: Dict[str, Any], row: Dict[str, Any]) -> None:
    ws = _get_history_ws(spreadsheet_id, service_account_info)
    ws.append_row([row.get(h, "") for h in HISTORY_HEADERS], value_input_option="USER_ENTERED")


def get_history_df(spreadsheet_id: str, service_account_info: Dict[str, Any]) -> pd.DataFrame:
    ws = _get_history_ws(spreadsheet_id, service_account_info)
    values = ws.get_all_records()
    if not values:
        return pd.DataFrame(columns=HISTORY_HEADERS)
    df = pd.DataFrame(values)
    for h in HISTORY_HEADERS:
        if h not in df.columns:
            df[h] = ""
    return df[HISTORY_HEADERS]
