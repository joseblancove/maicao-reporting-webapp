"""
Maicao Report Automation Web App v05

Streamlit app to generate the Maicao monthly PPT report from either:
1) an uploaded Excel model, or
2) a Google Sheet connected through a Google Cloud service account.

Run locally:
    streamlit run app.py

Deploy online:
    Upload this folder to GitHub and deploy with Streamlit Community Cloud or an internal Python host.
"""
from __future__ import annotations

import io
import json
import re
import tempfile
from pathlib import Path
from typing import Dict, Any, Optional

import streamlit as st
from openpyxl import Workbook

from generate_report_from_template import build_context, update_ppt, write_validation_report

ROOT = Path(__file__).resolve().parent
DEFAULT_EXCEL = ROOT / "Maicao_Reporte_Automation_Model_v04.xlsx"
DEFAULT_TEMPLATE = ROOT / "template" / "Maicao_Template_Visual_v02.pptx"
DEFAULT_OUTPUT_NAME = "Maicao_Reporte_Auto_Web_v05.pptx"

REQUIRED_SHEETS = [
    "00_Control",
    "02_Audience",
    "03_Paid_Media",
    "04_Squad",
    "05_MMPP",
    "07_Insights",
    "08_Recomendaciones",
    "09_Validaciones",
    "13_Platform_KPIs",
]


def extract_sheet_id(url_or_id: str) -> str:
    """Accept a raw Google Sheet ID or a full URL and return the spreadsheet ID."""
    s = (url_or_id or "").strip()
    if not s:
        raise ValueError("Debes pegar el ID o URL del Google Sheet.")
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", s)
    if m:
        return m.group(1)
    # Raw ID fallback.
    if re.fullmatch(r"[a-zA-Z0-9-_]{20,}", s):
        return s
    raise ValueError("No pude detectar el ID del Google Sheet. Pega la URL completa o el ID.")


def get_service_account_info(uploaded_file: Optional[Any]) -> Dict[str, Any]:
    """Load service account JSON from Streamlit secrets or uploaded JSON file."""
    # Streamlit secrets for hosted deployment.
    try:
        if "gcp_service_account" in st.secrets:
            return dict(st.secrets["gcp_service_account"])
    except Exception:
        pass

    if uploaded_file is None:
        raise ValueError(
            "Falta credencial Google. Sube el JSON de service account o configúralo en Streamlit secrets."
        )
    try:
        return json.loads(uploaded_file.getvalue().decode("utf-8"))
    except Exception as exc:
        raise ValueError(f"No pude leer el JSON de service account: {exc}")


def google_sheet_to_xlsx(spreadsheet_url_or_id: str, service_account_info: Dict[str, Any], output_path: Path) -> Path:
    """Download all Google Sheet tabs into an .xlsx file, preserving worksheet names."""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError as exc:
        raise RuntimeError(
            "Faltan dependencias Google. Instala requirements.txt: pip install -r requirements.txt"
        ) from exc

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    client = gspread.authorize(creds)
    spreadsheet_id = extract_sheet_id(spreadsheet_url_or_id)
    sh = client.open_by_key(spreadsheet_id)

    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    for worksheet in sh.worksheets():
        title = worksheet.title[:31]
        ws = wb.create_sheet(title=title)
        # UNFORMATTED_VALUE returns numbers as numbers when possible; fallback keeps display values.
        try:
            values = worksheet.get_all_values(value_render_option="UNFORMATTED_VALUE")
        except TypeError:
            values = worksheet.get_all_values()
        for row_idx, row in enumerate(values, start=1):
            for col_idx, value in enumerate(row, start=1):
                ws.cell(row=row_idx, column=col_idx, value=value)

    wb.save(output_path)
    return output_path


def save_uploaded_excel(uploaded_file: Any, output_path: Path) -> Path:
    output_path.write_bytes(uploaded_file.getvalue())
    return output_path


def generate_ppt(input_xlsx: Path, template_pptx: Path, strict: bool = False) -> tuple[bytes, str, list[str]]:
    ctx = build_context(str(input_xlsx))
    warnings = ctx.get("_WARNINGS", [])
    if strict and warnings:
        validation_stream = io.StringIO()
        validation_stream.write("VALIDACION: REVISAR\n\n")
        for warning in warnings:
            validation_stream.write(f"- {warning}\n")
        return b"", validation_stream.getvalue(), warnings

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        output_pptx = tmpdir_path / DEFAULT_OUTPUT_NAME
        validation_txt = tmpdir_path / "validation_report_v05.txt"
        update_ppt(str(template_pptx), str(output_pptx), ctx)
        write_validation_report(validation_txt, ctx)
        return output_pptx.read_bytes(), validation_txt.read_text(encoding="utf-8"), warnings


def sheet_name_check(xlsx_path: Path) -> list[str]:
    try:
        from openpyxl import load_workbook
        wb = load_workbook(xlsx_path, read_only=True, data_only=True)
        existing = set(wb.sheetnames)
        return [s for s in REQUIRED_SHEETS if s not in existing]
    except Exception:
        return []


st.set_page_config(page_title="Maicao Report Generator", page_icon="📊", layout="wide")

st.title("Maicao Report Generator · v05")
st.caption("Google Sheets / Excel → PowerPoint editable")

with st.sidebar:
    st.header("Fuente de datos")
    source = st.radio(
        "Selecciona cómo cargar la data",
        ["Google Sheets", "Subir Excel"],
        index=0,
    )
    strict_mode = st.checkbox("Modo estricto: detener si faltan datos", value=False)
    st.divider()
    st.markdown("**Plantilla PPT**")
    st.write("Usando plantilla incluida:")
    st.code(str(DEFAULT_TEMPLATE.name))

st.markdown(
    """
### Flujo
1. Actualiza el Google Sheet o Excel con el modelo v04/v05.  
2. Carga la fuente de datos aquí.  
3. Genera y descarga el PowerPoint editable.  
"""
)

input_xlsx_path: Optional[Path] = None
service_email_hint = ""

with tempfile.TemporaryDirectory() as tempdir:
    tempdir_path = Path(tempdir)
    if source == "Google Sheets":
        st.subheader("Conectar Google Sheets")
        st.info(
            "Para que funcione online, comparte el Google Sheet con el email del service account. "
            "También puedes subir el JSON aquí para la demo local."
        )
        sheet_url = st.text_input("URL o ID del Google Sheet")
        sa_file = st.file_uploader("Service account JSON", type=["json"], help="No se guarda; solo se usa durante esta sesión.")

        if sa_file is not None:
            try:
                info = json.loads(sa_file.getvalue().decode("utf-8"))
                service_email_hint = info.get("client_email", "")
                if service_email_hint:
                    st.success(f"Comparte el Sheet con: {service_email_hint}")
            except Exception:
                pass

        if st.button("Generar desde Google Sheets", type="primary"):
            try:
                sa_info = get_service_account_info(sa_file)
                input_xlsx_path = google_sheet_to_xlsx(sheet_url, sa_info, tempdir_path / "google_sheet_input.xlsx")
                missing_sheets = sheet_name_check(input_xlsx_path)
                if missing_sheets:
                    st.warning("Faltan hojas esperadas: " + ", ".join(missing_sheets))
                ppt_bytes, validation_text, warnings = generate_ppt(input_xlsx_path, DEFAULT_TEMPLATE, strict=strict_mode)
                if warnings:
                    st.warning("Validación: revisar alertas antes de enviar al cliente.")
                else:
                    st.success("Validación OK. Reporte generado.")
                st.text_area("Reporte de validación", validation_text, height=220)
                if ppt_bytes:
                    st.download_button(
                        "Descargar PPT editable",
                        data=ppt_bytes,
                        file_name=DEFAULT_OUTPUT_NAME,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )
            except Exception as exc:
                st.error(str(exc))

    else:
        st.subheader("Subir Excel")
        uploaded_excel = st.file_uploader("Excel modelo Maicao", type=["xlsx"])
        st.caption("También puedes usar el Excel incluido si estás corriendo la app localmente.")
        use_default = st.checkbox("Usar Excel incluido de ejemplo", value=False)

        if st.button("Generar desde Excel", type="primary"):
            try:
                if use_default:
                    input_xlsx_path = DEFAULT_EXCEL
                elif uploaded_excel is not None:
                    input_xlsx_path = save_uploaded_excel(uploaded_excel, tempdir_path / "uploaded_input.xlsx")
                else:
                    raise ValueError("Sube un Excel o marca 'Usar Excel incluido de ejemplo'.")
                missing_sheets = sheet_name_check(input_xlsx_path)
                if missing_sheets:
                    st.warning("Faltan hojas esperadas: " + ", ".join(missing_sheets))
                ppt_bytes, validation_text, warnings = generate_ppt(input_xlsx_path, DEFAULT_TEMPLATE, strict=strict_mode)
                if warnings:
                    st.warning("Validación: revisar alertas antes de enviar al cliente.")
                else:
                    st.success("Validación OK. Reporte generado.")
                st.text_area("Reporte de validación", validation_text, height=220)
                if ppt_bytes:
                    st.download_button(
                        "Descargar PPT editable",
                        data=ppt_bytes,
                        file_name=DEFAULT_OUTPUT_NAME,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )
            except Exception as exc:
                st.error(str(exc))

st.divider()
st.markdown(
    """
### Checklist para producción
- El Google Sheet debe tener las mismas pestañas y encabezados del Excel modelo.  
- La hoja `00_Control` define el mes activo.  
- La hoja `13_Platform_KPIs` alimenta KPIs por Instagram, Facebook y TikTok.  
- La hoja `02_Audience` alimenta seguidores y demografía.  
- La hoja `09_Validaciones` alimenta el overview ejecutivo.  
"""
)
