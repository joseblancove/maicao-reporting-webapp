from __future__ import annotations

import subprocess
import tempfile
from pathlib import Path
from typing import List, Tuple


def _find_soffice() -> str | None:
    for cmd in ["libreoffice", "soffice"]:
        try:
            r = subprocess.run(["which", cmd], capture_output=True, text=True, timeout=8)
            if r.returncode == 0 and r.stdout.strip():
                return r.stdout.strip().splitlines()[0]
        except Exception:
            pass
    return None


def render_pptx_to_images(pptx_bytes: bytes, scale: float = 1.35, max_slides: int | None = None) -> Tuple[List[bytes], str | None]:
    """Render a PPTX to PNG images using LibreOffice + PyMuPDF.

    Returns (images, error). If rendering is unavailable, images will be [] and
    error will explain how to enable it. The PPT generation itself is independent
    from this preview step.
    """
    if not pptx_bytes:
        return [], "No se recibió PPT para previsualizar."

    soffice = _find_soffice()
    if not soffice:
        return [], "Preview exacto no disponible: falta LibreOffice en el entorno."

    try:
        import fitz  # PyMuPDF
    except Exception as exc:
        return [], f"Preview exacto no disponible: falta PyMuPDF ({exc})."

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        pptx_path = tmp_path / "report.pptx"
        pptx_path.write_bytes(pptx_bytes)

        cmd = [
            soffice,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(tmp_path),
            str(pptx_path),
        ]
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        except subprocess.TimeoutExpired:
            return [], "La conversión del PPT a preview tardó demasiado."

        if result.returncode != 0:
            err = (result.stderr or result.stdout or "").strip()
            return [], f"No pude convertir el PPT a preview PDF. {err[:500]}"

        pdf_candidates = list(tmp_path.glob("*.pdf"))
        if not pdf_candidates:
            return [], "LibreOffice no generó el PDF de preview."

        pdf_path = pdf_candidates[0]
        try:
            doc = fitz.open(str(pdf_path))
            images: List[bytes] = []
            page_count = len(doc) if max_slides is None else min(len(doc), max_slides)
            for i in range(page_count):
                page = doc.load_page(i)
                pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                images.append(pix.tobytes("png"))
            doc.close()
            return images, None
        except Exception as exc:
            return [], f"No pude renderizar el PDF del preview: {exc}"
