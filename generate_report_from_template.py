#!/usr/bin/env python3
"""
Generador Maicao Visual v12

Mejoras v12:
- Validacion automatica antes de generar: avisa datos faltantes por hoja/plataforma/mes.
- Fallback visual: si falta un dato, la PPT muestra "Dato pendiente".
- Ajuste de tarjetas KPI para evitar que "vs mes anterior" se monte sobre el numero.
- Normalizacion de mes: acepta texto tipo "Marzo 2026" o fechas Excel tipo "mar-26".
- KPIs por plataforma desde Excel: hoja 13_Platform_KPIs.
- Barras dinamicas como shapes: recalcula altos/posiciones desde la data.
- Textos cualitativos dinamicos desde Google Sheets/Excel: hojas 15_Qualitative_Texts y 16_Action_Plan.

Uso Terminal:
  python3 generate_report_from_template.py

Salida default:
  output/Maicao_Reporte_Auto_Template_v13.pptx
"""
import argparse, json, re, sys
from pathlib import Path
from datetime import datetime, date
from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Pt, Inches
from media_utils import enrich_media_context, append_media_slides

ROOT = Path(__file__).resolve().parent
DEFAULTS_PATH = ROOT / "config" / "defaults.json"
MISSING = "Dato pendiente"
PLATFORMS = ["Instagram", "Facebook", "TikTok"]
MONTHS_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}
MONTHS_LOOKUP = {v.lower(): k for k, v in MONTHS_ES.items()}
MONTHS_LOOKUP.update({
    "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dic": 12,
    "jan": 1, "apr": 4, "aug": 8, "dec": 12,
})


def is_missing(v):
    return v is None or v == "" or str(v).strip().lower() in {"none", "nan", "na", "n/a"}


def normalize_month(v):
    """Return canonical Spanish month label, e.g. Marzo 2026."""
    if isinstance(v, (datetime, date)):
        return f"{MONTHS_ES[v.month]} {v.year}"
    if is_missing(v):
        return ""
    s = str(v).strip()
    # Excel sometimes stores strings like 2026-03-01 00:00:00
    for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"]:
        try:
            d = datetime.strptime(s, fmt)
            return f"{MONTHS_ES[d.month]} {d.year}"
        except Exception:
            pass
    # mar-26 / Mar-26 / marzo 2026 / MARZO 2026
    clean = s.replace(".", "").replace("_", " ").replace("-", " ")
    parts = [p for p in clean.split() if p]
    if len(parts) >= 2:
        mon = parts[0].lower()
        yr = parts[1]
        m = MONTHS_LOOKUP.get(mon)
        if m:
            try:
                y = int(yr)
                if y < 100:
                    y += 2000
                return f"{MONTHS_ES[m]} {y}"
            except Exception:
                pass
    # preserve user text with first letter capitalized when possible
    return s


def month_short(v):
    nm = normalize_month(v)
    return nm.split()[0].lower() if nm else "mes anterior"


def fmt_int(n, missing=MISSING):
    if is_missing(n):
        return missing
    try:
        return f"{int(round(float(n))):,}".replace(",", ".")
    except Exception:
        return str(n)


def fmt_money(n, missing=MISSING):
    if is_missing(n):
        return missing
    return "$" + fmt_int(n, missing="")


def fmt_pct(n, dec=1, comma=True, missing=MISSING):
    if is_missing(n):
        return missing
    try:
        s = f"{float(n):.{dec}f}%"
        return s.replace(".", ",") if comma else s
    except Exception:
        return str(n)


def fmt_m(n, dec=1, missing=MISSING):
    if is_missing(n):
        return missing
    try:
        return f"{float(n)/1_000_000:.{dec}f}M".replace(".", ",")
    except Exception:
        return str(n)


def fmt_k(n, dec=1, missing=MISSING):
    if is_missing(n):
        return missing
    try:
        return f"{float(n)/1_000:.{dec}f}K".replace(".", ",")
    except Exception:
        return str(n)


def to_float(v, default=None):
    """Parse numeric values from Excel or formatted strings like 1,7M / 88K / 1,4% / $850.000."""
    if is_missing(v):
        return default
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s:
        return default
    mult = 1.0
    su = s.upper()
    if "M" in su:
        mult = 1_000_000.0
    elif "K" in su:
        mult = 1_000.0
    # Keep digits, comma, dot and minus. Determine decimal separator.
    clean = re.sub(r"[^0-9,\.\-]", "", s)
    if not clean:
        return default
    if "," in clean and "." in clean:
        # Assume dots are thousands and comma is decimal: 1.234,5
        clean = clean.replace(".", "").replace(",", ".")
    elif "," in clean:
        clean = clean.replace(",", ".")
    elif "." in clean:
        # If there are multiple dots, treat them as thousands separators.
        if clean.count(".") > 1:
            clean = clean.replace(".", "")
        # If a single dot and three digits after it, likely thousands, e.g. 13.463
        else:
            left, right = clean.split(".")
            if len(right) == 3 and len(left) >= 1:
                clean = left + right
    try:
        return float(clean) * mult
    except Exception:
        return default


def diff_pct(new, old):
    if is_missing(old) or old == 0 or is_missing(new):
        return None
    try:
        return (float(new) - float(old)) / float(old) * 100
    except Exception:
        return None


def var_label(new, old, prev_label, dec=1, approx=False):
    d = diff_pct(new, old)
    if d is None:
        return MISSING
    label = fmt_pct(d, dec=dec).replace("-", "−")
    if approx:
        return f"{label} aprox vs {prev_label}"
    return f"{label} vs {prev_label}"


def yes(v):
    return str(v).strip().lower() in ["si", "sí", "yes", "true", "1", "x"]


def _normalize_row_aliases(row):
    """Add backwards-compatible aliases so the app supports both v08/v11 Spanish headers and v12 snake_case headers."""
    aliases = {
        'month': 'Mes', 'platform': 'Plataforma', 'views': 'Views', 'reach': 'Alcance',
        'interactions': 'Interacciones', 'contents': 'Contenidos', 'budget': 'Presupuesto',
        'views_prev': 'Views mes anterior', 'reach_prev': 'Alcance mes anterior',
        'interactions_prev': 'Interacciones mes anterior', 'er': 'ER %', 'er_prev': 'ER % mes anterior',
        'contents_prev': 'Contenidos mes anterior', 'budget_prev': 'Presupuesto mes anterior',
        'followers': 'Seguidores', 'followers_prev': 'Seguidores mes anterior', 'growth': 'Crecimiento',
        'women_pct': 'Mujeres %', 'men_pct': 'Hombres %', 'age_18_24': '18-24 %',
        'age_25_34': '25-34 %', 'age_35_44': '35-44 %', 'age_45_54': '45-54 %', 'age_55_plus': '55+ %',
        'amount': 'Monto', 'amount_prev': 'Monto mes anterior', 'include_overview': 'Incluir Overview',
        'member': 'Miembro', 'display_name': 'Talento', 'active': 'Incluir PPT',
        'display_name': 'Talento', 'brand': 'Marca', 'summary': 'Comentario general',
        'placeholder': 'Placeholder', 'text': 'Texto', 'use_in_ppt': 'Usar en PPT', 'active': 'Usar en PPT',
        'order': 'Orden', 'pillar': 'Pilar', 'action': 'Accion', 'goal': 'Meta',
        'kpi_focus': 'KPIs de control',
    }
    for src, dst in aliases.items():
        if dst not in row and src in row:
            row[dst] = row.get(src)
    # v12 squad member aliases to legacy full names used by original generator.
    if '04_Squad' not in row.get('_sheet_name', ''):
        pass
    if row.get('member') and row.get('Talento'):
        m = str(row.get('member')).strip().lower()
        full = {'skar': 'Skarleth Labra', 'busquilla': 'Busquilla', 'cami': 'Camila Andrade', 'disley': 'Disley Ramos'}.get(m)
        if full:
            row['Talento'] = full
    # v12 paid media: platform rows should also be included in platform budget unless it is Squad.
    if 'Incluir Plataforma' not in row:
        row['Incluir Plataforma'] = row.get('Plataforma') not in [None, '', 'Squad']
    # v12 raw content aliases for fallback logic.
    if 'include_platform' in row and 'Incluir Plataforma' not in row:
        row['Incluir Plataforma'] = row.get('include_platform')
    if 'include_overview' in row and 'Incluir Overview' not in row:
        row['Incluir Overview'] = row.get('include_overview')
    if 'content_type' in row and 'Formato' not in row:
        row['Formato'] = row.get('content_type')
    if 'views' in row and 'Visualizaciones' not in row:
        row['Visualizaciones'] = row.get('views')
    if 'interactions' in row and 'Interacciones totales' not in row:
        row['Interacciones totales'] = row.get('interactions')
    return row


def _detect_header_row(ws, preferred=4):
    candidates = [preferred, 1, 2, 3, 5]
    seen = set()
    for r in candidates:
        if r in seen or r < 1 or r > ws.max_row:
            continue
        seen.add(r)
        vals = [c.value for c in ws[r]]
        labels = {str(v).strip().lower() for v in vals if v is not None}
        # v12 snake_case headers or v08 Spanish headers.
        if {'month', 'platform'} & labels or {'mes', 'plataforma'} <= labels or {'campo', 'valor'} <= labels:
            return r
    return preferred if preferred <= ws.max_row else 1


def as_rows(ws, header_row=4):
    header_row = _detect_header_row(ws, header_row)
    headers = [c.value for c in ws[header_row]]
    rows = []
    for r in ws.iter_rows(min_row=header_row+1, values_only=True):
        if not any(v is not None for v in r):
            continue
        row = {str(headers[i]).strip(): r[i] if i < len(r) else None for i in range(len(headers)) if headers[i] is not None}
        row['_sheet_name'] = ws.title
        rows.append(_normalize_row_aliases(row))
    return rows


def control_dict(wb):
    ws = wb['00_Control']
    d = {}
    # v12: row 1 headers Campo / Valor. Legacy: values start row 5.
    rows = as_rows(ws, header_row=1)
    if rows and ('Campo' in rows[0] or 'field' in rows[0]):
        for r in rows:
            key = r.get('Campo') or r.get('field')
            val = r.get('Valor') if 'Valor' in r else r.get('value')
            if key:
                d[str(key)] = val
    else:
        for row in ws.iter_rows(min_row=5, max_col=2, values_only=True):
            if row[0]:
                d[str(row[0])] = row[1]
    return d


def validation_value(wb, metric, prefer_expected=True):
    ws = wb['09_Validaciones']
    # v12 layout: metric | value | previous_value | variation
    metric_map = {'Views':'views_total', 'Alcance':'reach_total', 'Interacciones':'interactions_total', 'Contenidos':'contents_total', 'ER':'er_total', 'Inversion':'investment_total'}
    rows = as_rows(ws, header_row=1)
    if rows and ('metric' in rows[0] or 'Metric' in rows[0]):
        want = metric_map.get(metric, metric)
        for r in rows:
            if str(r.get('metric') or '').strip() == want:
                return r.get('value')
    # legacy layout.
    for row in ws.iter_rows(min_row=5, values_only=True):
        if len(row) > 3 and row[1] == metric:
            return row[3] if prefer_expected and row[3] is not None else row[2]
    return None


def get_audience(wb, month, warnings):
    out = {}
    if '02_Audience' not in wb.sheetnames:
        warnings.append("Falta hoja 02_Audience.")
        return out
    for r in as_rows(wb['02_Audience']):
        p = r.get('Plataforma')
        if not p:
            continue
        row_month = normalize_month(r.get('Mes'))
        if row_month and row_month != month:
            continue
        out[str(p).strip()] = r
    for p in PLATFORMS:
        if p not in out:
            warnings.append(f"Falta fila en 02_Audience para {p} / {month}.")
    return out


def get_paid(wb, month):
    paid = {}
    if '03_Paid_Media' not in wb.sheetnames:
        return paid
    for r in as_rows(wb['03_Paid_Media']):
        row_month = normalize_month(r.get('Mes'))
        if row_month and row_month != month:
            continue
        p = r.get('Plataforma')
        if not p:
            continue
        paid.setdefault(str(p).strip(), 0)
        if yes(r.get('Incluir Plataforma')):
            paid[str(p).strip()] += r.get('Monto') or 0
    return paid


def platform_kpis_from_excel(wb, month, warnings):
    platforms = {}
    if '13_Platform_KPIs' in wb.sheetnames:
        for r in as_rows(wb['13_Platform_KPIs']):
            if not r.get('Plataforma'):
                continue
            row_month = normalize_month(r.get('Mes'))
            if row_month and row_month != month:
                continue
            p = str(r['Plataforma']).strip()
            platforms[p] = {
                'views': r.get('Views'),
                'views_prev': r.get('Views mes anterior'),
                'reach': r.get('Alcance'),
                'reach_prev': r.get('Alcance mes anterior'),
                'interactions': r.get('Interacciones'),
                'interactions_prev': r.get('Interacciones mes anterior'),
                'er': r.get('ER %'),
                'er_prev': r.get('ER % mes anterior'),
                'contents': r.get('Contenidos'),
                'contents_prev': r.get('Contenidos mes anterior'),
                'stories': r.get('Stories'),
                'reels': r.get('Reels'),
                'carousels': r.get('Carruseles'),
                'posts': r.get('Post'),
                'budget': r.get('Presupuesto'),
                'budget_prev': r.get('Presupuesto mes anterior'),
                'top1_name': r.get('Top 1 nombre'),
                'top1_value': r.get('Top 1 valor'),
                'top2_name': r.get('Top 2 nombre'),
                'top2_value': r.get('Top 2 valor'),
                'top3_name': r.get('Top 3 nombre'),
                'top3_value': r.get('Top 3 valor'),
                'insight': r.get('Insight corto'),
                'opportunity': r.get('Oportunidad / Nota'),
            }
        required = ['views', 'reach', 'interactions', 'er', 'contents']
        for p in PLATFORMS:
            if p not in platforms:
                warnings.append(f"Falta fila en 13_Platform_KPIs para {p} / {month}.")
            else:
                for f in required:
                    if is_missing(platforms[p].get(f)):
                        warnings.append(f"Falta {f} en 13_Platform_KPIs para {p} / {month}.")
        if platforms:
            return platforms

    # Fallback from 01_Content_Raw if the official sheet is absent or empty.
    content = as_rows(wb['01_Content_Raw']) if '01_Content_Raw' in wb.sheetnames else []
    fmt_map = {'STORY': 'stories', 'REEL': 'reels', 'VIDEO': 'reels', 'CARRUSEL': 'carousels', 'POST': 'posts'}
    for r in content:
        if normalize_month(r.get('Mes')) != month or not yes(r.get('Incluir Plataforma')):
            continue
        p = r.get('Plataforma')
        if not p:
            continue
        rec = platforms.setdefault(str(p).strip(), {'views':0,'reach':0,'interactions':0,'contents':0,'stories':0,'reels':0,'carousels':0,'posts':0})
        rec['views'] += r.get('Views') or 0
        rec['reach'] += r.get('Alcance') or 0
        rec['interactions'] += r.get('Interacciones totales') or 0
        rec['contents'] += 1
        f = str(r.get('Formato') or '').upper()
        if f in fmt_map:
            rec[fmt_map[f]] += 1
    for p, rec in platforms.items():
        rec['er'] = (rec['interactions'] / rec['reach'] * 100) if rec.get('reach') else None
    return platforms


def first_insight(wb, section, type_name=None, month=None):
    if '07_Insights' not in wb.sheetnames:
        return ''
    for r in as_rows(wb['07_Insights']):
        if month and normalize_month(r.get('Mes')) not in {month, ''}:
            continue
        if str(r.get('Sección')).strip().lower() == section.lower() and yes(r.get('Usar en PPT')):
            if type_name is None or str(r.get('Tipo')).strip().lower() == type_name.lower():
                parts = [r.get('Hallazgo'), r.get('Evidencia'), r.get('Recomendación')]
                return ' '.join(str(x) for x in parts if x)
    return ''


def filter_month_rows(wb, sheet_name, month, include_col=None):
    if sheet_name not in wb.sheetnames:
        return []
    rows = []
    for r in as_rows(wb[sheet_name]):
        row_month = normalize_month(r.get('Mes'))
        if row_month and row_month != month:
            continue
        if include_col and not yes(r.get(include_col)):
            continue
        rows.append(r)
    return rows


def get_qualitative_texts(wb, month, warnings):
    """Return a dict of placeholder -> text from 15_Qualitative_Texts.

    Expected headers at row 4:
    Mes | Slide | Seccion | Placeholder | Texto | Usar en PPT | Comentario
    """
    out = {}
    sheet = '15_Qualitative_Texts'
    if sheet not in wb.sheetnames:
        warnings.append('Falta hoja 15_Qualitative_Texts para textos cualitativos.')
        return out
    for r in as_rows(wb[sheet]):
        row_month = normalize_month(r.get('Mes'))
        if row_month and row_month != month:
            continue
        if not yes(r.get('Usar en PPT')):
            continue
        key = str(r.get('Placeholder') or '').strip()
        text = r.get('Texto')
        if key:
            out[key] = str(text) if not is_missing(text) else MISSING
    return out


def get_action_plan_texts(wb, month, warnings):
    """Return placeholders for slide 12 from 16_Action_Plan."""
    out = {}
    sheet = '16_Action_Plan'
    if sheet not in wb.sheetnames:
        warnings.append('Falta hoja 16_Action_Plan para plan de accion.')
        return out
    rows = []
    for r in as_rows(wb[sheet]):
        row_month = normalize_month(r.get('Mes'))
        if row_month and row_month != month:
            continue
        if not yes(r.get('Usar en PPT')):
            continue
        rows.append(r)
    rows.sort(key=lambda r: int(to_float(r.get('Orden'), 999) or 999))
    for idx, r in enumerate(rows[:4], start=1):
        out[f'PLAN_{idx}_PILLAR'] = str(r.get('Pilar') or MISSING)
        out[f'PLAN_{idx}_ACTION'] = str(r.get('Accion') or r.get('Acción') or MISSING)
        meta = r.get('Meta')
        out[f'PLAN_{idx}_META'] = str(meta) if not is_missing(meta) else MISSING
    # Optional KPI control line. If any row contains KPIs de control, prefer it.
    kpi_line = None
    for r in rows:
        if not is_missing(r.get('KPIs de control')):
            kpi_line = r.get('KPIs de control')
            break
    if kpi_line:
        text = str(kpi_line)
        out['PLAN_KPIS_CONTROL'] = text if text.lower().startswith('kpis de control') else f'KPIs de control: {text}'
    return out


def build_context(xlsx_path):
    defaults = json.loads(DEFAULTS_PATH.read_text(encoding='utf-8'))
    wb = load_workbook(xlsx_path, data_only=True)
    warnings = []
    ctrl = control_dict(wb)
    month = normalize_month(ctrl.get('Mes actual') or ctrl.get('active_month') or 'Marzo 2026')
    prev_month = normalize_month(ctrl.get('Mes anterior') or ctrl.get('previous_month') or 'Febrero 2026')
    prev_label = month_short(prev_month)
    month_upper = month.split()[0].upper() if month else MISSING
    prev = defaults['previous_month']

    overview = {
        'views': validation_value(wb, 'Views') or prev.get('views'),
        'reach': validation_value(wb, 'Alcance'),
        'interactions': validation_value(wb, 'Interacciones'),
        'contents': validation_value(wb, 'Contenidos'),
        'investment': sum((r.get('Monto') or 0) for r in filter_month_rows(wb, '03_Paid_Media', month) if yes(r.get('Incluir Overview'))),
        'er': None,
    }
    overview['er'] = (float(overview['interactions']) / float(overview['reach']) * 100) if overview.get('reach') else None
    for metric, val in [('Views', overview['views']), ('Alcance', overview['reach']), ('Interacciones', overview['interactions']), ('Contenidos', overview['contents'])]:
        if is_missing(val):
            warnings.append(f"Falta valor de overview para {metric} en 09_Validaciones.")

    audience = get_audience(wb, month, warnings)
    platform_kpis = platform_kpis_from_excel(wb, month, warnings)
    paid = get_paid(wb, month)

    # Merge audience and paid into platform records.
    for p, a in audience.items():
        platform_kpis.setdefault(p, {})
        platform_kpis[p].update({
            'followers': a.get('Seguidores'),
            'followers_prev': a.get('Seguidores mes anterior'),
            'followers_growth': a.get('Crecimiento'),
            'women_pct': a.get('Mujeres %'),
            'men_pct': a.get('Hombres %'),
            'age_18_24': a.get('18-24 %'),
            'age_25_34': a.get('25-34 %'),
            'age_35_44': a.get('35-44 %'),
            'age_45_54': a.get('45-54 %'),
            'age_55': a.get('55+ %'),
        })
    for p in PLATFORMS:
        rec = platform_kpis.get(p, {})
        if is_missing(rec.get('followers')):
            warnings.append(f"Falta Seguidores en 02_Audience para {p} / {month}.")
    for p, amount in paid.items():
        platform_kpis.setdefault(p, {})
        if is_missing(platform_kpis[p].get('budget')):
            platform_kpis[p]['budget'] = amount

    squad_rows = filter_month_rows(wb, '04_Squad', month, include_col='Incluir PPT')
    squad_tot = {
        'views': sum(r.get('Views') or 0 for r in squad_rows),
        'reach': sum(r.get('Alcance') or 0 for r in squad_rows),
        'interactions': sum(r.get('Interacciones') or 0 for r in squad_rows),
        'budget': sum(r.get('Monto') or 0 for r in squad_rows),
    }
    squad_tot['er'] = (squad_tot['interactions'] / squad_tot['reach'] * 100) if squad_tot['reach'] else None
    squad = {r['Talento']: r for r in squad_rows if r.get('Talento')}

    mmpp_rows = filter_month_rows(wb, '05_MMPP', month, include_col='Usar en PPT')
    mmpp = mmpp_rows[0] if mmpp_rows else {}
    if not mmpp:
        warnings.append(f"Falta fila activa en 05_MMPP para {month}.")

    ctx = {
        'CLIENTE': ctrl.get('Cliente') or ctrl.get('client_name') or 'Maicao',
        'MES': month or MISSING,
        'MES_UPPER': month_upper,
        'MES_PREV': prev_month or MISSING,
        'MES_PREV_LABEL': prev_label,
        'TITLE_META': f"{ctrl.get('Cliente') or ctrl.get('client_name') or 'Maicao'} · {month}",
        'OVERVIEW_VIEWS': fmt_int(overview['views']),
        'OVERVIEW_VIEWS_M': fmt_m(overview['views']),
        'OVERVIEW_REACH': fmt_int(overview['reach']),
        'OVERVIEW_REACH_M': fmt_m(overview['reach']),
        'OVERVIEW_INTERACTIONS': fmt_int(overview['interactions']),
        'OVERVIEW_INTERACTIONS_K': fmt_k(overview['interactions']),
        'OVERVIEW_ER': fmt_pct(overview['er']),
        'OVERVIEW_INVESTMENT': fmt_money(overview['investment']),
        'OVERVIEW_CONTENTS': fmt_int(overview['contents']),
        'OVERVIEW_VIEWS_VAR': var_label(overview['views'], prev.get('views'), prev_label),
        'OVERVIEW_REACH_VAR': var_label(overview['reach'], prev.get('reach'), prev_label),
        'OVERVIEW_INTERACTIONS_VAR': var_label(overview['interactions'], prev.get('interactions'), prev_label),
        'OVERVIEW_ER_VAR': var_label(overview['er'], prev.get('er'), prev_label),
        'OVERVIEW_INVESTMENT_VAR': var_label(overview['investment'], prev.get('investment'), prev_label),
        'PREV_VIEWS_M': fmt_m(prev['views']),
        'PREV_INTERACTIONS_K': fmt_k(prev['interactions']),
        'PREV_ER_DOT': fmt_pct(prev['er'], comma=False),
        'CURR_ER_DOT': fmt_pct(overview['er'], comma=False),
        'PREV_INVESTMENT': fmt_int(prev['investment']),
        'CURR_INVESTMENT': fmt_int(overview['investment']),
        'PREV_CONTENTS': fmt_int(prev['contents']),
    }

    for name, prefix in [('Instagram', 'IG'), ('Facebook', 'FB'), ('TikTok', 'TT')]:
        p = platform_kpis.get(name, {})
        followers_growth = p.get('followers_growth')
        if is_missing(followers_growth) and not is_missing(p.get('followers')) and not is_missing(p.get('followers_prev')):
            try:
                followers_growth = (p.get('followers') or 0) - (p.get('followers_prev') or 0)
            except Exception:
                followers_growth = None
        if not is_missing(followers_growth):
            fg = float(followers_growth)
            growth_label = f"+{fmt_int(fg, missing='')} vs mes anterior" if fg >= 0 else f"{fmt_int(fg, missing='')} vs mes anterior"
        else:
            growth_label = MISSING
        ctx.update({
            f'{prefix}_VIEWS': fmt_int(p.get('views')),
            f'{prefix}_VIEWS_M': fmt_m(p.get('views')),
            f'{prefix}_REACH': fmt_int(p.get('reach')),
            f'{prefix}_REACH_M': fmt_m(p.get('reach')),
            f'{prefix}_INTERACTIONS': fmt_int(p.get('interactions')),
            f'{prefix}_ER': fmt_pct(p.get('er')),
            f'{prefix}_FOLLOWERS': fmt_int(p.get('followers')),
            f'{prefix}_FOLLOWERS_GROWTH': growth_label,
            f'{prefix}_CONTENTS': fmt_int(p.get('contents')),
            f'{prefix}_BUDGET': fmt_money(p.get('budget')),
            f'{prefix}_VIEWS_VAR': var_label(p.get('views'), p.get('views_prev'), prev_label, approx=(prefix=='FB')),
            f'{prefix}_REACH_VAR': var_label(p.get('reach'), p.get('reach_prev'), prev_label, approx=(prefix=='FB')),
            f'{prefix}_INTERACTIONS_VAR': var_label(p.get('interactions'), p.get('interactions_prev'), prev_label),
            f'{prefix}_ER_PREV': fmt_pct(p.get('er_prev')),
            f'{prefix}_BUDGET_PREV': fmt_money(p.get('budget_prev')),
            f'{prefix}_CONTENTS_VAR': (f"+{fmt_int((p.get('contents') or 0) - (p.get('contents_prev') or 0), missing='')} vs {prev_label}" if not is_missing(p.get('contents_prev')) and (p.get('contents') or 0) >= (p.get('contents_prev') or 0) else f"{fmt_int((p.get('contents') or 0) - (p.get('contents_prev') or 0), missing='')} vs {prev_label}" if not is_missing(p.get('contents_prev')) else MISSING),
            f'{prefix}_STORIES': fmt_int(p.get('stories')),
            f'{prefix}_REELS': fmt_int(p.get('reels')),
            f'{prefix}_CAROUSELS': fmt_int(p.get('carousels')),
            f'{prefix}_POSTS': fmt_int(p.get('posts')),
            f'{prefix}_AGE_18_24': fmt_pct(p.get('age_18_24'), 0),
            f'{prefix}_AGE_25_34': fmt_pct(p.get('age_25_34'), 0),
            f'{prefix}_AGE_35_44': fmt_pct(p.get('age_35_44'), 0),
            f'{prefix}_AGE_45_54': fmt_pct(p.get('age_45_54'), 0),
            f'{prefix}_AGE_55': fmt_pct(p.get('age_55'), 0),
            f'{prefix}_TOP1_NAME': p.get('top1_name') or MISSING,
            f'{prefix}_TOP1_VALUE': p.get('top1_value') or MISSING,
            f'{prefix}_TOP2_NAME': p.get('top2_name') or MISSING,
            f'{prefix}_TOP2_VALUE': p.get('top2_value') or MISSING,
            f'{prefix}_TOP3_NAME': p.get('top3_name') or MISSING,
            f'{prefix}_TOP3_VALUE': p.get('top3_value') or MISSING,
            f'{prefix}_INSIGHT': p.get('insight') or MISSING,
            f'{prefix}_OPPORTUNITY': p.get('opportunity') or MISSING,
        })

    ctx.update({
        'SQUAD_VIEWS': fmt_int(squad_tot['views'] if squad_rows else None),
        'SQUAD_VIEWS_M': fmt_m(squad_tot['views'] if squad_rows else None),
        'SQUAD_REACH': fmt_int(squad_tot['reach'] if squad_rows else None),
        'SQUAD_REACH_M': fmt_m(squad_tot['reach'] if squad_rows else None),
        'SQUAD_INTERACTIONS': fmt_int(squad_tot['interactions'] if squad_rows else None),
        'SQUAD_ER': fmt_pct(squad_tot['er']),
        'SQUAD_BUDGET': fmt_money(squad_tot['budget'] if squad_rows else None),
    })
    talent_map = {'Skarleth Labra': 'SKAR', 'Busquilla': 'BUSQUI', 'Camila Andrade': 'CAMI', 'Disley Ramos': 'DISLEY'}
    for t, prefix in talent_map.items():
        r = squad.get(t, {})
        val = r.get('Views')
        ctx[f'{prefix}_VIEWS_M'] = fmt_m(val) if val and val >= 1_000_000 else fmt_k(val, 0)
        ctx[f'{prefix}_ER'] = fmt_pct(r.get('ER %'))

    ctx.update({
        'MMPP_VIEWS': fmt_int(mmpp.get('Views')),
        'MMPP_REACH': fmt_int(mmpp.get('Alcance')),
        'MMPP_INTERACTIONS': fmt_int(mmpp.get('Interacciones')),
        'MMPP_ER': fmt_pct(mmpp.get('ER %')),
        'MMPP_CONTENTS': fmt_int(mmpp.get('Contenidos')),
        'MMPP_COMMENT': mmpp.get('Comentario general') or MISSING,
    })

    # v08: dynamic qualitative copy from the Google Sheet / Excel model.
    # Keys in the sheet are already aligned to tokens used by the PPT mapping.
    qualitative = get_qualitative_texts(wb, month, warnings)
    action_plan = get_action_plan_texts(wb, month, warnings)
    ctx.update(qualitative)
    ctx.update(action_plan)

    # Numeric values used by v08 to resize bars/shapes dynamically.
    def pv(platform, field):
        return to_float(platform_kpis.get(platform, {}).get(field), 0)

    ctx['_CHART_VALUES'] = {
        # Slide 4: channel portfolio.
        'slide4_views': [pv('Instagram', 'views'), pv('Facebook', 'views'), pv('TikTok', 'views')],
        'slide4_reach': [pv('Instagram', 'reach'), pv('Facebook', 'reach'), pv('TikTok', 'reach')],
        'slide4_er': [pv('Instagram', 'er'), pv('Facebook', 'er'), pv('TikTok', 'er')],
        # Slide 5: Instagram.
        'slide5_mix': [pv('Instagram', 'stories'), pv('Instagram', 'reels'), pv('Instagram', 'carousels')],
        'slide5_top': [to_float(platform_kpis.get('Instagram', {}).get('top1_value'), 0), to_float(platform_kpis.get('Instagram', {}).get('top2_value'), 0), to_float(platform_kpis.get('Instagram', {}).get('top3_value'), 0)],
        # Slide 6: Facebook.
        'slide6_top': [to_float(platform_kpis.get('Facebook', {}).get('top1_value'), 0), to_float(platform_kpis.get('Facebook', {}).get('top2_value'), 0), to_float(platform_kpis.get('Facebook', {}).get('top3_value'), 0)],
        'slide6_formats': [pv('Facebook', 'stories'), pv('Facebook', 'reels'), pv('Facebook', 'carousels'), pv('Facebook', 'posts')],
        # Slide 7: TikTok.
        'slide7_age': [pv('TikTok', 'age_18_24'), pv('TikTok', 'age_25_34'), pv('TikTok', 'age_35_44'), pv('TikTok', 'age_45_54'), pv('TikTok', 'age_55')],
        'slide7_top': [to_float(platform_kpis.get('TikTok', {}).get('top1_value'), 0), to_float(platform_kpis.get('TikTok', {}).get('top2_value'), 0), to_float(platform_kpis.get('TikTok', {}).get('top3_value'), 0)],
        # Slide 9: Squad.
        'slide9_views': [to_float(squad.get('Skarleth Labra', {}).get('Views'), 0), to_float(squad.get('Busquilla', {}).get('Views'), 0), to_float(squad.get('Camila Andrade', {}).get('Views'), 0), to_float(squad.get('Disley Ramos', {}).get('Views'), 0)],
        'slide9_er': [to_float(squad.get('Skarleth Labra', {}).get('ER %'), 0), to_float(squad.get('Busquilla', {}).get('ER %'), 0), to_float(squad.get('Camila Andrade', {}).get('ER %'), 0), to_float(squad.get('Disley Ramos', {}).get('ER %'), 0)],
    }

    # v12: media/assets context for Top 3, MMPP, Squad and Competition visual slides.
    ctx = enrich_media_context(wb, ctx, month, ctrl, warnings)

    ctx['_WARNINGS'] = warnings
    return ctx


GLOBAL_TEXT_TO_TOKEN = {
    'Maicao · Marzo 2026': 'TITLE_META',
    'MARZO 2026': 'MES_UPPER',
}

SLIDE_TEXT_TO_TOKEN = {
    1: {'16.5M': 'OVERVIEW_VIEWS_M','8.2M': 'OVERVIEW_REACH_M','110.7K': 'OVERVIEW_INTERACTIONS_K'},
    2: {
        '16.486.179': 'OVERVIEW_VIEWS','+24,5% vs febrero': 'OVERVIEW_VIEWS_VAR','8.211.420': 'OVERVIEW_REACH',
        '+0,2% vs febrero': 'OVERVIEW_REACH_VAR','1,3%': 'OVERVIEW_ER','-55,2% vs febrero': 'OVERVIEW_ER_VAR',
        '110.695': 'OVERVIEW_INTERACTIONS','-56,3% vs febrero': 'OVERVIEW_INTERACTIONS_VAR','$3.800.400': 'OVERVIEW_INVESTMENT',
        '+78,4% vs febrero': 'OVERVIEW_INVESTMENT_VAR','Febrero vs Marzo': 'MES_PREV','13,2M': 'PREV_VIEWS_M',
        '16,5M': 'OVERVIEW_VIEWS_M','235,7K': 'PREV_INTERACTIONS_K','110,7K': 'OVERVIEW_INTERACTIONS_K'},
    3: {'2.9%': 'PREV_ER_DOT','1.3%': 'CURR_ER_DOT','2.130.000': 'PREV_INVESTMENT','3.800.400': 'CURR_INVESTMENT','140': 'PREV_CONTENTS','156': 'OVERVIEW_CONTENTS'},
    4: {
        '8,2M': 'IG_VIEWS_M','3,2M': 'FB_VIEWS_M','4,5M': 'TT_VIEWS_M','4,0M': 'IG_REACH_M','1,9M': ['FB_REACH_M', 'TT_REACH_M'],
        '1,4%': 'IG_ER','0,2%': 'FB_ER','0,9%': 'TT_ER','332.382 seguidores · 58.723 interacciones': 'IG_FOLLOWERS_INTERACTIONS',
        '565.501 seguidores · 4.656 interacciones': 'FB_FOLLOWERS_INTERACTIONS','30.120 seguidores · 17.148 interacciones': 'TT_FOLLOWERS_INTERACTIONS'},
    5: {
        '332.382': 'IG_FOLLOWERS','+1.389 vs mes anterior': 'IG_FOLLOWERS_GROWTH','92': 'IG_CONTENTS','+10 vs febrero': 'IG_CONTENTS_VAR','1,4%': 'IG_ER','antes 3,2%': 'IG_ER_PREV_LINE','$1.670.000': 'IG_BUDGET','antes $850.000': 'IG_BUDGET_PREV_LINE','68': 'IG_STORIES','13': 'IG_REELS','10': 'IG_CAROUSELS','Bubble Maybelline': 'IG_TOP1_NAME','1,51M': 'IG_TOP1_VALUE','Check Bienvenida': 'IG_TOP2_NAME','763K': 'IG_TOP2_VALUE','Make up SAMY': 'IG_TOP3_NAME','1,69M': 'IG_TOP3_VALUE'},
    6: {
        '565.501': 'FB_FOLLOWERS','+430 vs mes anterior': 'FB_FOLLOWERS_GROWTH','3.223.306': 'FB_VIEWS','+69% aprox vs febrero': 'FB_VIEWS_VAR','1.905.477': 'FB_REACH','+63% aprox vs febrero': 'FB_REACH_VAR','0,2%': 'FB_ER','antes 0,3%': 'FB_ER_PREV_LINE','Prod. $1.000': 'FB_TOP1_NAME','526K': 'FB_TOP1_VALUE','Máscara Bubble': 'FB_TOP2_NAME','508K': 'FB_TOP2_VALUE','SAMY': 'FB_TOP3_NAME','531K': 'FB_TOP3_VALUE','19': 'FB_STORIES','9': 'FB_REELS','2': 'FB_CAROUSELS','1': 'FB_POSTS'},
    7: {
        '30.120': 'TT_FOLLOWERS','+2.768 vs mes anterior': 'TT_FOLLOWERS_GROWTH','4.522.490': 'TT_VIEWS','+15,7% vs febrero': 'TT_VIEWS_VAR','17.148': 'TT_INTERACTIONS','+7,8% vs febrero': 'TT_INTERACTIONS_VAR','0,9%': 'TT_ER','antes 0,7%': 'TT_ER_PREV_LINE','36%': 'TT_AGE_18_24','32%': 'TT_AGE_25_34','13%': 'TT_AGE_35_44','8%': 'TT_AGE_45_54','10%': 'TT_AGE_55','Maquillaje principiantes': 'TT_TOP1_NAME','2.880': 'TT_TOP1_VALUE','EGC Garnier': 'TT_TOP2_NAME','630': 'TT_TOP2_VALUE','UGC hábitos piel': 'TT_TOP3_NAME','1.387': 'TT_TOP3_VALUE'},
    9: {
        '4.637.075': 'SQUAD_VIEWS','2.020.826': 'SQUAD_REACH','60.146': 'SQUAD_INTERACTIONS','3,0%': 'SQUAD_ER','2,05M': 'SKAR_VIEWS_M','88K': 'BUSQUI_VIEWS_M','811K': 'CAMI_VIEWS_M','1,69M': 'DISLEY_VIEWS_M','4,5%': 'SKAR_ER','1,1%': 'BUSQUI_ER','0,8%': 'CAMI_ER','2,2%': 'DISLEY_ER'},
    10: {'3.336.786': 'MMPP_VIEWS','1.443.025': 'MMPP_REACH','13.463': 'MMPP_INTERACTIONS','1,0%': 'MMPP_ER','Mix de 14 contenidos': 'MMPP_CONTENTS_LINE'},
}


# v08: qualitative text mapping. These map existing template text fragments to
# editable placeholders in 15_Qualitative_Texts / 16_Action_Plan.
QUALITATIVE_TEXT_TO_TOKEN = {
    2: {
        'La inversión amplificó la visibilidad, pero el engagement cayó por cambio de contexto: de festivales/verano a vuelta a la rutina. La prioridad ahora es retención y conversación, no solo alcance.': 'EXEC_BUSINESS_READING',
        'Más alcance pagado': 'EXEC_TAG_1',
        'Menos chispa orgánica': 'EXEC_TAG_2',
        'TikTok motor': 'EXEC_TAG_3',
    },
    3: {
        'Cambio de temporada': 'CHANGE_TITLE_1',
        'La conversación espontánea bajó al salir de hitos de alto interés como festivales/verano.': 'CHANGE_BODY_1',
        'Inversión crece': 'CHANGE_TITLE_2',
        'El paid empuja visibilidad (+24,5% views), pero diluye la tasa al abrir públicos menos fidelizados.': 'CHANGE_BODY_2',
        'Nuevo mix creativo': 'CHANGE_TITLE_3',
        'EGC, UGC y creators acercan la marca. Los contenidos útiles y auténticos son los que más se guardan/comentan.': 'CHANGE_BODY_3',
        'Decisión recomendada: mantener escala, pero rediseñar el contenido para retener: hooks de 3 segundos + contenido útil + CTA conversacional.': 'CHANGE_DECISION',
    },
    4: {
        'Inspiración + cultura pop': 'ROLE_INSTAGRAM',
        'Comunidad masiva + ofertas': 'ROLE_FACEBOOK',
        'Descubrimiento + contenido útil': 'ROLE_TIKTOK',
        'Lectura: IG concentra visibilidad e interacción; Facebook sostiene comunidad y alcance eficiente; TikTok es la apuesta de crecimiento cualitativo por autenticidad.': 'PORTFOLIO_READING',
    },
    5: {
        '• Colaboraciones con creadoras: empujan comentarios desde sus comunidades.': 'IG_WHAT_WORKED_1',
        '• Cultura pop y educativo: contenido tipo “Te enseño a...” genera valor y conversión.': 'IG_WHAT_WORKED_2',
        '• Audiencia femenina dominante: 92,3% mujeres; mayor foco 25-34.': 'IG_WHAT_WORKED_3',
        'Pasar de piezas aisladas a series guardables: tutoriales, comparativas, hacks y CTA de conversación.': 'IG_OPTIMIZATION',
    },
    6: {
        'Facebook sigue siendo el canal de comunidad más grande. Las piezas de oferta generan mayor reacción y conversación, especialmente cuando el beneficio es directo.': 'FB_VISUAL_READING',
        'Las stories no tienen totales completos por error de plataforma antes del 20 de marzo.': 'FB_DATA_NOTE',
    },
    7: {
        'Funciona cuando habla el lenguaje de la plataforma: auténtico, real, educativo y menos publicitario.': 'TT_CONTENT_PRINCIPLE',
        'Replicar contenidos de trabajadoras y formatos con datos, precios y beneficios concretos. Potenciar CTA a comentar, guardar y seguir.': 'TT_OPPORTUNITY',
    },
    8: {
        'Volumen, frecuencia y exposición continua.': 'TT_MIX_UGC_DESC',
        'Cercanía, confianza y conversación con consejeras.': 'TT_MIX_EGC_DESC',
        'Menor volumen, mayor engagement por oportunidad cultural.': 'TT_MIX_CAMPAIGN_DESC',
        'Lectura: UGC debe seguir alimentando volumen; EGC debe usarse para confianza; campañas deben seleccionarse por oportunidad de conversación, no solo por calendario.': 'TT_MIX_READING',
    },
    9: {
        'Skar sostiene interacción alta; Disley abre volumen nuevo con tono humorístico y cercano. Para próximos meses, combinar perfiles de alto alcance con micro-nichos de expertise.': 'SQUAD_LEARNING',
        '70% creators de performance, 20% nichos estratégicos, 10% apuestas en tendencia.': 'SQUAD_RULE',
    },
    10: {
        'El mix permitió ampliar la presencia de productos, con comentarios positivos asociados a calidad y precios. El contenido tiene potencial para crecer si se empaqueta en series de beneficios concretos.': 'MMPP_QUAL_READING',
        'Calidad': 'MMPP_TAG_1',
        'Precio': 'MMPP_TAG_2',
        'Valor práctico': 'MMPP_TAG_3',
        'Próximo paso: convertir menciones de marca propia en “pruebas visuales” comparativas: antes/después, precio vs resultado, rutina completa.': 'MMPP_NEXT_STEP',
    },
    11: {
        'Liderazgo': 'COMP_FB_TITLE',
        'Maicao lidera comunidad; crecimiento estable, oportunidad en oferta y contenido gráfico.': 'COMP_FB_BODY',
        'Posición intermedia': 'COMP_IG_TITLE',
        'DBS domina estética; Maicao crece con contenido y creators.': 'COMP_IG_BODY',
        'Brecha de crecimiento': 'COMP_TT_TITLE',
        'Dr. Simi y el contenido social/humano marcan la conversación; Maicao tiene el mayor espacio para posicionarse.': 'COMP_TT_BODY',
        'Competidores activan Pinterest-like, creators, eventos y contenido humano. Maicao puede diferenciarse con utilidad + cercanía + retail expertise.': 'COMP_BENCHMARK',
    },
    12: {
        'Retención': 'PLAN_1_PILLAR',
        'Hooks de 3 segundos, captions con promesa clara y edición nativa para cada plataforma.': 'PLAN_1_ACTION',
        'Meta: subir ER': 'PLAN_1_META',
        'Contenido útil': 'PLAN_2_PILLAR',
        'Series guardables: tutoriales, hacks, “precio vs resultado”, rutinas y comparativas.': 'PLAN_2_ACTION',
        'Meta: más guardados': 'PLAN_2_META',
        'Conversación': 'PLAN_3_PILLAR',
        'CTA de comentario y votación. Preguntas directas por edad, necesidad y ocasión de uso.': 'PLAN_3_ACTION',
        'Meta: interacción real': 'PLAN_3_META',
        'Nichos': 'PLAN_4_PILLAR',
        'Talentos de skincare, maquillaje pro y Gen Z. Activaciones ocasionales con perfiles en tendencia.': 'PLAN_4_ACTION',
        'Meta: relevancia': 'PLAN_4_META',
        'KPIs de control: ER por formato · guardados por contenido · retención de video · crecimiento de seguidores · ratio paid/orgánico': 'PLAN_KPIS_CONTROL',
    },
}


def enrich_context(ctx):
    ctx['IG_FOLLOWERS_INTERACTIONS'] = f"{ctx.get('IG_FOLLOWERS',MISSING)} seguidores · {ctx.get('IG_INTERACTIONS',MISSING)} interacciones"
    ctx['FB_FOLLOWERS_INTERACTIONS'] = f"{ctx.get('FB_FOLLOWERS',MISSING)} seguidores · {ctx.get('FB_INTERACTIONS',MISSING)} interacciones"
    ctx['TT_FOLLOWERS_INTERACTIONS'] = f"{ctx.get('TT_FOLLOWERS',MISSING)} seguidores · {ctx.get('TT_INTERACTIONS',MISSING)} interacciones"
    for prefix in ['IG', 'FB', 'TT']:
        ctx[f'{prefix}_ER_PREV_LINE'] = f"antes {ctx.get(f'{prefix}_ER_PREV')}" if ctx.get(f'{prefix}_ER_PREV') not in [None, '', MISSING] else MISSING
        ctx[f'{prefix}_BUDGET_PREV_LINE'] = f"antes {ctx.get(f'{prefix}_BUDGET_PREV')}" if ctx.get(f'{prefix}_BUDGET_PREV') not in [None, '', MISSING] else MISSING
    return ctx


def resolve_token(token, replacements, original):
    if token == 'MMPP_CONTENTS_LINE':
        val = replacements.get('MMPP_CONTENTS')
        return f"Mix de {val if val not in [None, '', MISSING] else MISSING} contenidos"
    val = replacements.get(token, original)
    if val is None or val == "":
        return MISSING
    return str(val)


def replace_text_in_shape(shape, replacements, text_map, state):
    if not hasattr(shape, 'text_frame') or shape.text_frame is None:
        return 0
    count = 0
    for paragraph in shape.text_frame.paragraphs:
        runs = list(paragraph.runs)
        if not runs:
            continue
        full = ''.join(run.text for run in runs)
        new_full = full
        for old, token_spec in sorted(text_map.items(), key=lambda x: len(x[0]), reverse=True):
            if old not in new_full:
                continue
            if isinstance(token_spec, list):
                idx = state.get(old, 0)
                token = token_spec[min(idx, len(token_spec)-1)]
                state[old] = idx + 1
            else:
                token = token_spec
            val = resolve_token(token, replacements, old)
            new_full = new_full.replace(old, val, 1)
        if new_full != full:
            runs[0].text = new_full
            for run in runs[1:]:
                run.text = ''
            count += 1
    return count


def tune_kpi_cards(prs):
    """Small layout correction for KPI cards in slides 5-7.
    Moves growth labels slightly down and reduces their font size so they do not overlap with large numbers.
    """
    for slide_idx in [5, 6, 7]:
        if slide_idx > len(prs.slides):
            continue
        slide = prs.slides[slide_idx-1]
        for shape in slide.shapes:
            if not hasattr(shape, 'text_frame') or shape.text_frame is None:
                continue
            txt = shape.text.strip()
            if 'vs mes anterior' in txt or txt.startswith('antes '):
                # Move down a touch inside the KPI tile; reduce text size.
                shape.top = shape.top + Inches(0.08)
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(8.5)
            # Shrink very large follower/money numbers in KPI cards slightly.
            if re.fullmatch(r'(\$)?[0-9\.]+', txt):
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        if run.font.size and run.font.size.pt > 28:
                            run.font.size = Pt(25)




# -----------------------------------------------------------------------------
# Dynamic chart bars v08
# -----------------------------------------------------------------------------
# v08 stops relying on the internal PowerPoint shape order at runtime.
# It tags bars/labels with stable technical names and then updates shapes by name.
# If a template is still untagged, it will tag the known v02/v06 template once using
# the fallback indices below. After that, charts are robust to minor position edits.

BAR_SPECS = [
    # key, slide number, bar names, label names, fallback bar indices, fallback label indices
    ('slide4_views', 4,
     ['bar_channel_views_ig','bar_channel_views_fb','bar_channel_views_tt'],
     ['label_channel_views_ig','label_channel_views_fb','label_channel_views_tt'],
     [6,9,12], [8,11,14]),
    ('slide4_reach', 4,
     ['bar_channel_reach_ig','bar_channel_reach_fb','bar_channel_reach_tt'],
     ['label_channel_reach_ig','label_channel_reach_fb','label_channel_reach_tt'],
     [16,19,22], [18,21,24]),
    ('slide4_er', 4,
     ['bar_channel_er_ig','bar_channel_er_fb','bar_channel_er_tt'],
     ['label_channel_er_ig','label_channel_er_fb','label_channel_er_tt'],
     [26,29,32], [28,31,34]),

    ('slide5_mix', 5,
     ['bar_ig_mix_stories','bar_ig_mix_reels','bar_ig_mix_carousels'],
     ['label_ig_mix_stories','label_ig_mix_reels','label_ig_mix_carousels'],
     [22,25,28], [24,27,30]),
    ('slide5_top', 5,
     ['bar_ig_top_1','bar_ig_top_2','bar_ig_top_3'],
     ['label_ig_top_1','label_ig_top_2','label_ig_top_3'],
     [32,35,38], [34,37,40]),

    ('slide6_top', 6,
     ['bar_fb_top_1','bar_fb_top_2','bar_fb_top_3'],
     ['label_fb_top_1','label_fb_top_2','label_fb_top_3'],
     [22,25,28], [24,27,30]),
    ('slide6_formats', 6,
     ['bar_fb_format_stories','bar_fb_format_reels','bar_fb_format_carousels','bar_fb_format_post'],
     ['label_fb_format_stories','label_fb_format_reels','label_fb_format_carousels','label_fb_format_post'],
     [32,35,38,41], [34,37,40,43]),

    ('slide7_age', 7,
     ['bar_tt_age_18_24','bar_tt_age_25_34','bar_tt_age_35_44','bar_tt_age_45_54','bar_tt_age_55'],
     ['label_tt_age_18_24','label_tt_age_25_34','label_tt_age_35_44','label_tt_age_45_54','label_tt_age_55'],
     [22,25,28,31,34], [24,27,30,33,36]),
    ('slide7_top', 7,
     ['bar_tt_top_1','bar_tt_top_2','bar_tt_top_3'],
     ['label_tt_top_1','label_tt_top_2','label_tt_top_3'],
     [38,41,44], [40,43,46]),

    ('slide9_views', 9,
     ['bar_squad_views_skar','bar_squad_views_busquilla','bar_squad_views_cami','bar_squad_views_disley'],
     ['label_squad_views_skar','label_squad_views_busquilla','label_squad_views_cami','label_squad_views_disley'],
     [22,25,28,31], [24,27,30,33]),
    ('slide9_er', 9,
     ['bar_squad_er_skar','bar_squad_er_busquilla','bar_squad_er_cami','bar_squad_er_disley'],
     ['label_squad_er_skar','label_squad_er_busquilla','label_squad_er_cami','label_squad_er_disley'],
     [35,38,41,44], [37,40,43,46]),
]


def _shape_bottom(shape):
    return int(shape.top + shape.height)


def _safe_shape(slide, idx):
    try:
        return slide.shapes[idx]
    except Exception:
        return None


def _find_shape_by_name(slide, name):
    for shape in slide.shapes:
        if getattr(shape, 'name', None) == name:
            return shape
    return None


def tag_chart_shapes(prs):
    """Tag chart bars and labels with stable names when the template is untagged.

    This is intentionally conservative: it only uses fallback indices for the approved
    v02/v06 template structure. If a template has already been tagged, it leaves it alone.
    Returns a diagnostics list.
    """
    diagnostics = []
    tagged = 0
    missing = 0
    for key, slide_no, bar_names, label_names, bar_indices, label_indices in BAR_SPECS:
        if slide_no > len(prs.slides):
            diagnostics.append(f"{key}: slide {slide_no} no existe")
            missing += len(bar_names)
            continue
        slide = prs.slides[slide_no-1]
        # If all bars already named, do nothing.
        existing = sum(1 for n in bar_names if _find_shape_by_name(slide, n) is not None)
        if existing == len(bar_names):
            diagnostics.append(f"{key}: barras nombradas OK ({existing}/{len(bar_names)})")
            continue
        # Otherwise tag by fallback indices.
        for idx, name in zip(bar_indices, bar_names):
            sh = _safe_shape(slide, idx)
            if sh is not None:
                sh.name = name
                tagged += 1
            else:
                diagnostics.append(f"{key}: no encontre barra fallback index {idx} para {name}")
                missing += 1
        for idx, name in zip(label_indices, label_names):
            sh = _safe_shape(slide, idx)
            if sh is not None:
                sh.name = name
                tagged += 1
            else:
                diagnostics.append(f"{key}: no encontre label fallback index {idx} para {name}")
        diagnostics.append(f"{key}: tagged by fallback indices")
    diagnostics.append(f"Template tagging: {tagged} shapes tagged, {missing} missing")
    return diagnostics


def _collect_named_shapes(slide, names, fallback_indices=None):
    shapes = []
    missing = []
    for pos, name in enumerate(names):
        sh = _find_shape_by_name(slide, name)
        if sh is None and fallback_indices and pos < len(fallback_indices):
            # Emergency fallback for legacy templates.
            sh = _safe_shape(slide, fallback_indices[pos])
            if sh is not None:
                sh.name = name
        if sh is None:
            missing.append(name)
        shapes.append(sh)
    return shapes, missing


def _scale_named_bar_group(slide, values, bar_names, label_names=None, fallback_bar_indices=None, fallback_label_indices=None, min_height=12000):
    vals = [to_float(v, 0) or 0 for v in (values or [])]
    bars, missing_bars = _collect_named_shapes(slide, bar_names, fallback_bar_indices)
    labels, missing_labels = _collect_named_shapes(slide, label_names or [], fallback_label_indices)
    valid_bars = [b for b in bars if b is not None]
    if not valid_bars:
        return 0, missing_bars, missing_labels
    max_val = max(vals) if vals else 0
    if max_val <= 0:
        return 0, missing_bars, missing_labels
    # Use the largest existing bar as the chart maximum, preserving user design.
    max_height = max(int(b.height) for b in valid_bars)
    bottom = max(_shape_bottom(b) for b in valid_bars)
    changes = 0
    for pos, b in enumerate(bars):
        if b is None:
            continue
        value = vals[pos] if pos < len(vals) else 0
        new_h = min_height if value <= 0 else max(min_height, int(max_height * value / max_val))
        b.height = new_h
        b.top = bottom - new_h
        changes += 1
        if labels and pos < len(labels) and labels[pos] is not None:
            lab = labels[pos]
            lab.top = max(0, int(b.top - lab.height - 35000))
            changes += 1
    return changes, missing_bars, missing_labels


def update_dynamic_bars(prs, context):
    """v08: update editable PowerPoint bar shapes by stable object names."""
    values = context.get('_CHART_VALUES', {}) or {}
    diagnostics = []
    changes = 0
    tag_diags = tag_chart_shapes(prs)
    diagnostics.extend(tag_diags)
    missing_total = 0
    for key, slide_no, bar_names, label_names, bar_indices, label_indices in BAR_SPECS:
        if slide_no > len(prs.slides):
            diagnostics.append(f"{key}: slide {slide_no} no existe")
            continue
        changed, missing_bars, missing_labels = _scale_named_bar_group(
            prs.slides[slide_no-1],
            values.get(key, []),
            bar_names,
            label_names,
            bar_indices,
            label_indices,
        )
        changes += changed
        missing_total += len(missing_bars)
        if missing_bars:
            diagnostics.append(f"{key}: barras faltantes: {', '.join(missing_bars)}")
        else:
            diagnostics.append(f"{key}: barras actualizadas OK ({len(bar_names)})")
        if missing_labels:
            diagnostics.append(f"{key}: labels faltantes: {', '.join(missing_labels)}")
    context['_BAR_DIAGNOSTICS'] = diagnostics
    context['_BAR_CHANGES'] = changes
    context['_BAR_MISSING'] = missing_total
    return changes


def update_ppt(template_path, output_path, context, asset_service_account_info=None):
    context = enrich_context(context)
    prs = Presentation(template_path)
    changes = 0
    for idx, slide in enumerate(prs.slides, start=1):
        text_map = dict(GLOBAL_TEXT_TO_TOKEN)
        text_map.update(SLIDE_TEXT_TO_TOKEN.get(idx, {}))
        text_map.update(QUALITATIVE_TEXT_TO_TOKEN.get(idx, {}))
        state = {}
        for shape in slide.shapes:
            changes += replace_text_in_shape(shape, context, text_map, state)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        changes += replace_text_in_shape(cell, context, text_map, state)
    tune_kpi_cards(prs)
    changes += update_dynamic_bars(prs, context)
    context['_MEDIA_SLIDES_ADDED'] = append_media_slides(prs, context, asset_service_account_info)
    prs.save(output_path)
    return changes


def write_validation_report(path, ctx):
    warnings = ctx.get('_WARNINGS', [])
    bar_diags = ctx.get('_BAR_DIAGNOSTICS', [])
    lines = []
    lines.append("VALIDACION MAICAO REPORTING STUDIO v13")
    lines.append(f"Mes: {ctx.get('MES')}")
    lines.append("")
    if warnings:
        lines.append("Estado data: REVISAR")
        lines.append("")
        lines.append("Alertas de data:")
        for w in warnings:
            lines.append(f"- {w}")
    else:
        lines.append("Estado data: OK")
        lines.append("No se detectaron campos obligatorios faltantes.")
    lines.append("")
    lines.append("Estado textos cualitativos:")
    qualitative_keys = [k for k in ctx.keys() if k.isupper() and (k.startswith(('EXEC_', 'CHANGE_', 'ROLE_', 'PORTFOLIO_', 'IG_', 'FB_', 'TT_', 'SQUAD_', 'MMPP_', 'COMP_', 'PLAN_')))]
    qual_values = [ctx.get(k) for k in qualitative_keys if k not in {'IG_VIEWS','IG_REACH','IG_INTERACTIONS','IG_ER','FB_VIEWS','FB_REACH','FB_INTERACTIONS','FB_ER','TT_VIEWS','TT_REACH','TT_INTERACTIONS','TT_ER','SQUAD_VIEWS','SQUAD_REACH','SQUAD_INTERACTIONS','SQUAD_ER','MMPP_VIEWS','MMPP_REACH','MMPP_INTERACTIONS','MMPP_ER'}]
    lines.append(f"- Campos cualitativos cargados: {sum(1 for v in qual_values if v not in [None, '', MISSING])}")
    lines.append("")
    lines.append("Estado barras dinamicas:")
    lines.append(f"- Cambios aplicados: {ctx.get('_BAR_CHANGES', 0)}")
    lines.append(f"- Barras faltantes: {ctx.get('_BAR_MISSING', 0)}")
    if bar_diags:
        for d in bar_diags:
            lines.append(f"- {d}")
    else:
        lines.append("- No hay diagnostico de barras. Verifica que update_ppt haya corrido.")
    lines.append("")
    lines.append("Estado media/assets:")
    lines.append(f"- Slides visuales agregadas: {ctx.get('_MEDIA_SLIDES_ADDED', 0)}")
    top3 = ctx.get('_TOP3_MEDIA', {}) or {}
    for platform, items in top3.items():
        lines.append(f"- Top 3 {platform}: {len(items)}/3 piezas")
    path.write_text("\n".join(lines), encoding='utf-8')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', default=str(ROOT / 'Maicao_Reporte_Input_Model_v13_MEDIA_HISTORY.xlsx'))
    ap.add_argument('--template', default=str(ROOT / 'template' / 'Maicao_Template_Visual_v02.pptx'))
    ap.add_argument('--output', default=str(ROOT / 'output' / 'Maicao_Reporte_Auto_Template_v13.pptx'))
    ap.add_argument('--strict', action='store_true', help='Detener generacion si hay datos obligatorios faltantes.')
    args = ap.parse_args()
    ctx = build_context(args.input)
    Path(args.output).parent.mkdir(parents=True, exist_ok=True)
    report_path = Path(args.output).parent / 'validation_report_v13.txt'

    warnings = ctx.get('_WARNINGS', [])
    if warnings:
        print("VALIDACION: REVISAR")
        for w in warnings:
            print(f"- {w}")
        print(f"Reporte de validacion: {report_path}")
        if args.strict:
            print("Generacion detenida por --strict.")
            sys.exit(2)
    else:
        print("VALIDACION: OK")

    changes = update_ppt(args.template, args.output, ctx)
    write_validation_report(report_path, ctx)
    print(f"OK: {args.output}")
    print(f"Text replacements applied: {changes}")
    print("Fuente KPI plataforma: Excel hoja 13_Platform_KPIs")
    print(f"Barras dinamicas: {ctx.get('_BAR_CHANGES', 0)} cambios; faltantes {ctx.get('_BAR_MISSING', 0)}")
    print(f"Reporte de validacion: {report_path}")


if __name__ == '__main__':
    main()
