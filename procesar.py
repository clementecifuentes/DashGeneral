"""
procesar.py  —  Lee todos los archivos de datos/ y genera data.json para el dashboard.
Correr: python procesar.py
"""

import json
import re
from pathlib import Path
from datetime import date, datetime

import pandas as pd
import pdfplumber

BASE = Path(__file__).parent / "datos"
ROOT = Path(__file__).parent

# ── Mapeo de códigos de sucursal (viejo y nuevo) → nombre canónico ──────────
SUCURSAL_MAP = {
    # códigos nuevos (desde ago 2025)
    "10": "Central", "20": "Tandil", "30": "Tresa",
    "40": "Neco",    "21": "Juárez", "50": "Olava",
    # códigos viejos
    "12": "Central", "23": "Tandil", "33": "Tresa",  "32": "Tresa",
    "39": "Neco",    "9":  "Juárez", "19": "Juárez", "61": "Olava",
    "37": "Juárez",
    # TBD: 2 y 3 (nuevas sucursales, dejar vacío por ahora)
}

SUCURSALES_NORM = {
    "TRES ARROYOS": "Tresa", "TANDIL": "Tandil", "NECOCHEA": "Neco",
    "B JUAREZ": "Juárez", "OLAVARRIA": "Olava",
}

SUCURSALES = ["Tresa", "Tandil", "Neco", "Juárez", "Olava"]

# ── JP codes por categoría ───────────────────────────────────────────────────
JP_MONITOREO = {"jp267", "jp269"}   # Paquete de monitoreo (único criterio separado)

# ── Ciudad → sucursal (para hectáreas conectadas) ────────────────────────────
CIUDAD_SUC = {
    "Tandil": "Tandil", "Ayacucho": "Tandil", "Lobería": "Tandil", "Rauch": "Tandil",
    "Tres Arroyos": "Tresa", "Coronel Dorrego": "Tresa",
    "Adolfo Gonzales Chaves": "Tresa", "San Cayetano": "Tresa",
    "Necochea": "Neco", "Quequén": "Neco",
    "Benito Juárez": "Juárez", "Juárez": "Juárez",
    # Olavarría — agregar ciudades que aparezcan en el historial
    "Olavarría": "Olava", "Olavarria": "Olava", "Azul": "Olava",
    "Laprida": "Olava", "General La Madrid": "Olava",
}

FY_START_MONTH = 11  # noviembre


def latest_file(pattern: str) -> Path | None:
    files = sorted(BASE.glob(pattern), reverse=True)
    return files[0] if files else None

def all_files(pattern: str) -> list[Path]:
    """Todos los archivos que coinciden con el patrón, ordenados cronológicamente."""
    return sorted(BASE.glob(pattern))


def mes_fy(dt: date) -> int:
    """Mes dentro del FY: Nov=1, Dic=2, … Oct=12"""
    return ((dt.month - FY_START_MONTH) % 12) + 1


# ── 1. KPIs por sucursal (ISG sheet) ────────────────────────────────────────
def leer_kpis_isg() -> dict:
    path = latest_file("kpis_pan_fy26_*.xlsx")
    if not path:
        return {}

    df = pd.read_excel(path, sheet_name="ISG", header=None)
    result = {}
    current = None

    for _, row in df.iterrows():
        val0 = str(row[0]).strip() if pd.notna(row[0]) else ""

        # Nueva sucursal (case-insensitive)
        if val0.upper() in SUCURSALES_NORM:
            current = SUCURSALES_NORM[val0.upper()]
            result[current] = {}
            continue

        if current is None or val0 in ("", "Conectividad", "nan"):
            continue

        # Fila de KPI
        try:
            obj       = float(row[1]) if pd.notna(row[1]) else None
            real_suc  = float(row[2]) if pd.notna(row[2]) else None
            cumpl     = float(row[3]) if pd.notna(row[3]) else (real_suc / obj if obj and real_suc else 0)
            csc       = float(row[5]) if len(row) > 5 and pd.notna(row[5]) else None
            real_jd   = float(row[6]) if len(row) > 6 and pd.notna(row[6]) else None
        except (TypeError, ValueError):
            continue

        result[current][val0] = {
            "objetivo":      round(obj, 2)      if obj      is not None else None,
            "real":          round(real_suc, 2) if real_suc is not None else None,
            "real_csc":      round(csc, 2)      if csc      is not None else None,
            "real_total_jd": round(real_jd, 2)  if real_jd  is not None else None,
            "cumplimiento":  round(min(cumpl, 1.0), 4) if cumpl is not None else 0,
        }

    # ── Sobreescribir Ecosistemas Conectados desde PABLO.ECOSISTEMAS OBJETIVOS ──
    eco_path = BASE / "PABLO.ECOSISTEMAS OBJETIVOS.xlsx"
    if eco_path.exists():
        df_eco = pd.read_excel(eco_path, header=None)
        ECO_NORM = {
            "TANDIL": "Tandil", "NECOCHEA": "Neco", "TRES ARROYOS": "Tresa",
            "B. JUÁREZ": "Juárez", "B. JUAREZ": "Juárez", "B JUÁREZ": "Juárez",
            "B JUAREZ": "Juárez", "OLAVARRIA": "Olava", "OLAVARRÍA": "Olava",
        }
        for _, row in df_eco.iterrows():
            lbl = str(row.iloc[1]).strip().upper() if pd.notna(row.iloc[1]) else ""
            suc = ECO_NORM.get(lbl)
            if not suc:
                continue
            try:
                obj  = float(row.iloc[2]) if pd.notna(row.iloc[2]) else None
                real = float(row.iloc[3]) if pd.notna(row.iloc[3]) else None
            except (TypeError, ValueError):
                continue
            if obj is None and real is None:
                continue
            cumpl = round(min(real / obj, 1.0), 4) if obj and real else 0
            if suc not in result:
                result[suc] = {}
            result[suc]["Ecosistemas Conectados"] = {
                "objetivo":      obj,
                "real":          real,
                "real_csc":      None,
                "real_total_jd": None,
                "cumplimiento":  cumpl,
            }

    return result


# ── 2. Paquetes de monitoreo ─────────────────────────────────────────────────
def mover_vencidos_a_bajas(path: Path, hoy: date) -> None:
    """Mueve contratos de monitoreo vencidos de la hoja Monitoreo a Bajas."""
    from openpyxl import load_workbook

    wb = load_workbook(path)
    ws_mon = wb["Monitoreo"]
    ws_baj = wb["Bajas"]

    # Leer encabezados de ambas hojas
    headers_mon = [c.value for c in next(ws_mon.iter_rows(min_row=1, max_row=1))]
    headers_baj = [c.value for c in next(ws_baj.iter_rows(min_row=1, max_row=1))]

    # Índice de la columna Vencimiento en Monitoreo
    try:
        idx_venc = headers_mon.index("Vencimiento contrato")
    except ValueError:
        return

    # Series de Bajas ya existentes para evitar duplicados
    try:
        idx_serie_baj = headers_baj.index("numero de serie")
        series_bajas = {
            ws_baj.cell(r, idx_serie_baj + 1).value
            for r in range(2, ws_baj.max_row + 1)
        }
    except ValueError:
        series_bajas = set()

    try:
        idx_serie_mon = headers_mon.index("numero de serie")
    except ValueError:
        idx_serie_mon = None

    rows_a_mover = []
    rows_a_mantener = [list(next(ws_mon.iter_rows(min_row=1, max_row=1, values_only=True)))]

    for row in ws_mon.iter_rows(min_row=2, values_only=True):
        venc = row[idx_venc]
        if venc is None:
            rows_a_mantener.append(list(row))
            continue
        if hasattr(venc, "date"):
            venc_date = venc.date()
        else:
            try:
                venc_date = pd.to_datetime(venc).date()
            except Exception:
                rows_a_mantener.append(list(row))
                continue

        serie = row[idx_serie_mon] if idx_serie_mon is not None else None
        if venc_date < hoy and serie not in series_bajas:
            rows_a_mover.append(list(row))
            if serie:
                series_bajas.add(serie)
        else:
            rows_a_mantener.append(list(row))

    if not rows_a_mover:
        return

    # Agregar vencidos a Bajas
    for row in rows_a_mover:
        # Alinear columnas con Bajas (rellenar con None si faltan)
        row_baj = [None] * len(headers_baj)
        for i, val in enumerate(row):
            if i < len(headers_baj):
                row_baj[i] = val
        ws_baj.append(row_baj)

    # Reescribir Monitoreo sin los vencidos
    for row in ws_mon.iter_rows(min_row=2, max_row=ws_mon.max_row):
        for cell in row:
            cell.value = None
    for i, row_data in enumerate(rows_a_mantener[1:], start=2):
        for j, val in enumerate(row_data, start=1):
            ws_mon.cell(row=i, column=j, value=val)

    import tempfile, shutil, os
    tmp = Path(tempfile.mktemp(suffix=".xlsx", dir=path.parent))
    try:
        wb.save(tmp)
        shutil.move(str(tmp), str(path))
        print(f"  -> {len(rows_a_mover)} contrato(s) vencido(s) movido(s) a Bajas")
    except PermissionError:
        if tmp.exists():
            os.remove(tmp)
        print("  ! Archivo abierto en otra app — cerrar Excel y volver a correr procesar.py")


def leer_paquetes() -> dict:
    # Usar ContactoPaquetesvigentes.xlsx en datos/ (fuente completa)
    path = BASE / "ContactoPaquetesvigentes.xlsx"
    if not path.exists():
        # Fallback al archivo en datos/
        path = latest_file("paquetes_fy26_*.xlsx")
    if not path:
        return {}

    hoy = date.today()

    monitoreo_raw = pd.read_excel(path, sheet_name="Listado",
                                  usecols=["Nombre", "Telefono", "Sucursal", "Tipo de Paquete",
                                           "maquina", "Vencimiento contrato", "Producto", "JP"])
    monitoreo_raw = monitoreo_raw.dropna(how="all")
    monitoreo_raw["Vencimiento contrato"] = pd.to_datetime(
        monitoreo_raw["Vencimiento contrato"], errors="coerce")

    NORM = {"Tres Arroyos": "Tresa", "Necochea": "Neco",
            "Olavarria": "Olava", "Olavarría": "Olava",
            "Juarez": "Juárez", "B Juarez": "Juárez"}
    monitoreo_raw["Sucursal"] = monitoreo_raw["Sucursal"].str.strip().str.title().replace(NORM)

    # Solo contratos de monitoreo con paquete Plata/Oro (excluir señales, licencias sueltas, etc.)
    monitoreo = monitoreo_raw[
        (monitoreo_raw["Producto"].str.lower().str.strip() == "contrato de monitoreo") &
        (monitoreo_raw["Tipo de Paquete"].isin(["Plata", "Oro"]))
    ].copy()

    mon_suc = monitoreo.groupby("Sucursal").size().to_dict()
    # Excluir sucursales irrelevantes (Azul es zona de Olava, ya cubierta)
    SUCURSALES_SET = set(SUCURSALES) | {"Central"}
    sucursales_list = sorted(s for s in mon_suc if s in SUCURSALES_SET)

    por_sucursal = [
        {"sucursal": s, "monitoreo": int(mon_suc.get(s, 0))}
        for s in sucursales_list
    ]

    # Contratos que vencen en los próximos 60 días
    hoy_ts   = pd.Timestamp(hoy)
    limite   = hoy_ts + pd.Timedelta(days=60)
    venc_col = monitoreo["Vencimiento contrato"]
    vencen_60d = monitoreo[
        venc_col.notna() & (venc_col >= hoy_ts) & (venc_col <= limite)
    ]
    vencen_60d_suc = vencen_60d.groupby("Sucursal").size().to_dict()

    # Lista de contratos vigentes para la tabla detalle
    contratos_lista = []
    for _, row in monitoreo.sort_values("Sucursal").iterrows():
        venc = row["Vencimiento contrato"]
        venc_str = venc.strftime("%Y-%m-%d") if pd.notna(venc) else None
        dias_restantes = int((venc - hoy_ts).days) if pd.notna(venc) else None
        tel_val = row.get("Telefono", None)
        if pd.notna(tel_val):
            try:
                tel = str(int(float(tel_val)))
            except (ValueError, TypeError):
                tel = str(tel_val).strip()
            if tel in ("nan", "None"):
                tel = ""
        else:
            tel = ""
        contratos_lista.append({
            "nombre":         str(row["Nombre"]).strip(),
            "sucursal":       str(row["Sucursal"]).strip(),
            "tipo":           str(row["Tipo de Paquete"]).strip(),
            "maquina":        str(row["maquina"]).strip() if pd.notna(row["maquina"]) else "",
            "telefono":       tel,
            "vencimiento":    venc_str,
            "dias_restantes": dias_restantes,
        })

    return {
        "total_monitoreo":  len(monitoreo),
        "vencen_60d_total": len(vencen_60d),
        "vencen_60d_suc":   {s: int(vencen_60d_suc.get(s, 0)) for s in sucursales_list},
        "por_sucursal":     por_sucursal,
        "contratos_lista":  contratos_lista,
    }


# ── 2b. Bajas de monitoreo ───────────────────────────────────────────────────
def leer_bajas() -> dict:
    path = BASE / "ContactoPaquetesvigentes.xlsx"
    if not path.exists():
        return {"total": 0, "lista": []}

    try:
        df = pd.read_excel(path, sheet_name="Bajas", header=1)
    except Exception:
        return {"total": 0, "lista": []}

    NORM_BAJ = {
        "Tres Arroyos": "Tresa", "Necochea": "Neco",
        "Olavarria": "Olava", "Olavarría": "Olava",
        "Juarez": "Juárez", "B Juarez": "Juárez",
    }

    lista = []
    for _, row in df.iterrows():
        nombre = str(row.get("Nombre", "")).strip()
        if not nombre or nombre in ("nan", "None"):
            continue
        venc = row.get("Vencimiento contrato", None)
        venc_str = None
        venc_raw = str(venc).strip() if pd.notna(venc) else ""
        if venc_raw and not venc_raw.startswith("-") and venc_raw not in ("nan", "None", ""):
            try:
                venc_dt = pd.to_datetime(venc_raw, errors="coerce")
                if pd.notna(venc_dt):
                    venc_str = venc_dt.strftime("%Y-%m-%d")
            except Exception:
                pass
        serie = str(row.get("numero de serie", "")).strip()
        suc_raw = str(row.get("Sucursal", "")).strip()
        if suc_raw in ("nan", "None", ""):
            continue
        suc = NORM_BAJ.get(suc_raw, suc_raw)

        tel_val = row.get("Telefono", None)
        if pd.notna(tel_val):
            try:
                tel = str(int(float(tel_val)))
            except (ValueError, TypeError):
                tel = str(tel_val).strip()
            if tel in ("nan", "None"):
                tel = ""
        else:
            tel = ""

        lista.append({
            "nombre":      nombre,
            "sucursal":    suc,
            "vencimiento": venc_str,
            "serie":       serie if serie not in ("nan", "None") else "",
            "telefono":    tel,
        })

    return {"total": len(lista), "lista": lista}


# ── 3. Activaciones ──────────────────────────────────────────────────────────
def _leer_sheet_activaciones(path: Path, sheet: str) -> pd.DataFrame:
    """Lee una hoja de activaciones del Excel master y normaliza columnas."""
    try:
        df = pd.read_excel(path, sheet_name=sheet)
    except Exception:
        return pd.DataFrame()
    df.columns = df.columns.str.strip()
    df["Fecha"] = pd.to_datetime(df.get("Fecha"), errors="coerce")
    df = df.dropna(subset=["Fecha"])
    suc_col = "Unnamed: 2" if "Unnamed: 2" in df.columns else ("Uen" if "Uen" in df.columns else None)
    if suc_col:
        df["suc_cod"] = pd.to_numeric(df[suc_col], errors="coerce").fillna(0).astype(int).astype(str)
    else:
        df["suc_cod"] = "0"
    df["sucursal"] = df["suc_cod"].map(SUCURSAL_MAP).fillna("Otro")
    usd_col = "Resultado USD" if "Resultado USD" in df.columns else None
    if usd_col:
        df["resultado_usd"] = pd.to_numeric(
            df[usd_col].astype(str).str.replace(",", "."), errors="coerce").fillna(0)
    else:
        df["resultado_usd"] = 0.0
    tarea_col = "Tarea" if "Tarea" in df.columns else None
    if tarea_col:
        df["tarea"] = df[tarea_col]
    else:
        df["tarea"] = ""
    return df[["Fecha", "sucursal", "tarea", "resultado_usd"]]

def leer_activaciones() -> dict:
    # Busca el Excel master (puede haber varios — toma el más reciente)
    path = latest_file("activaciones_fy26_master*.xlsx")
    if not path:
        return {}

    hoy = date.today()
    fy_inicio = date(hoy.year if hoy.month >= FY_START_MONTH else hoy.year - 1,
                     FY_START_MONTH, 1)

    # Determina qué hojas leer según el FY actual
    fy_label = str(fy_inicio.year + 1)  # FY26 → hoja "Activaciones 2026"
    sheet_name = f"Activaciones {fy_label}"

    df = _leer_sheet_activaciones(path, sheet_name)
    if df.empty:
        return {}

    df_fy = df[df["Fecha"].dt.date >= fy_inicio]

    mes_actual = hoy.month
    anio_actual = hoy.year
    df_mes = df_fy[(df_fy["Fecha"].dt.month == mes_actual) &
                   (df_fy["Fecha"].dt.year == anio_actual)]

    por_sucursal_fy = (df_fy.groupby("sucursal")
                       .agg(activaciones=("tarea", "count"), usd=("resultado_usd", "sum"))
                       .reset_index()
                       .to_dict("records"))

    return {
        "total_fy": int(len(df_fy)),
        "usd_fy": round(float(df_fy["resultado_usd"].sum()), 2),
        "total_mes": int(len(df_mes)),
        "usd_mes": round(float(df_mes["resultado_usd"].sum()), 2),
        "fy_inicio": fy_inicio.strftime("%Y-%m-%d"),
        "por_sucursal": por_sucursal_fy,
    }


# ── 4. Precision Ag (PDF) ────────────────────────────────────────────────────
def leer_precision_ag() -> dict:
    root_pdfs = sorted(BASE.glob("Precision Ag*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
    path = root_pdfs[0] if root_pdfs else latest_file("precision_ag_fy26_*.pdf")
    if not path:
        return {}

    result = {}
    # data_refreshed usa re.DOTALL porque la fecha aparece en la línea siguiente al label
    patterns = {
        "highly_engaged_orgs": r"Highly Engaged Orgs\s+([\d,\.]+)",
        "digitally_engaged_orgs": r"Digitally Engaged Orgs\s+([\d,\.]+)",
        "opportunity_ha": r"Opportunity\s+([\d\.]+)\s+100%",
        "engaged_ha_r12": r"Current Engaged (?:Hectares|Acre) \(R12\)\s+([\d,\.]+)",
        "connected_machines": r"Connected Ag Machines\s+([\d,\.]+)",
        "data_refreshed": r"Data Refreshed as of:.*?(\d+/\d+/\d+)",
    }

    text = ""
    with pdfplumber.open(path) as pdf:
        page0 = pdf.pages[0]
        for page in pdf.pages[:2]:
            t = page.extract_text()
            if t:
                text += t + "\n"

        # highly_engaged_ha_5step: el texto principal está garbled por superposición
        # de capas de texto en el PDF; se extrae con bounding box de la columna "5 Step"
        try:
            crop5 = page0.crop((330, 255, 395, 365))
            t5 = crop5.extract_text() or ""
            m5 = re.search(r"5 Steps\s+([\d.]+)K", t5)
            if m5:
                result["highly_engaged_ha_5step"] = round(float(m5.group(1)) * 1000, 0)
        except Exception:
            pass

    def parse_es_num(s: str):
        """Convierte número en formato español (1.234,56 o 1.234) a float."""
        s = s.strip()
        if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", s):
            return s  # es una fecha, no un número
        # si tiene coma decimal: 787,594 → 787594 o 787.594
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        else:
            s = s.replace(".", "")
        try:
            return float(s)
        except ValueError:
            return s

    for key, pat in patterns.items():
        m = re.search(pat, text, re.DOTALL)
        if m:
            result[key] = parse_es_num(m.group(1))

    return result


# ── 5. Resumen Ventas AMS (Paquete Esencial + Licencias) ────────────────────
def _extraer_uen_cuenta(cuenta: str) -> str | None:
    try:
        return str(int(str(cuenta).split("-")[0].strip()))
    except (ValueError, AttributeError):
        return None


def _extraer_jp(articulo: str) -> str:
    m = re.search(r"JP0+(\d+)", str(articulo), re.IGNORECASE)
    return ("jp" + str(int(m.group(1)))) if m else "otro"


def _cat_jp(jp: str) -> str:
    return "Paq. Monitoreo" if jp in JP_MONITOREO else "Otras"


def leer_ams() -> dict:
    path = BASE / "datos" / "Sol.PUKs.xlsx"
    if not path.exists():
        path = BASE / "Sol.PUKs.xlsx"
    if not path.exists():
        return {}

    wb = pd.read_excel(path, sheet_name=None)

    # ── Paquete Esencial: hoja "PUKs x Sucursal", bloque "2026 PUK" ────────────
    paquete_esencial = {}
    _puk_sheet = next((s for s in wb if s.strip() == "PUKs x Sucursal"), None)
    if _puk_sheet:
        df_puk = wb[_puk_sheet]
        # Buscar fila cabecera del bloque: col A = "Sucursal"
        hdr_idx = df_puk[df_puk.iloc[:, 0].astype(str).str.strip().str.lower() == "sucursal"].index
        if len(hdr_idx):
            last_hdr = hdr_idx[-1]  # último bloque = "2026 PUK" (el más reciente)
            header_row = df_puk.iloc[last_hdr]
            # Columna TOTAL: última columna con valor "TOTAL"
            total_col_idx = None
            for j, v in enumerate(header_row):
                if str(v).strip().upper() == "TOTAL":
                    total_col_idx = j
            if total_col_idx is None:
                total_col_idx = df_puk.shape[1] - 1
            # Leer filas de datos hasta encontrar "TOTAL" en col A
            for i in range(last_hdr + 1, last_hdr + 15):
                if i >= len(df_puk):
                    break
                row = df_puk.iloc[i]
                raw = str(row.iloc[0]).strip()
                if raw.upper().startswith("TOTAL") or raw in ("", "nan"):
                    continue
                m = re.search(r"-\s*(\d+)$", raw)
                if not m:
                    continue
                suc = SUCURSAL_MAP.get(m.group(1))
                if suc:
                    val = row.iloc[total_col_idx]
                    total = int(val) if pd.notna(val) and str(val) != "nan" else 0
                    paquete_esencial[suc] = paquete_esencial.get(suc, 0) + total

    # ── Licencias: desde activaciones_fy26_master, hoja "Activaciones 2026" ──
    act_path = BASE / "activaciones_fy26_master.xlsx"
    licencias_suc = {suc: {"total": 0, "otras": 0} for suc in SUCURSALES}
    if act_path.exists():
        df_act = pd.read_excel(act_path, sheet_name="Activaciones 2026")
        df_act["uen"]      = pd.to_numeric(df_act["Unnamed: 2"], errors="coerce").fillna(0).astype(int).astype(str)
        df_act["sucursal"] = df_act["uen"].map(SUCURSAL_MAP).fillna("Sin mapear")
        df_act["jp"]       = df_act["Tarea"].astype(str).str.strip().str.lower()
        for suc in SUCURSALES:
            sub = df_act[df_act["sucursal"] == suc]
            otras = sub[~sub["jp"].isin(JP_MONITOREO)]
            licencias_suc[suc] = {
                "total": int(len(sub)),
                "otras": int(len(otras)),
            }

    return {
        "paquete_esencial": paquete_esencial,
        "licencias":        licencias_suc,
    }


# ── 6. Conectividad: Hectáreas + Máquinas ────────────────────────────────────
def leer_conectividad() -> dict:
    # ── Hectáreas: usar el archivo más reciente de Historial ─────────────────
    historial_dir = (BASE / "Historial Hectareas Conectadas") if (BASE / "Historial Hectareas Conectadas").exists() \
                    else (ROOT / "Historial Hectareas Conectadas")
    historial_files = (
        [f for f in historial_dir.glob("*.xlsx") if "OBJETIVOS" not in f.name.upper() and not f.name.startswith("~$")]
        if historial_dir.exists() else []
    )
    ha_path = max(historial_files, key=lambda p: p.stat().st_mtime) if historial_files else None

    hectareas = {}
    if ha_path:
        import openpyxl as _opxl
        _wb_do = _opxl.load_workbook(ha_path, data_only=True)
        _ws = _wb_do[_wb_do.sheetnames[1]]  # "2026.Héctareas Conecta"

        # Columnas openpyxl (1-based) para cada sucursal — confirmadas por el usuario
        _SUC_COLS = {"Tandil": 3, "Neco": 7, "Tresa": 11, "Juárez": 15, "Olava": 20}
        # Filas de totales de bloque donde está el AOR (col openpyxl 8)
        # (row, col_openpyxl) — Olava sin AOR propio (mezclado con Azul, no separable)
        _AOR_CELLS = {"Tandil": (23, 8), "Neco": (31, 8), "Tresa": (39, 8), "Juárez": (46, 8)}

        def _cv(row, col):
            try:
                v = _ws.cell(row=row, column=col).value
                return round(float(v), 2) if v is not None else 0.0
            except Exception:
                return 0.0

        for _suc, _col in _SUC_COLS.items():
            _ha_csc      = _cv(87, _col)   # Ha conectadas aportadas por CSC
            _ha_suc      = _cv(89, _col)   # Ha conectadas aportadas por suc
            _ha_alt_csc  = _cv(88, _col)   # Profundidad aportadas por CSC
            _ha_alt_suc  = _cv(90, _col)   # Profundidad aportadas por suc
            _ha_alt_obj  = _cv(84, _col)   # Objetivo profundidad ha
            _has_tot     = round(_ha_csc + _ha_suc, 2)
            _ha_alt_tot  = round(_ha_alt_csc + _ha_alt_suc, 2)
            _aor_cell = _AOR_CELLS.get(_suc)
            _aor      = _cv(*_aor_cell) if _aor_cell else 0.0
            hectareas[_suc] = {
                "has_conectadas":     _has_tot,
                "ha_csc":             _ha_csc,
                "ha_suc":             _ha_suc,
                "ha_alt_conectadas":  _ha_alt_tot,   # CSC + suc (igual lógica que has_conectadas)
                "ha_alt_csc":         _ha_alt_csc,
                "ha_alt_suc":         _ha_alt_suc,
                "ha_alt_objetivo":    _ha_alt_obj,
                "ha_altamente_excel": 0.0,   # se llena desde tabla resumen col L
                "area_oportunidad":   _aor,
                "pct_cumplimiento":   round(_has_tot / _aor, 4) if _aor > 0 else 0.0,
            }

        # Tabla resumen filas 98-105:
        #   col F (idx 5)  → ha_alt_conectadas total por sucursal y fila Totales
        #   col L (idx 11) → ha_altamente_excel (tabla comparativa)
        # ha_alt_conectadas de sucursales ya fue calculado como row88+row90,
        # pero la fila "Totales" se lee aquí para el KPI global
        _ALTEX_MAP = {
            "central":      "Central",
            "necochea":     "Neco",
            "tres arroyos": "Tresa",
            "juarez":       "Juárez",
            "juárez":       "Juárez",
            "tandil":       "Tandil",
            "azul":         "Olava",
            "totales":      "_prof_total",
        }
        for _row in _ws.iter_rows(min_row=98, max_row=105, values_only=True):
            _lbl = str(_row[0]).strip().lower() if _row[0] else ""
            _suc = _ALTEX_MAP.get(_lbl)
            if not _suc:
                continue
            if _suc not in hectareas:
                hectareas[_suc] = {"ha_alt_conectadas": 0.0, "ha_altamente_excel": 0.0}
            # Para "_total": leer el total directo de la fila Totales del Excel
            if _suc == "_prof_total":
                try:
                    hectareas["_prof_total"]["ha_alt_conectadas"] = round(float(_row[5]), 2)
                except Exception:
                    pass
            try:
                hectareas[_suc]["ha_altamente_excel"] = round(float(_row[11]), 2)
            except Exception:
                pass

        print(f"  Hectareas: {ha_path.name}")

    # ── Máquinas conectadas ───────────────────────────────────────────────────
    # Fuente: Ezequiel.Maquinas conectadas.xlsx (columnas Ciudad + Registration Status)
    CIUDAD_MAQ = {
        "TANDIL":       "Tandil",
        "NECOCHEA":     "Neco",
        "TRES ARROYOS": "Tresa",
        "BTO. JUAREZ":  "Juárez",
        "OLAVARRIA":    "Olava",
        "OLAVARRÍA":    "Olava",
        "STOCK":        "Central",
    }
    maquinas = {}
    maq_path = BASE / "Ezequiel.Maquinas conectadas.xlsx"
    if maq_path.exists():
        df_maq     = pd.read_excel(maq_path)
        col_ciudad = next((c for c in df_maq.columns if "Ciudad" in str(c) or "CIUDAD" in str(c)), None)
        col_status = next((c for c in df_maq.columns if "Registration Status" in str(c)), None)
        if col_ciudad:
            df_maq["sucursal"] = df_maq[col_ciudad].fillna("STOCK").astype(str).str.strip().str.upper().map(
                lambda x: CIUDAD_MAQ.get(x, "Sin mapear")
            )
            for suc in ["Central"] + SUCURSALES:
                sub = df_maq[df_maq["sucursal"] == suc]
                if len(sub) > 0:
                    maquinas[suc] = {
                        "conectadas": int(len(sub)),
                        "total":      int(len(sub)),
                    }
        print(f"  Maquinas: {maq_path.name}")

    return {"hectareas": hectareas, "maquinas": maquinas}


# ── 7. CSC One Page Report (PDF) ─────────────────────────────────────────────
def leer_csc() -> dict:
    root_pdfs = sorted(list((BASE / "datos").glob("One Page Report CSC*.pdf")) + list(BASE.glob("One Page Report CSC*.pdf")),
                       key=lambda p: p.stat().st_mtime, reverse=True)
    path = root_pdfs[0] if root_pdfs else None
    if not path:
        return {}

    pages = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            pages.append(page.extract_text() or "")

    p1 = pages[0] if pages else ""
    p3 = pages[2] if len(pages) > 2 else ""
    result = {}

    # Fecha de actualización
    m = re.search(r"Last Updated\s+(\d+/\d+/\d+)", p1)
    if m:
        result["last_updated"] = m.group(1)

    # ── % Service CSC (FYTD) ─────────────────────────────────────────────────
    # El bloque tiene 6 gauges en fila; pdfplumber extrae los valores de los gauges
    # más llenos (derecha) primero: 78.26%, 64.62%, luego en la misma línea los más
    # bajos: 11.57% 4.3% 17.2% 78.5%. El primer valor de esa tercera línea = % Service CSC.
    m = re.search(r"% Service CSC \(FYTD\).*?[\d.]+%\n[\d.]+%\n([\d.]+)%", p1, re.DOTALL)
    if m:
        result["pct_service_csc"] = float(m.group(1)) / 100

    # ── Expert Alerts Worked (FYTD) ──────────────────────────────────────────
    # Gauge: "96.72%\nGoal: 85.0%"
    m = re.search(r"% EA Worked.*?([\d.]+%)\s*\n.*?Goal:\s*([\d.]+)%", p1, re.DOTALL)
    if m:
        result["expert_alerts_worked_pct"] = float(m.group(1).replace("%", "")) / 100
        result["expert_alerts_worked_obj"] = float(m.group(2)) / 100
    else:
        result["expert_alerts_worked_obj"] = 0.85

    # ── % EA Conversion (High & Critical) ────────────────────────────────────
    # El gauge va precedido del icono "supera meta" en el PDF; ese icono es el ancla
    # correcta (de lo contrario re.search agarra el primer % de toda la pagina).
    m = re.search(chr(0xE718) + r"\s*([\d.]+)%", p1)
    if m:
        result["expert_alerts_conversion_pct"] = float(m.group(1)) / 100
    else:
        m2 = re.search(r"([\d.]+)%\s*\nGoal:\s*10\.0%", p1)
        if m2:
            result["expert_alerts_conversion_pct"] = float(m2.group(1)) / 100
    result["expert_alerts_conversion_obj"] = 0.10

    # ── % Tech x SAR (FYTD) ──────────────────────────────────────────────────
    # Más confiable: calcular de "Total 3" (SAR techs) / "Total 21" (all techs)
    # en la sección "List of technicians Tech that use SAR"
    m_techs = re.search(
        r"List of technicians Tech that use SAR.*?Total\s+(\d+).*?Total\s+(\d+)",
        p1, re.DOTALL
    )
    if m_techs:
        t1, t2 = int(m_techs.group(1)), int(m_techs.group(2))
        sar_techs   = min(t1, t2)
        total_techs = max(t1, t2)
        result["pct_tech_sar"] = round(sar_techs / total_techs, 4) if total_techs > 0 else 0.0
    else:
        # Fallback: buscar "(Blank) XX.XX%" del gráfico de barras
        m = re.search(r"\(Blank\)\s*([\d.]+%)", p1)
        if m:
            result["pct_tech_sar"] = float(m.group(1).replace("%", "")) / 100
    result["pct_tech_sar_obj"] = 0.40   # objetivo: 40%

    # ── CSC Service Revenue ───────────────────────────────────────────────────
    m = re.search(r"CSC Service Revenue.*?([\d.]+K)\s+[\d.]+K", p1, re.DOTALL)
    if m:
        result["csc_service_revenue_k"] = float(m.group(1).replace("K", ""))

    # ── Proactive Connections R12 y Meaningful Proactive Connections R12 ──────
    # En el PDF aparecen como "1.2" y "1.0" = 1.2% y 1.0% → almacenar como fracción
    # Valores reales (no son porcentajes): "CSC Parts Revenue CSC Parts Revenue\n1.2\n1.0"
    m = re.search(r"CSC Parts Revenue CSC Parts Revenue\s*\n([\d.]+)\s*\n([\d.]+)", p1)
    if m:
        result["proactive_connections_r12"]  = float(m.group(1))
        result["meaningful_connections_r12"] = float(m.group(2))
    result["proactive_connections_obj"]  = 0.10  # objetivo: 10%
    result["meaningful_connections_obj"] = 0.06  # objetivo: 6%

    # ── TC Metrics desde página 1 (One Page view, misma vista que el usuario ve) ──
    # Formato real: "32.24%\n1\n152 49 63.78" — el \n1\n es ruido del PDF
    m = re.search(
        r"Total Cases Cases TC Resolved % TC Resolved.*?([\d.]+)%\s*\n\d+\s*\n(\d+)\s+(\d+)",
        p1, re.DOTALL
    )
    if m:
        result["tc_resolved_pct"]  = float(m.group(1)) / 100
        result["tc_total_cases"]   = int(m.group(2))
        result["tc_resolved"]      = int(m.group(3))

    # Gauges: "Feedback\n24.49% 99.15%\nA&T Crop Care 88.19%"
    # grupo1=% Closed <2h, grupo2=% Well Prepared, grupo3=% Closed w/ Feedback
    m = re.search(
        r"Feedback\s*\n([\d.]+)%\s+([\d.]+)%\s*\n[^\n]*?(\d+\.\d+)%",
        p1
    )
    if m:
        result["tc_closed_lt2h_pct"]     = float(m.group(1)) / 100
        result["tc_well_prepared_pct"]   = float(m.group(2)) / 100
        result["tc_closed_feedback_pct"] = float(m.group(3)) / 100
    # Objetivos fijos TC
    result["tc_resolved_obj"]        = 0.35
    result["tc_closed_lt2h_obj"]     = 0.70
    result["tc_well_prepared_obj"]   = 0.95
    result["tc_closed_feedback_obj"] = 0.80

    return result


# ── 8. Objetivos PAN (PAN Ceres Tolvas FY26.xlsx, hoja Objetivos) ────────────
def leer_pan_objetivos() -> dict:
    path = BASE / "PAN Ceres Tolvas FY26.xlsx"
    if not path.exists():
        print("  ! PAN Ceres Tolvas FY26.xlsx no encontrado")
        return {}

    try:
        df = pd.read_excel(path, sheet_name="Objetivos", header=None)
    except Exception as e:
        print(f"  ! Error leyendo PAN objetivos: {e}")
        return {}

    # col index → clave interna; col 11 = "OBJETIVO TOTAL sucursales"
    COL_MAP = {5: "Tandil", 6: "Neco", 7: "Tresa", 8: "Juárez", 9: "Olava", 11: "Total"}
    # row index (header=None, todos los índices +1 respecto a lectura con header)
    ROW_MAP = {
        3:  "hec_conectadas",
        5:  "hec_altamente",
        7:  "maquinas",
        9:  "ecosistemas",
        18: "paquete_esencial",
        21: "licencias",
    }
    INT_FIELDS = {"maquinas", "ecosistemas", "paquete_esencial", "licencias"}

    result = {suc: {} for suc in COL_MAP.values()}

    for row_idx, campo in ROW_MAP.items():
        if row_idx >= len(df):
            continue
        row = df.iloc[row_idx]
        for col_idx, suc in COL_MAP.items():
            if col_idx >= len(row):
                continue
            val = row.iloc[col_idx]
            try:
                if not pd.notna(val):
                    result[suc][campo] = None
                    continue
                fval = float(val)
                if campo in INT_FIELDS:
                    result[suc][campo] = int(round(fval, 0))
                else:
                    result[suc][campo] = round(fval, 2)
            except (TypeError, ValueError):
                result[suc][campo] = None

    return result


# ── 9b. Códigos de licencias y precios ───────────────────────────────────────
def leer_codigos() -> list:
    path = BASE / "Precios Licencias, Señales y Paquetes de Monitoreo (1).xlsx"
    if not path.exists():
        return []
    try:
        import openpyxl as _opxl
        wb = _opxl.load_workbook(path, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = [r for r in ws.iter_rows(values_only=True)]
    except Exception:
        return []

    sections = []
    current = None
    i = 0
    while i < len(rows):
        row = rows[i]
        vals = [v for v in row if v is not None]
        if not vals:
            i += 1
            continue
        col0 = str(row[0]).strip() if row[0] else ""
        # Fila de título de sección principal (col A tiene texto, resto vacío)
        rest_empty = all(row[j] is None for j in range(1, len(row)))
        if rest_empty and col0 and col0 not in ("Código", "Codigo"):
            current = {"titulo": col0, "columnas": [], "filas": []}
            sections.append(current)
            i += 1
            continue
        # Fila de encabezados de columna
        if col0 in ("Código", "Codigo") and current is not None:
            current["columnas"] = [str(v).strip() if v else "" for v in row if v is not None]
            i += 1
            continue
        # Fila de datos
        if current is not None and col0:
            fila = []
            for v in row:
                if v is None:
                    fila.append("")
                elif isinstance(v, float) and v == int(v):
                    fila.append(int(v))
                else:
                    fila.append(str(v).strip())
            # Recortar al ancho de columnas definidas
            n = len(current["columnas"]) if current["columnas"] else len(fila)
            fila = fila[:n]
            current["filas"].append(fila)
        i += 1
    return sections


# ── 9. Facturación y Órdenes de Reparación ───────────────────────────────────
def leer_facturacion() -> dict:
    import glob as _glob
    paths = sorted(
        BASE.glob("Facturacion y ordenes de reparacion*.xlsx"),
        key=lambda p: p.stat().st_mtime, reverse=True
    )
    if not paths:
        return {}
    path = paths[0]

    import openpyxl as _opxl
    wb = _opxl.load_workbook(path, data_only=True)
    if "Hoja1" not in wb.sheetnames:
        return {}
    ws = wb["Hoja1"]

    def _v(row, col):
        try:
            val = ws.cell(row=row, column=col).value
            return round(float(val), 2) if val is not None else 0.0
        except (TypeError, ValueError):
            return 0.0

    def _i(row, col):
        try:
            val = ws.cell(row=row, column=col).value
            return int(val) if val is not None else 0
        except (TypeError, ValueError):
            return 0

    return {
        "archivo": path.name,
        "taller_isg": {
            "cantidad":        _i(3, 4),
            "monto_repuesto":  _v(3, 5),
            "monto_mo":        _v(3, 6),
            "total":           _v(3, 7),
        },
        "ordenes_isg": {
            "cantidad":        _i(4, 4),
            "monto_repuesto":  _v(4, 5),
            "monto_mo":        _v(4, 6),
            "total":           _v(4, 7),
        },
        "ju": {
            "cantidad":        _i(5, 4),
            "monto_repuesto":  _v(5, 5),
            "monto_mo":        _v(5, 6),
            "total":           _v(5, 7),
        },
        "js10": {
            "cantidad":        _i(6, 4),
            "monto_repuesto":  _v(6, 5),
            "monto_mo":        _v(6, 6),
            "total":           _v(6, 7),
        },
        "licencias_jp": {
            "cantidad":        _i(7, 4),
            "monto_repuesto":  _v(7, 5),
            "monto_mo":        _v(7, 6),
            "total":           _v(7, 7),
        },
        "nota_ventas": {
            "cantidad":        _i(8, 4),
            "monto_mo":        _v(8, 6),
        },
        "total": {
            "cantidad":        _i(9, 4),
            "monto_repuesto":  _v(9, 5),
            "monto_mo":        _v(9, 6),
            "total":           _v(9, 7),
        },
    }


# ── Main ─────────────────────────────────────────────────────────────────────
def main():
    print("Leyendo KPIs ISG...")
    kpis = leer_kpis_isg()
    print("Leyendo paquetes de monitoreo...")
    paquetes = leer_paquetes()
    print("Leyendo bajas de monitoreo...")
    bajas = leer_bajas()
    print("Leyendo activaciones...")
    activaciones = leer_activaciones()
    print("Leyendo Precision Ag PDF...")
    precision_ag = leer_precision_ag()
    print("Leyendo Sol.PUKs...")
    ams = leer_ams()
    print("Leyendo conectividad (hectáreas + máquinas)...")
    conectividad = leer_conectividad()
    print("Leyendo CSC One Page Report...")
    csc = leer_csc()
    print("Leyendo Facturación y Órdenes de Reparación...")
    facturacion = leer_facturacion()
    print("Leyendo objetivos PAN...")
    pan_obj = leer_pan_objetivos()
    print("Leyendo códigos de licencias...")
    codigos = leer_codigos()

    # ── Sobreescribir "Paquete esencial" en kpis_sucursal con datos de AMS ──────
    for _suc, _real in (ams.get("paquete_esencial") or {}).items():
        if _suc in kpis and "Paquete esencial" in kpis[_suc]:
            _obj = kpis[_suc]["Paquete esencial"].get("objetivo") or 0
            kpis[_suc]["Paquete esencial"]["real"] = _real
            kpis[_suc]["Paquete esencial"]["cumplimiento"] = round(min(_real / _obj, 1.0), 4) if _obj else 0

    # ── Expert Connect / Adopcion / Utilizacion desde PABLO.ECOSISTEMAS OBJETIVOS.xlsx ──
    expert_connect = {}
    adopcion = []
    utilizacion = []
    eco_path_ec = BASE / "PABLO.ECOSISTEMAS OBJETIVOS.xlsx"
    if eco_path_ec.exists():
        try:
            import openpyxl as _opxl_ec
            _wb_ec = _opxl_ec.load_workbook(eco_path_ec, data_only=True)
            if "EXPERT CONNECT" in _wb_ec.sheetnames:
                _ws_ec = _wb_ec["EXPERT CONNECT"]
                for _row in _ws_ec.iter_rows(min_row=3, values_only=True):
                    _lbl = str(_row[1]).strip() if _row[1] else None
                    if not _lbl or _lbl == "None":
                        continue
                    _obj  = float(_row[2]) if _row[2] is not None else None
                    _fytd = float(_row[3]) if _row[3] is not None else None
                    expert_connect[_lbl] = {"objetivo": _obj, "fytd": _fytd}
            if "ADOPCIÓN" in _wb_ec.sheetnames:
                _ws_a = _wb_ec["ADOPCIÓN"]
                for _row in _ws_a.iter_rows(min_row=2, values_only=True):
                    _lbl = str(_row[1]).strip() if _row[1] else None
                    if not _lbl or _lbl == "None":
                        continue
                    _obj  = float(_row[2]) if _row[2] is not None else None
                    _fytd = float(_row[3]) if _row[3] is not None else None
                    adopcion.append({"label": _lbl, "objetivo": _obj, "fytd": _fytd})
            if "UTILIZACION" in _wb_ec.sheetnames:
                _ws_u = _wb_ec["UTILIZACION"]
                for _row in _ws_u.iter_rows(min_row=2, values_only=True):
                    _lbl = str(_row[1]).strip() if _row[1] else None
                    if not _lbl or _lbl == "None":
                        continue
                    _obj  = float(_row[2]) if _row[2] is not None else None
                    # col E (index 4) es el FYTD; puede venir como "65,3%" o float
                    _raw_fytd = _row[4] if len(_row) > 4 else None
                    if isinstance(_raw_fytd, str):
                        _raw_fytd = _raw_fytd.strip().rstrip('%').replace(',', '.')
                        try:
                            _fytd = float(_raw_fytd) / 100
                        except ValueError:
                            _fytd = None
                    else:
                        _fytd = float(_raw_fytd) if _raw_fytd is not None else None
                    utilizacion.append({"label": _lbl, "objetivo": _obj, "fytd": _fytd})
        except Exception:
            pass

    data = {
        "generado":       datetime.now().strftime("%Y-%m-%d %H:%M"),
        "kpis_sucursal":  kpis,
        "paquetes":       paquetes,
        "bajas":          bajas,
        "activaciones":   activaciones,
        "precision_ag":   precision_ag,
        "ams":            ams,
        "conectividad":   conectividad,
        "csc":            csc,
        "expert_connect": expert_connect,
        "adopcion":       adopcion,
        "utilizacion":    utilizacion,
        "facturacion":    facturacion,
        "pan_objetivos":  pan_obj,
        "codigos":        codigos,
    }

    root = Path(__file__).parent
    (root / "data.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    (root / "data.js").write_text(
        "const DATA = " + json.dumps(data, ensure_ascii=False) + ";", encoding="utf-8")
    print(f"OK archivos generados: data.json y data.js")

    # ── HTML standalone: inyecta data.js dentro del HTML → un solo archivo ──
    html_src  = root / "dashboard.html"
    html_dest = root / "Dash_CSC.html"
    if html_src.exists():
        data_js_inline = "const DATA = " + json.dumps(data, ensure_ascii=False) + ";"
        html_content   = html_src.read_text(encoding="utf-8")
        # Reemplaza el <script src="data.js"> por el bloque de datos inline
        standalone = html_content.replace(
            '<script src="data.js"></script>',
            f'<script>\n{data_js_inline}\n</script>'
        )
        html_dest.write_text(standalone, encoding="utf-8")
        (root / "index.html").write_text(standalone, encoding="utf-8")
        print(f"OK archivo standalone generado: {html_dest.name}")


if __name__ == "__main__":
    main()
