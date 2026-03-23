"""
Convierte horarios_uni_completo.xlsx → actualiza FIA_DATA en index.html
Uso: python build_data.py
"""
import json
import re
import sys
import openpyxl

XLSX_PATH = "horarios_uni_completo.xlsx"
HTML_PATH = "index.html"

# Mapeo sección → especialidad (actualizar si cambia en cada periodo)
SECC_ESP = {
    "E": "IS",
    "F": "IH",
    "G": "IA",
    "H": "IS",
    "I": "CB",
    "J": "IA",
}

TIPO_MAP = {
    "TEORIA": "T", "TEORÍA": "T",
    "PRACTICA": "P", "PRÁCTICA": "P",
    "LABORATORIO": "L", "LAB": "L",
    "SEMINARIO": "S",
}


def parse_horario(h):
    """'MA 10-12' → ('MA', 10, 12)"""
    parts = str(h).strip().split()
    if len(parts) != 2:
        return None
    rng = parts[1].split("-")
    if len(rng) != 2:
        return None
    try:
        return parts[0].upper(), int(rng[0]), int(rng[1])
    except ValueError:
        return None


def get_ciclo(cod):
    c = str(cod)[2] if len(str(cod)) > 2 else ""
    return int(c) if c.isdigit() else 11


def load_rows():
    wb = openpyxl.load_workbook(XLSX_PATH, read_only=True, data_only=True)
    rows = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        all_rows = list(ws.iter_rows(values_only=True))

        # Buscar fila de encabezado
        hi = -1
        headers = []
        for i, row in enumerate(all_rows):
            cells = [str(c).upper().strip() if c is not None else "" for c in row]
            if "COD" in cells and "CURSO" in cells:
                hi = i
                headers = cells
                break
        if hi < 0:
            continue

        def ix(name):
            try:
                return headers.index(name)
            except ValueError:
                return -1

        iC = ix("COD")
        iCu = ix("CURSO")
        iS = ix("SECC")
        iTi = ix("TIPO")
        iH = ix("HORARIO")
        iD = ix("DIA")
        iHi = ix("H INI")
        iHf = ix("H FIN")
        iSa = ix("AULA") if ix("AULA") >= 0 else ix("SALON")
        iDo = ix("DOCENTE")
        iCy = ix("CICLO")
        fia_fmt = iH >= 0

        for row in all_rows[hi + 1:]:
            if not row[iC] or not row[iCu]:
                continue

            cod = str(row[iC]).strip()
            curso = str(row[iCu]).strip()
            secc = str(row[iS] or "").strip().upper() if iS >= 0 else ""
            tipo_raw = str(row[iTi] or "").strip().upper() if iTi >= 0 else ""
            tipo = TIPO_MAP.get(tipo_raw, "T")
            ciclo = int(row[iCy]) if iCy >= 0 and row[iCy] and str(row[iCy]).strip().isdigit() else get_ciclo(cod)
            salon = str(row[iSa] or "").strip() if iSa >= 0 else ""
            docente = str(row[iDo] or "").strip() if iDo >= 0 else ""
            esp = SECC_ESP.get(secc, "")

            if fia_fmt:
                ph = parse_horario(row[iH])
                if not ph:
                    continue
                dia, h0, h1 = ph
                rows.append({"esp": esp, "cod": cod, "secc": secc, "curso": curso,
                              "docente": docente, "tipo": tipo, "dia": dia,
                              "hIni": h0, "hFin": h1, "salon": salon})
            else:
                if iD < 0 or not row[iD]:
                    continue
                try:
                    h0, h1 = int(row[iHi]), int(row[iHf])
                except (TypeError, ValueError):
                    continue
                dia = str(row[iD]).strip().upper()
                rows.append({"esp": esp, "cod": cod, "secc": secc, "curso": curso,
                              "docente": docente, "tipo": tipo, "dia": dia,
                              "hIni": h0, "hFin": h1, "salon": salon})

    wb.close()
    return rows


def update_html(rows):
    with open(HTML_PATH, "r", encoding="utf-8") as f:
        content = f.read()

    new_data = "const FIA_DATA=" + json.dumps(rows, ensure_ascii=False, separators=(",", ":")) + ";"
    new_content, n = re.subn(r"const FIA_DATA=\[.*?\];", new_data, content)

    if n == 0:
        print("ERROR: no se encontró 'const FIA_DATA=[...]' en index.html", file=sys.stderr)
        sys.exit(1)

    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(new_content)

    print(f"OK: {len(rows)} sesiones escritas en FIA_DATA")


if __name__ == "__main__":
    rows = load_rows()
    if not rows:
        print("ERROR: no se encontraron filas en el xlsx", file=sys.stderr)
        sys.exit(1)
    update_html(rows)
