#!/usr/bin/env python3
"""Importa de forma segura la rutina Excel a Hevy Pro.

Por defecto solo hace preview. Usa --apply para crear las rutinas.
La API key se solicita en Terminal y no se guarda.
"""
from __future__ import annotations

import argparse
import difflib
import getpass
import json
import os
import re
import sys
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
import uuid
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

BASE_URL = "https://api.hevyapp.com"
ROUTINE_SHEETS = ["Superior A", "Inferior A", "Superior B", "Inferior B", "Brazos Casa"]

ALIASES = {
    "Press inclinado mancuernas": ["Incline Bench Press (Dumbbell)"],
    "Remo T-Bar con apoyo": ["Chest Supported T Bar Row", "T Bar Row"],
    "Chest Press máquina": ["Chest Press (Machine)"],
    "Jalón al pecho polea": ["Lat Pulldown (Cable)"],
    "Elevación lateral unilateral polea": ["Single Arm Lateral Raise (Cable)"],
    "Pushdown tríceps": ["Triceps Pushdown", "Triceps Rope Pushdown"],
    "Curl inclinado mancuernas": ["Seated Incline Curl (Dumbbell)", "Incline Curl (Dumbbell)"],
    "Sentadilla barra": ["Squat (Barbell)"],
    "Prensa": ["Leg Press (Machine)"],
    "Curl femoral acostado": ["Lying Leg Curl (Machine)"],
    "Hip Thrust": ["Hip Thrust (Barbell)", "Hip Thrust (Machine)"],
    "Pantorrilla Smith": ["Standing Calf Raise (Smith)"],
    "Press plano mancuernas": ["Bench Press (Dumbbell)"],
    "Remo sentado agarre ancho": ["Seated Cable Row - Bar Wide Grip"],
    "Press hombro mancuernas": ["Shoulder Press (Dumbbell)", "Overhead Press (Dumbbell)"],
    "Fly de pecho máquina": ["Chest Fly (Machine)", "Butterfly (Pec Deck)"],
    "Jalón al pecho / agarre alternativo": ["Lat Pulldown - Close Grip (Cable)", "Lat Pulldown (Cable)"],
    "Cruce inverso en polea": ["Reverse Fly (Cable)", "Rear Delt Cable Fly", "Face Pull"],
    "Extensión tríceps overhead polea": ["Overhead Triceps Extension (Cable)"],
    "Curl martillo": ["Hammer Curl (Dumbbell)"],
    "Peso muerto rumano": ["Romanian Deadlift (Barbell)"],
    "Sentadilla búlgara mancuernas": ["Bulgarian Split Squat (Dumbbell)"],
    "Extensión cuádriceps": ["Leg Extension (Machine)"],
    "Abducción máquina": ["Hip Abduction (Machine)"],
    "Aducción máquina": ["Hip Adduction (Machine)"],
    "Curl barra Z": ["EZ Bar Biceps Curl"],
    "Pushdown cuerda": ["Triceps Rope Pushdown"],
    "Extensión overhead polea": ["Overhead Triceps Extension (Cable)"],
    "Press francés barra Z": ["Skullcrusher (EZ Bar)", "EZ Bar Skullcrusher", "Skullcrusher (Barbell)"],
}


def normalize(value: str) -> str:
    value = unicodedata.normalize("NFKD", str(value))
    value = "".join(c for c in value if not unicodedata.combining(c)).lower()
    return " ".join(re.sub(r"[^a-z0-9]+", " ", value).split())


def col_index(ref: str) -> int:
    letters = re.match(r"([A-Z]+)", ref).group(1)
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


def read_xlsx_rows(path: Path) -> dict[str, list[list]]:
    ns = {
        "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    with zipfile.ZipFile(path) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rel_map = {r.attrib["Id"]: r.attrib["Target"] for r in rels}
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall("x:si", ns):
                shared.append("".join((t.text or "") for t in si.iterfind(".//x:t", ns)))
        result = {}
        for sh in wb.findall("x:sheets/x:sheet", ns):
            name = sh.attrib["name"]
            rid = sh.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = rel_map[rid]
            xml_path = ("xl/" + target.lstrip("/")).replace("xl/xl/", "xl/")
            root = ET.fromstring(z.read(xml_path))
            rows = []
            for row in root.findall("x:sheetData/x:row", ns):
                vals, max_col = {}, -1
                for cell in row.findall("x:c", ns):
                    idx = col_index(cell.attrib.get("r", "A1"))
                    max_col = max(max_col, idx)
                    typ = cell.attrib.get("t")
                    v = cell.find("x:v", ns)
                    val = None if v is None else (v.text or "")
                    if val is not None and typ == "s":
                        val = shared[int(val)]
                    elif val is not None and typ == "n":
                        try:
                            num = float(val)
                            val = int(num) if num.is_integer() else num
                        except ValueError:
                            pass
                    vals[idx] = val
                if max_col >= 0:
                    arr = [None] * (max_col + 1)
                    for idx, val in vals.items():
                        arr[idx] = val
                    rows.append(arr)
            result[name] = rows
        return result


def parse_plan(excel: Path) -> dict[str, list[dict]]:
    sheets = read_xlsx_rows(excel)
    plan = {}
    for sheet in ROUTINE_SHEETS:
        rows = sheets.get(sheet)
        if not rows:
            raise ValueError(f"No se encontró la hoja '{sheet}'.")
        header = next((i for i, r in enumerate(rows) if r and r[0] == "Ejercicio"), None)
        if header is None:
            raise ValueError(f"No se encontró la tabla en '{sheet}'.")
        items = []
        for row in rows[header + 1:]:
            row = list(row) + [None] * 8
            try:
                series = int(row[1])
            except (TypeError, ValueError):
                continue
            items.append({
                "exercise": str(row[0]), "sets": series, "reps": str(row[2]),
                "load": str(row[3]), "rir": str(row[4]), "rpe": str(row[5]),
                "rest": str(row[6]), "progression": str(row[7]),
            })
        plan[sheet] = items
    return plan


def parse_rep_range(text: str) -> tuple[int, int]:
    nums = [int(x) for x in re.findall(r"\d+", text)]
    if not nums:
        raise ValueError(f"Repeticiones inválidas: {text}")
    return (nums[0], nums[0]) if len(nums) == 1 else (nums[0], nums[1])


def parse_weight(text: str):
    if not text or "calibr" in text.lower():
        return None
    m = re.search(r"(\d+(?:[.,]\d+)?)", text)
    return float(m.group(1).replace(",", ".")) if m else None


def parse_rest(text: str):
    values = []
    for part in re.split(r"[–—-]", text or ""):
        part = part.strip()
        if not re.search(r"\d", part):
            continue
        try:
            if ":" in part:
                m, s = part.split(":", 1)
                values.append(int(m) * 60 + int(s))
            else:
                values.append(int(float(part)))
        except ValueError:
            pass
    return max(values) if values else None


class HevyAPI:
    def __init__(self, api_key: str): self.api_key = api_key

    def request(self, method: str, path: str, payload=None):
        headers = {"api-key": self.api_key, "Accept": "application/json", "User-Agent": "api-hevypro/2.0"}
        data = None
        if payload is not None:
            data = json.dumps(payload, ensure_ascii=False).encode()
            headers["Content-Type"] = "application/json"
        req = urllib.request.Request(BASE_URL + path, data=data, headers=headers, method=method)
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                raw = resp.read().decode()
                return json.loads(raw) if raw else None
        except urllib.error.HTTPError as e:
            raw = e.read().decode(errors="replace")
            raise RuntimeError(f"HTTP {e.code} {method} {path}: {raw}") from None

    def get(self, path): return self.request("GET", path)
    def post(self, path, payload): return self.request("POST", path, payload)

    def list_all(self, endpoint: str, key: str, page_size=100):
        page, out = 1, []
        while True:
            q = urllib.parse.urlencode({"page": page, "pageSize": page_size})
            body = self.get(f"{endpoint}?{q}")
            out.extend(body.get(key, []))
            if page >= int(body.get("page_count", page)):
                return out
            page += 1


def clean_api_key(raw: str) -> str:
    key = (raw or "").strip()
    if key.lower().startswith("api-key:"):
        key = key.split(":", 1)[1].strip()
    if len(key) >= 2 and key[0] == key[-1] and key[0] in {'"', "'"}:
        key = key[1:-1].strip()
    try:
        return str(uuid.UUID(key))
    except ValueError:
        raise ValueError("La HEVY_API_KEY no tiene formato UUID válido.") from None


def score(query: str, title: str) -> float:
    return difflib.SequenceMatcher(None, normalize(query), normalize(title)).ratio()


def resolve_templates(plan, templates, mapping_file: Path):
    saved = {}
    if mapping_file.exists():
        try: saved = json.loads(mapping_file.read_text())
        except Exception: pass
    by_id = {t["id"]: t for t in templates}
    exact = {}
    for t in templates:
        exact.setdefault(normalize(t["title"]), []).append(t)
    required = []
    for items in plan.values():
        for item in items:
            if item["exercise"] not in required:
                required.append(item["exercise"])
    resolved = {}
    for ex in required:
        if saved.get(ex) in by_id:
            resolved[ex] = by_id[saved[ex]]
            continue
        found = None
        for alias in ALIASES.get(ex, []):
            matches = exact.get(normalize(alias), [])
            if len(matches) == 1:
                found = matches[0]; break
        if found is None:
            candidates = sorted(
                templates,
                key=lambda t: max(score(q, t["title"]) for q in [ex] + ALIASES.get(ex, [])),
                reverse=True,
            )[:8]
            print(f"\nSelecciona ejercicio para: {ex}")
            for i, c in enumerate(candidates, 1):
                print(f"  {i}. {c['title']} [{c.get('equipment', '')}]")
            while True:
                try:
                    idx = int(input("Opción: ")) - 1
                    found = candidates[idx]
                    break
                except (ValueError, IndexError):
                    print("Opción inválida")
        resolved[ex] = found
        saved[ex] = found["id"]
    mapping_file.write_text(json.dumps(saved, ensure_ascii=False, indent=2))
    return resolved


def build_payload(title, items, resolved):
    exercises = []
    for item in items:
        rep_start, _ = parse_rep_range(item["reps"])
        notes = (
            f"Rango objetivo: {item['reps']} | RIR: {item['rir']} | RPE: {item['rpe']} | "
            f"Carga inicial: {item['load']} | Progresión: {item['progression']}"
        )
        exercises.append({
            "exercise_template_id": resolved[item["exercise"]]["id"],
            "superset_id": None,
            "rest_seconds": parse_rest(item["rest"]),
            "notes": notes,
            "sets": [{"type": "normal", "weight_kg": parse_weight(item["load"]), "reps": rep_start}
                     for _ in range(item["sets"])],
        })
    return {"routine": {"title": title, "folder_id": None,
                         "notes": "Rutina 6–8 semanas importada desde Excel.",
                         "exercises": exercises}}


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--excel", default="Rutina_Entrenamiento_5_Dias_Hevy.xlsx")
    p.add_argument("--mapping", default="hevy_mapping.json")
    p.add_argument("--apply", action="store_true")
    args = p.parse_args()

    excel = Path(args.excel)
    if not excel.exists(): sys.exit(f"No encuentro el Excel: {excel}")
    plan = parse_plan(excel)
    print("Rutinas encontradas:", ", ".join(plan))

    raw = os.getenv("HEVY_API_KEY") or getpass.getpass("Pega tu HEVY_API_KEY: ")
    try: api_key = clean_api_key(raw)
    except ValueError as e: sys.exit(f"ERROR API KEY: {e}")

    api = HevyAPI(api_key)
    api.get("/v1/exercise_templates?page=1&pageSize=1")
    print("Conexión con Hevy validada.")
    templates = api.list_all("/v1/exercise_templates", "exercise_templates")
    print(f"Catálogo cargado: {len(templates)} ejercicios")

    resolved = resolve_templates(plan, templates, Path(args.mapping))
    print("\n=== PREVIEW ===")
    for name, items in plan.items():
        print(f"\n{name}")
        for item in items:
            print(f"- {item['exercise']} -> {resolved[item['exercise']]['title']} | {item['sets']} x {item['reps']} | {item['load']}")

    if not args.apply:
        print("\nMODO PREVIEW: no se creó nada. Si está correcto, ejecuta con --apply")
        return

    existing = {r.get("title") for r in api.list_all("/v1/routines", "routines", 10)}
    for name, items in plan.items():
        if name in existing:
            print(f"SKIP: {name} ya existe")
            continue
        api.post("/v1/routines", build_payload(name, items, resolved))
        print(f"CREADA: {name}")


if __name__ == "__main__":
    main()
