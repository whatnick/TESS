"""Package upstream TESS Rev 00 exports for an Elecrow RFQ, not production."""

from __future__ import annotations

import argparse
import csv
import hashlib
import io
import json
import math
import re
import subprocess
import zipfile
from collections import Counter
from pathlib import Path
from urllib.parse import quote

import openpyxl
import pymupdf
import xlrd

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "TESS_00" / "OUTPUT"
ODB = SOURCE / "ODB" / "odb"
PCB = ODB / "steps" / "pcb"
UPSTREAM_REVISION = "ee5d5cd45d179993bc479158fda230e5350106f7"
EXCLUDED_HEADERS = {"P2", "P3", "P6", "P7"}
BOM_COLUMNS = [
    "*Reference", "*Manufacturer Part number", "*QTY",
    "*Original (Y/N)", "Description", "Purchase Link",
]
COPPER_AND_MASK = {
    "GTL": "TESS-rev00-L1-top.gbr",
    "G1": "TESS-rev00-L2-inner.gbr",
    "G2": "TESS-rev00-L3-inner.gbr",
    "GBL": "TESS-rev00-L4-bottom.gbr",
    "GTS": "TESS-rev00-top-mask.gbr",
    "GBS": "TESS-rev00-bottom-mask.gbr",
    "GTO": "TESS-rev00-top-silk.gbr",
    "GBO": "TESS-rev00-bottom-silk.gbr",
}
PASTE = {"GTP": "TESS-rev00-top-paste.gbr", "GBP": "TESS-rev00-bottom-paste.gbr"}
HOLDS = [
    "RFQ ONLY: no manufacturing, procurement, payment or substitutions authorised.",
    "Native Altium schematic/PCB sources and ERC/DRC reports are absent. "
    "No ERC, DRC, schematic parity or supplier DFM pass is claimed.",
    "Four copper layers; preserve upstream layer order and routing. "
    "ODB board_thickness=0.0292566 inch (0.74311764 mm). Confirm actual "
    "stackup, finished thickness and copper weights with Phyx and Elecrow. "
    "Do not silently substitute 1.6 mm or a standard stackup.",
    "Castellated plated holes lie on both long edges. Preserve them; "
    "supplier must confirm castellated-hole capability and finishing.",
    "Confirm finish, mask, silk, panelisation and stencil design before production. "
    "The PDF's rendered colours are not fabrication specifications.",
    "U2 storage: C2875854 and listed alternate C5349318 had zero JLC stock "
    "during 2026-11-08 sourcing. Supplier must source and quote exact compatible "
    "SD-interface storage; do not replace with generic SPI NAND.",
    "L1/L2: BOM value says 2.2 uH, but specified C97003 / "
    "WPN201610H1R0MTY01 is 1 uH. Quote the named MPN; resolve intended "
    "inductance with the designer before manufacture.",
    "P1: use exact Wurth 620102131822. Upstream alternate C2888424 "
    "is a switch, not this battery connector; it is excluded from purchasing.",
    "J1: BOM names TYPEC-304S-ACP16 but lists different connector codes; "
    "placement instead lists C720630. Quote the exact named connector and "
    "confirm footprint; do not auto-substitute listed alternatives.",
    "P4 C2998855 had only three pieces in JLC stock during 2026-11-08 "
    "sourcing. Confirm exact top-contact connector availability and fit.",
    "SW1/SW2: use exact Wurth 434153017835; listed C33129656 resolves "
    "to a different manufacturer/model. Any substitution requires approval.",
    "Y2: C7294625 was not found in the pricing service. Quote a 24 MHz "
    "2520 crystal with 9 pF load and 10 ppm specification; confirm exact MPN.",
    "Generic passives, LEDs and crystals have specification entries rather "
    "than invented MPNs. Elecrow must propose exact MPNs for approval. "
    "Interpret upstream '2u2F' as 2.2 uF, not 2.2 F.",
    "117 fitted components per board; P2/P3/P6/P7 are DNP because they are "
    "absent from the upstream BOM. TP references are test features, not parts. "
    "Retain their copper and holes. P1 and P5 remain fitted.",
    "Positions preserve Altium component-centre coordinates and rotations, "
    "including bottom-side values: no mirroring or rotation translation. "
    "Assembler must reconcile orientation, pin 1 and machine centroids.",
    "Request first-article inspection and functional checks before completing "
    "the batch: boot, storage, rails, charge voltage/polarity, Wi-Fi, and "
    "external Feather SPI1 with a connected peripheral.",
]


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def upstream_bytes(path: Path) -> bytes:
    relative = path.relative_to(ROOT).as_posix()
    data = subprocess.check_output(["git", "show", f"{UPSTREAM_REVISION}:{relative}"], cwd=ROOT)
    local = path.read_bytes()
    if path.suffix.lower() in {".xls", ".pdf"}:
        matches = local == data
    else:
        matches = local.replace(b"\r\n", b"\n") == data.replace(b"\r\n", b"\n")
    if not matches:
        raise ValueError(f"Source differs from pinned upstream revision: {relative}")
    return data


def read_bom() -> list[dict]:
    path = SOURCE / "BOM" / "BoM Default-TESS_00(Default).xls"
    sheet = xlrd.open_workbook(str(path)).sheet_by_index(0)
    header = next(i for i in range(sheet.nrows) if sheet.cell_value(i, 0) == "Footprint")
    columns = sheet.row_values(header)
    rows = []
    for i in range(header + 1, sheet.nrows):
        values = sheet.row_values(i)
        if not values[3]:
            continue
        row = dict(zip(columns, values))
        row["refs"] = [r.strip() for r in row["Designator"].split(",")]
        qty = row["Quantity"]
        if int(qty) != qty or int(qty) != len(row["refs"]):
            raise ValueError(f"BOM reference/quantity mismatch: {row['Designator']}")
        row["Quantity"] = int(qty)
        rows.append(row)
    refs = [r for row in rows for r in row["refs"]]
    if len(refs) != len(set(refs)):
        raise ValueError("Duplicate BOM reference")
    return rows


def read_positions() -> list[dict]:
    path = SOURCE / "Pick Place" / "Pick Place for TESS_00.csv"
    text = path.read_text(encoding="utf-8-sig")
    # Omit the upstream absolute workstation path and report preamble.
    rows = list(csv.DictReader(io.StringIO(text[text.index('"Designator"'): ])))
    refs = [row["Designator"] for row in rows]
    if len(refs) != len(set(refs)):
        raise ValueError("Duplicate placement reference")
    return rows


def fitted_positions(bom: list[dict], positions: list[dict]) -> tuple[list[dict], list[dict]]:
    refs = {r for row in bom for r in row["refs"]}
    available = {row["Designator"] for row in positions}
    if refs - available:
        raise ValueError(f"Missing placements: {sorted(refs - available)}")
    excluded = [r for r in positions if r["Designator"] not in refs]
    unexpected = [
        r["Designator"] for r in excluded
        if r["Designator"] not in EXCLUDED_HEADERS and not re.fullmatch(r"TP\d+", r["Designator"])
    ]
    if unexpected:
        raise ValueError(f"Unclassified non-BOM placements: {unexpected}")
    return [row for row in positions if row["Designator"] in refs], excluded


def odb_drills(path: Path) -> dict[float, list[tuple[float, float]]]:
    text = path.read_text()
    if not text.startswith("UNITS=INCH"):
        raise ValueError("Unsupported ODB drill units")
    symbols = {
        int(i): float(d) * 0.0254
        for i, d in re.findall(r"(?m)^\$(\d+) r([\d.]+)$", text)
    }
    groups: dict[float, list[tuple[float, float]]] = {}
    for line in text.splitlines():
        if not line.startswith("P "):
            continue
        fields = line.split()
        if fields[4:7] != ["P", "0", "0"]:
            raise ValueError(f"Unsupported drill feature: {line}")
        diameter = symbols[int(fields[3])]
        groups.setdefault(diameter, []).append((float(fields[1]) * 25.4, float(fields[2]) * 25.4))
    if not groups:
        raise ValueError("Empty drill layer")
    return groups


def write_drills(path: Path, groups: dict[float, list[tuple[float, float]]], plating: str) -> None:
    tools = sorted(groups)
    lines = ["M48", f";TYPE={plating}", ";Explicit decimal metric coordinates", "METRIC"]
    lines += [f"T{i:02}C{d:.6f}" for i, d in enumerate(tools, 1)]
    lines += ["%", "G90", "G05"]
    for i, d in enumerate(tools, 1):
        lines.append(f"T{i:02}")
        lines.extend(f"X{x:.6f}Y{y:.6f}" for x, y in groups[d])
    path.write_text("\n".join(lines + ["M30", ""]), encoding="ascii")


def original_drills(text: str) -> list[tuple[float, float, float]]:
    if ";FILE_FORMAT=2:5" not in text or "INCH,LZ" not in text:
        raise ValueError("Unexpected upstream Excellon format")
    tools = {
        int(tool): float(diameter) * 25.4
        for tool, diameter in re.findall(r"(?m)^T(\d+)F.*?C([\d.]+)$", text)
    }
    current_tool = None
    x = y = None
    points = []
    for line in text.splitlines():
        if re.fullmatch(r"T\d+", line):
            current_tool = int(line[1:])
        elif line.startswith(("X", "Y")):
            for axis, digits in re.findall(r"([XY])(-?\d+)", line):
                # Altium's export includes leading zeros and omits trailing zeros.
                # Corroborate this interpretation against independent ODB points.
                sign = -1 if digits.startswith("-") else 1
                value = sign * int(digits.lstrip("-").ljust(7, "0")) / 100_000 * 25.4
                if axis == "X":
                    x = value
                else:
                    y = value
            if current_tool not in tools or x is None or y is None:
                raise ValueError("Uninitialised Excellon coordinate/tool")
            points.append((tools[current_tool], x, y))
    return points


def outline_segments() -> list[dict]:
    text = (PCB / "layers" / "outline" / "features").read_text()
    if not text.startswith("UNITS=INCH"):
        raise ValueError("Unsupported outline units")
    segments = []
    for line in text.splitlines():
        if not line.startswith(("L ", "A ")):
            continue
        fields = line.split()
        segment = {
            "start": (float(fields[1]) * 25.4, float(fields[2]) * 25.4),
            "end": (float(fields[3]) * 25.4, float(fields[4]) * 25.4),
            "arc": fields[0] == "A",
        }
        if segment["arc"]:
            segment["centre"] = (float(fields[5]) * 25.4, float(fields[6]) * 25.4)
            if fields[-1] not in {"Y", "N"}:
                raise ValueError("Unsupported arc direction")
            segment["clockwise"] = fields[-1] == "Y"
        segments.append(segment)
    if not segments:
        raise ValueError("No ODB outline")
    return segments


def write_outline(path: Path) -> dict:
    remaining = outline_segments()
    current = remaining[0]["start"]
    first = current
    ordered = []
    max_join = 0.0
    while remaining:
        candidates = [
            (math.dist(current, s[endpoint]), i, endpoint)
            for i, s in enumerate(remaining) for endpoint in ("start", "end")
        ]
        distance, index, endpoint = min(candidates)
        if distance > 0.0001:
            raise ValueError("Outline is open or has multiple contours")
        max_join = max(max_join, distance)
        s = dict(remaining.pop(index))
        if endpoint == "end":
            s["start"], s["end"] = s["end"], s["start"]
            if s["arc"]:
                s["clockwise"] = not s["clockwise"]
        s["start"] = current
        ordered.append(s)
        current = s["end"]
    closure = math.dist(current, first)
    if closure > 0.0001:
        raise ValueError("Open outline")
    ordered[-1]["end"] = first
    max_join = max(max_join, closure)

    def coordinate(value: float) -> str:
        return str(round(value * 1_000_000))

    def xy(point: tuple[float, float]) -> str:
        return f"X{coordinate(point[0])}Y{coordinate(point[1])}"

    lines = [
        "G04 Sole outline from upstream ODB analytic line/arc features*",
        "%FSLAX46Y46*%", "%MOMM*%", "%LPD*%", "%ADD10C,0.100000*%",
        "G75*", "D10*", f"{xy(first)}D02*",
    ]
    for s in ordered:
        if s["arc"]:
            i = s["centre"][0] - s["start"][0]
            j = s["centre"][1] - s["start"][1]
            command = "G02" if s["clockwise"] else "G03"
            lines.append(f"{command}{xy(s['end'])}I{coordinate(i)}J{coordinate(j)}D01*")
        else:
            lines.append(f"G01{xy(s['end'])}D01*")
    path.write_text("\n".join(lines + ["M02*", ""]), encoding="ascii")
    points = [s[k] for s in ordered for k in ("start", "end")]
    return {
        "width_mm": round(max(p[0] for p in points) - min(p[0] for p in points), 4),
        "height_mm": round(max(p[1] for p in points) - min(p[1] for p in points), 4),
        "segments": len(ordered),
        "arcs": sum(s["arc"] for s in ordered),
        "max_endpoint_join_normalisation_mm": max_join,
    }


def purchasing(row: dict) -> tuple[str, str, str]:
    ref = row["refs"][0]
    if ref.startswith("R"):
        return str(row["Supplier Part Number 3"]), "N", ""
    if ref.startswith("C"):
        spec = row["Description"].replace("CAP # ", "").replace("2u2F", "2.2uF")
        return spec, "N", ""
    if ref.startswith("D"):
        return "Green LED 0402", "N", ""
    if ref == "Y1":
        return "32.768kHz 2012 crystal 12.5pF", "N", ""
    if ref == "Y2":
        return "24MHz 2520 crystal 9pF 10ppm", "N", ""
    exact = {
        "L1": "WPN201610H1R0MTY01", "L3": "WPN252012H2R2MT",
        "U3": "SY8892ARC", "P5": "62101021021",
    }
    mpn = exact.get(ref, str(row["Value"]).strip())
    if not mpn:
        raise ValueError(f"No purchasing identity for {ref}")
    url = f"https://www.mouser.com/c/?q={quote(mpn, safe='')}"
    if ref in {"U1", "U2", "U3", "U5", "U6", "L1", "L3"}:
        code = row["Supplier Part Number 1"]
        url = f"https://www.lcsc.com/product-detail/{code}.html"
    return mpn, "Y", url


def write_csv(path: Path, fields: list[str], rows: list[list]) -> None:
    with path.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(fields)
        writer.writerows(rows)


def archive(path: Path, files: list[Path]) -> None:
    names = [p.name for p in files]
    if len(names) != len(set(names)):
        raise ValueError("Duplicate archive filename")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for file in sorted(files):
            info = zipfile.ZipInfo(file.name, (2024, 11, 15, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, file.read_bytes())
    with zipfile.ZipFile(path) as z:
        if z.testzip() is not None:
            raise ValueError(f"Archive CRC failure: {path}")


def prepare(boards: int) -> Path:
    if boards <= 0:
        raise ValueError("Board count must be positive")
    output = ROOT / "manufacturing" / "elecrow-rev00"
    output.mkdir(parents=True, exist_ok=True)
    fab = output / "fabrication"
    fab.mkdir(exist_ok=True)
    bom = read_bom()
    fitted, excluded = fitted_positions(bom, read_positions())
    counts = Counter(r["Layer"] for r in fitted)
    if set(counts) - {"TopLayer", "BottomLayer"}:
        raise ValueError("Unknown assembly side")
    for row in fitted:
        x, y = float(row["Center-X(mm)"]), float(row["Center-Y(mm)"])
        if not (0 <= x <= 22.8601 and 0 <= y <= 50.8001):
            raise ValueError(f"Placement outside board: {row['Designator']}")
        if not 0 <= float(row["Rotation"]) < 360:
            raise ValueError(f"Invalid rotation: {row['Designator']}")
    outline = write_outline(fab / "TESS-rev00-outline.gbr")
    source_files = [
        SOURCE / "BOM" / "BoM Default-TESS_00(Default).xls",
        SOURCE / "Pick Place" / "Pick Place for TESS_00.csv",
        SOURCE / "Gerber" / "TESS_00.TXT",
        SOURCE / "TESS_00.PDF", PCB / "profile",
        PCB / "layers" / "outline" / "features", ODB / "misc" / "attrlist",
        ODB / "matrix" / "matrix",
    ]
    fabrications = [fab / "TESS-rev00-outline.gbr"]
    for suffix, filename in COPPER_AND_MASK.items():
        source = SOURCE / "Gerber" / f"TESS_00.{suffix}"
        if "M02*" not in source.read_text():
            raise ValueError(f"Incomplete Gerber: {source}")
        destination = fab / filename
        destination.write_bytes(upstream_bytes(source))
        fabrications.append(destination)
        source_files.append(source)
    drills = {}
    all_odb_drills = []
    edge_holes = 0
    for name, label in [
        ("drill_plated_bottom_layer-top_layer", "PTH"),
        ("drill_non-plated_bottom_layer-top_layer", "NPTH"),
    ]:
        source = PCB / "layers" / name / "features"
        groups = odb_drills(source)
        all_odb_drills.extend((d, x, y) for d, points in groups.items() for x, y in points)
        destination = fab / f"TESS-rev00-{label}.drl"
        write_drills(destination, groups, "PLATED" if label == "PTH" else "NON_PLATED")
        drills[label] = sum(len(points) for points in groups.values())
        for points in groups.values():
            for x, y in points:
                if not (-0.0001 <= x <= 22.8601 and -0.0001 <= y <= 50.8001):
                    raise ValueError("Drill outside source board bounds")
                if label == "PTH" and (abs(x) < 0.0001 or abs(x - 22.86) < 0.0001):
                    edge_holes += 1
        fabrications.append(destination)
        source_files.extend([source, PCB / "layers" / name / "tools"])
    original = original_drills((SOURCE / "Gerber" / "TESS_00.TXT").read_text())
    if len(original) != len(all_odb_drills):
        raise ValueError("Original Excellon and ODB drill counts differ")
    unmatched = list(all_odb_drills)
    for point in original:
        distance, index = min(
            (max(abs(x - y) for x, y in zip(point, other)), i)
            for i, other in enumerate(unmatched)
        )
        if distance > 0.001:
            raise ValueError("Original Excellon and ODB drill geometry differ")
        unmatched.pop(index)

    notes = (
        f"TESS Rev 00 - ELECROW RFQ ONLY - {boards} assemblies\n"
        f"{len(bom)} BOM groups; {len(fitted)} fitted parts per board; "
        f"{len(fitted) * boards} fitted parts for the batch.\n"
        f"Top: {counts['TopLayer']}; bottom: {counts['BottomLayer']} per board.\n"
        f"Outline: {outline['width_mm']} x {outline['height_mm']} mm, four copper layers.\n"
        f"PTH hits: {drills['PTH']}; NPTH hits: {drills['NPTH']}; "
        f"castellated edge PTH hits: {edge_holes}.\n"
        "Copper/mask/silk copied byte-for-byte from upstream. Sole outline "
        "regenerated from ODB analytic arcs/lines, preserving common origin. "
        "Metric absolute PTH/NPTH drill files generated from ODB points; "
        "original combined inch drill and mechanical drawings are NOT duplicate "
        "fab inputs. GTP/GBP go only in the assembly support archive.\n"
        "BOM *QTY is per board; do not multiply both quantity and assembly count.\n"
        "Original=N identifies generic specification entries requiring proposed MPNs; "
        "Original=Y requires the named part (supplier approval holds still apply).\n"
        "Excluded references: " + ", ".join(sorted(r["Designator"] for r in excluded)) + "\n\n"
        + "\n\n".join(f"{i}. {hold}" for i, hold in enumerate(HOLDS, 1)) + "\n"
    )
    note_path = output / "PCBA-NOTES.txt"
    note_path.write_text(notes, encoding="utf-8")
    fab_note = fab / "FABRICATION-NOTES.txt"
    fab_note.write_text(
        notes + "\nNo added order marks on functional surfaces, no artwork changes, "
        "no clipping castellations. Quote a suitable panel; first-article hold applies.\n",
        encoding="utf-8",
    )
    fabrications.append(fab_note)
    bom_rows = []
    for row in bom:
        mpn, original, url = purchasing(row)
        description = f"{row['Description']}; footprint={row['Footprint']}"
        if row["refs"][0] == "L1":
            description += "; HOLD: upstream value 2.2uH conflicts with named 1uH MPN"
        bom_rows.append([row["Designator"], mpn, row["Quantity"], original, description, url])
    bom_csv = output / "TESS-rev00-Elecrow-BOM.csv"
    write_csv(bom_csv, BOM_COLUMNS, bom_rows)
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Elecrow BOM"
    worksheet.append(BOM_COLUMNS)
    for row in bom_rows:
        worksheet.append(row)
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for column, width in {"A": 65, "B": 45, "C": 12, "D": 20, "E": 85, "F": 65}.items():
        worksheet.column_dimensions[column].width = width
    bom_xlsx = output / "TESS-rev00-Elecrow-BOM.xlsx"
    workbook.save(bom_xlsx)
    loaded = openpyxl.load_workbook(bom_xlsx, read_only=True)
    expected = [
        tuple(None if value == "" else value for value in row)
        for row in [BOM_COLUMNS, *bom_rows]
    ]
    if list(loaded.active.values) != expected:
        raise ValueError("Workbook read-back differs from BOM")
    loaded.close()

    placement = output / "TESS-rev00-positions.csv"
    write_csv(
        placement,
        ["Reference", "PosX(mm)", "PosY(mm)", "Side", "Rotation(deg)", "Footprint"],
        [[r["Designator"], r["Center-X(mm)"], r["Center-Y(mm)"],
          "Top" if r["Layer"] == "TopLayer" else "Bottom", r["Rotation"], r["Footprint"]]
         for r in sorted(fitted, key=lambda r: r["Designator"])],
    )
    reconciliation = output / "TESS-rev00-placement-exclusions.csv"
    write_csv(
        reconciliation, ["Reference", "Reason"],
        [[r["Designator"], "DNP header absent from upstream BOM" if r["Designator"] in EXCLUDED_HEADERS
          else "Test feature, not purchased part"] for r in excluded],
    )
    support = [placement, reconciliation, note_path]
    with pymupdf.open(SOURCE / "TESS_00.PDF") as pdf:
        for name, start, end in [
            ("schematics", 0, 3), ("upstream-board-drawings", 4, 9),
            ("upstream-3d-overviews", 10, 11),
        ]:
            destination = output / f"TESS-rev00-{name}.pdf"
            with pymupdf.open() as extracted:
                extracted.insert_pdf(pdf, from_page=start, to_page=end)
                extracted.save(destination)
            support.append(destination)
    for suffix, filename in PASTE.items():
        source = SOURCE / "Gerber" / f"TESS_00.{suffix}"
        destination = output / filename
        destination.write_bytes(upstream_bytes(source))
        support.append(destination)
        source_files.append(source)
    fab_zip = output / "TESS-rev00-Elecrow-Gerbers.zip"
    archive(fab_zip, fabrications)
    placement_zip = output / "TESS-rev00-Elecrow-Assembly.zip"
    archive(placement_zip, support)
    outputs = [bom_csv, bom_xlsx, *fabrications, *support, fab_zip, placement_zip]
    manifest = {
        "status": "RFQ_ONLY_NOT_MANUFACTURING_RELEASED",
        "source_repository": "https://github.com/phyx-be/TESS",
        "upstream_revision": UPSTREAM_REVISION,
        "board_revision": "00", "assembly_quantity": boards,
        "bom_groups": len(bom), "fitted_per_board": len(fitted),
        "fitted_per_batch": len(fitted) * boards, "placements_by_side": dict(counts),
        "outline": outline, "copper_layers": 4, "drill_hits": drills,
        "castellated_edge_holes": edge_holes,
        "source_board_thickness_mm": 0.0292566 * 25.4,
        "validation": {
            "bom_placement_parity": "passed", "workbook_readback": "passed",
            "archive_crc": "passed", "placement_and_drill_bounds": "passed",
            "original_excellon_vs_odb_drill_geometry": "passed within 0.001 mm",
            "native_erc_drc_parity": "not run; native CAD sources absent",
            "supplier_dfm": "not requested", "live_supplier_upload": "not performed",
        },
        "holds": HOLDS,
        "source_hash_basis": "Pinned upstream Git blobs; local checkout newline conversion ignored",
        "source_sha256": {
            str(p.relative_to(ROOT)): hashlib.sha256(upstream_bytes(p)).hexdigest()
            for p in source_files
        },
        "output_sha256": {str(p.relative_to(output)): sha256(p) for p in set(outputs)},
        "archive_member_sha256": {
            z.name: {p.name: sha256(p) for p in files}
            for z, files in [(fab_zip, fabrications), (placement_zip, support)]
        },
    }
    (output / "manifest.json").write_text(json.dumps(manifest, indent=2, sort_keys=True) + "\n")
    print(f"{len(bom)} BOM groups; {len(fitted)} fitted/board; {boards} boards")
    print(f"Gerbers: {fab_zip}\nBOM: {bom_xlsx}\nOther files: {placement_zip}")
    print("RFQ ONLY: resolve PCBA-NOTES.txt holds before manufacturing.")
    return output


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--boards", type=int, default=10)
    prepare(parser.parse_args().boards)
