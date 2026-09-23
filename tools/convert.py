#!/usr/bin/env python3
"""
Wandelt die Excel-Medienliste in Dateien für die Suchseite um.

Aufruf (im Hauptordner des Projekts):
    python tools/convert.py                          # liest data/Medienliste.xlsx, schreibt nach data/
    python tools/convert.py datei.xlsx ausgabeordner # andere Ein- und Ausgabe
    python tools/convert.py --strict                 # bricht mit Fehler ab, wenn Spendernamen gefunden werden

Erzeugt im Ausgabeordner:
    medien.json       die bereinigten Daten
    medien.js         dieselben Daten, wird von index.html geladen
    pruefbericht.md   Auffälligkeiten in der Excel (Text, für die GitHub-Zusammenfassung)
    pruefbericht.html dasselbe als Webseite

Benötigt nur:  pip install openpyxl
Die Spalte "Spende" wird nie übernommen. Ihre Werte tauchen auch in keinem Bericht auf.
"""
import argparse
import datetime as dt
import html
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent

# Spaltenüberschrift in der Excel  ->  Feldname in den Daten.
# Zugeordnet wird über die Überschrift, die Reihenfolge der Spalten ist egal.
# "spende" wird nur zur Sicherheitsprüfung gelesen (Feld beginnt mit _) und nie ausgegeben.
COLUMNS = {
    "autoren": "author",
    "titel": "title",
    "medien-nr.": "id",
    "signatur": "signature",
    "publikationsdatum": "year",
    "verlag": "publisher",
    "seiten": "pages",
    "isbn": "isbn",
    "standort": "location",
    "anzahl cd": "discs",
    "medienart": "type",
    "sprache": "language",
    "zugang": "added",
    "zugang (datum)": "addedDate",
    "spende": "_donation",
}
REQUIRED = {"author", "title"}  # Blätter ohne diese Spalten (z. B. "Rückenschilder") werden übersprungen


# ---------------------------------------------------------------- Bereinigung
def text(v):
    """Beliebigen Zellwert in bereinigten Text umwandeln (leer -> None)."""
    if v is None:
        return None
    if isinstance(v, float):
        if v != v:  # NaN
            return None
        if v.is_integer():
            v = int(v)
    s = re.sub(r"\s+", " ", str(v)).strip()
    return s or None


def to_int(v):
    if v is None:
        return None
    try:
        return int(float(str(v).replace(",", ".").strip()))
    except ValueError:
        return None


def clean_year(v):
    n = to_int(v)
    if n is not None and 1000 <= n <= 2100:
        return n
    return None  # "k.A." und ähnliches -> keine Angabe


def clean_isbn(v):
    s = text(v)
    if not s:
        return None
    s = re.sub(r"\.0$", "", s)
    s = re.sub(r"[^0-9Xx]", "", s).upper()
    return s or None


def isbn_valid(d):
    if len(d) == 13 and d.isdigit():
        s = sum(int(c) * (1 if i % 2 == 0 else 3) for i, c in enumerate(d[:12]))
        return (10 - s % 10) % 10 == int(d[12])
    if len(d) == 10 and re.fullmatch(r"\d{9}[\dX]", d):
        s = sum((10 - i) * (10 if c == "X" else int(c)) for i, c in enumerate(d))
        return s % 11 == 0
    return False


def iso_date(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    return None


# ---------------------------------------------------------------- Einlesen
def read_workbook(path):
    wb = load_workbook(path, data_only=True, read_only=True)
    items, skipped_sheets = [], []
    for ws in wb.worksheets:
        rows = ws.iter_rows(values_only=True)
        header = next(rows, None)
        if not header:
            skipped_sheets.append(ws.title)
            continue
        colmap = {}
        for idx, h in enumerate(header):
            key = COLUMNS.get(str(h).strip().lower()) if h is not None else None
            if key and key not in colmap.values():
                colmap[idx] = key
        if not REQUIRED.issubset(colmap.values()):
            skipped_sheets.append(ws.title)
            continue
        for rownum, row in enumerate(rows, start=2):
            raw = {key: (row[idx] if idx < len(row) else None) for idx, key in colmap.items()}
            if not text(raw.get("title")) and not text(raw.get("author")):
                continue  # leere Zeile
            item = {
                "author": text(raw.get("author")),
                "title": text(raw.get("title")),
                "id": text(raw.get("id")),
                "signature": text(raw.get("signature")),
                "year": clean_year(raw.get("year")),
                "publisher": text(raw.get("publisher")),
                "pages": to_int(raw.get("pages")),
                "isbn": clean_isbn(raw.get("isbn")),
                "location": text(raw.get("location")),
                "discs": text(raw.get("discs")),
                "type": text(raw.get("type")),
                "language": text(raw.get("language")),
            }
            added = text(raw.get("added"))
            item["added"] = None if (added is None or added.lower() == "bestand") else added
            item["addedDate"] = iso_date(raw.get("addedDate"))
            item["_donation"] = text(raw.get("_donation"))
            item["_sheet"], item["_row"] = ws.title, rownum
            items.append(item)
    return items, skipped_sheets


# ---------------------------------------------------------------- Prüfbericht
def short(s, n=70):
    s = s or ""
    return s if len(s) <= n else s[: n - 1] + "…"


def analyse(items):
    """Liefert die Auffälligkeiten als Liste von Abschnitten (Titel, Hinweis, Tabellen)."""

    def ref(i):
        return (i["_sheet"], i["_row"], i["id"] or "–", short(i["title"]))

    head = ["Blatt", "Zeile", "Medien-Nr.", "Titel"]
    F = []

    missing_id = [i for i in items if not i["id"]]
    by_id = defaultdict(list)
    for i in items:
        if i["id"]:
            by_id[i["id"]].append(i)
    dup_id = {k: v for k, v in by_id.items() if len(v) > 1}
    F.append(dict(title="Medien-Nr. fehlt", count=len(missing_id),
                  hint="Ohne Nummer lässt sich ein Medium nicht eindeutig zuordnen.",
                  tables=[(None, head, [ref(i) for i in missing_id])] if missing_id else []))
    rows = [(k, i["_sheet"], i["_row"], short(i["title"])) for k, v in sorted(dup_id.items()) for i in v]
    F.append(dict(title="Medien-Nr. mehrfach vergeben", count=len(dup_id),
                  hint=f"{len(dup_id)} Nummern, die bei mehreren Einträgen stehen ({len(rows)} Zeilen).",
                  tables=[(None, ["Medien-Nr.", "Blatt", "Zeile", "Titel"], rows)] if rows else []))

    no_isbn = [i for i in items if not i["isbn"]]
    bad_isbn = [i for i in items if i["isbn"] and not isbn_valid(i["isbn"])]
    by_isbn = defaultdict(list)
    for i in items:
        if i["isbn"]:
            by_isbn[i["isbn"]].append(i)
    dup_isbn = {k: v for k, v in by_isbn.items() if len(v) > 1}
    F.append(dict(title="ISBN fehlt", count=len(no_isbn),
                  hint="Diese Medien sind über die ISBN-Suche nicht auffindbar.",
                  tables=[(None, head, [ref(i) for i in no_isbn])] if no_isbn else []))
    F.append(dict(title="ISBN ungültig", count=len(bad_isbn),
                  hint="Falsche Länge oder Prüfziffer stimmt nicht (Tippfehler oder Eigencode?). Die Suche findet sie trotzdem.",
                  tables=[(None, ["ISBN"] + head, [(i["isbn"],) + ref(i) for i in bad_isbn])] if bad_isbn else []))
    rows = [(k, i["_sheet"], i["_row"], i["id"] or "–", short(i["title"])) for k, v in sorted(dup_isbn.items()) for i in v]
    F.append(dict(title="ISBN mehrfach vorhanden", count=len(dup_isbn),
                  hint="Kann gewollt sein (mehrere Exemplare), sonst Doppelerfassung.",
                  tables=[(None, ["ISBN", "Blatt", "Zeile", "Medien-Nr.", "Titel"], rows)] if rows else []))

    no_type = [i for i in items if not i["type"]]
    per_sheet = Counter(i["_sheet"] for i in no_type)
    tables = []
    if no_type:
        tables.append((None, ["Blatt", "Einträge ohne Medienart"], list(per_sheet.items())))
        few = [i for i in no_type if per_sheet[i["_sheet"]] <= 30]
        if few:
            tables.append(("Einzeln aufgeführt (Blätter mit bis zu 30 Fällen)", head, [ref(i) for i in few]))
    F.append(dict(title="Medienart fehlt", count=len(no_type),
                  hint="Diese Medien erscheinen nur unter „Ohne Angabe“ im Filter Medienart.", tables=tables))

    no_sig = [i for i in items if not i["signature"]]
    F.append(dict(title="Signatur fehlt", count=len(no_sig),
                  hint="Ohne Signatur kann das Rückenschild in der Suche nicht angezeigt werden.",
                  tables=[(None, head, [ref(i) for i in no_sig])] if no_sig else []))
    variants = defaultdict(Counter)
    for i in items:
        if i["signature"]:
            variants[i["signature"].lower()][i["signature"]] += 1
    mixed = [v for v in variants.values() if len(v) > 1]
    F.append(dict(title="Signatur uneinheitlich geschrieben", count=len(mixed),
                  hint="Gleiche Signatur mit unterschiedlicher Groß-/Kleinschreibung.",
                  tables=[(None, ["Schreibweisen"], [(" / ".join(f"{s} ({n}×)" for s, n in v.items()),) for v in mixed])] if mixed else []))

    # Sicherheitsprüfung: in "Spende" darf nur "x" oder nichts stehen. Die Werte selbst werden NICHT ausgegeben.
    donors = [i for i in items if i["_donation"] and i["_donation"].lower() != "x"]
    F.append(dict(title="Spalte „Spende“ enthält Namen statt „x“", count=len(donors),
                  hint="Spendernamen dürfen nicht veröffentlicht werden. Bitte durch „x“ ersetzen (nur Zeilennummern, keine Namen aufgeführt).",
                  tables=[(None, ["Blatt", "Zeile"], [(i["_sheet"], i["_row"]) for i in donors])] if donors else []))
    return F, donors


def md_table(head, rows):
    esc = lambda c: str(c).replace("|", "/")
    out = ["| " + " | ".join(esc(h) for h in head) + " |", "|" + "|".join("---" for _ in head) + "|"]
    return "\n".join(out + ["| " + " | ".join(esc(c) for c in r) + " |" for r in rows])


def render_md(items, findings, skipped, source):
    per_sheet = Counter(i["_sheet"] for i in items)
    L = ["# Prüfbericht Medienliste\n",
         f"Quelle: `{source}`, erstellt am {dt.datetime.now().strftime('%d.%m.%Y %H:%M')}\n",
         "Auffälligkeiten in der Excel-Datei. Nichts davon verhindert die Suche, aber die Punkte lassen sich in der Excel korrigieren. "
         "*Zeile* ist die Zeilennummer in Excel.\n",
         "## Übersicht\n",
         md_table(["Blatt", "Einträge"], list(per_sheet.items()) + [("**Gesamt**", f"**{len(items)}**")])]
    if skipped:
        L.append(f"\nÜbersprungene Blätter (keine Spalten *Autoren* und *Titel*): {', '.join(skipped)}")
    L.append("\n## Auffälligkeiten\n")
    L.append(md_table(["Prüfung", "Anzahl"], [(f["title"], f["count"]) for f in findings]) + "\n")
    for f in findings:
        L.append(f"### {f['title']}: {f['count']}\n")
        L.append(f["hint"] + "\n")
        for cap, head, rows in f["tables"]:
            body = (cap + ":\n\n" if cap else "") + md_table(head, rows)
            L.append(f"<details>\n<summary>Liste anzeigen</summary>\n\n{body}\n\n</details>\n")
    return "\n".join(L)


def html_table(head, rows):
    e = lambda c: html.escape(str(c))
    return ("<div class=\"scroll\"><table><thead><tr>" + "".join(f"<th>{e(h)}</th>" for h in head) + "</tr></thead><tbody>" +
            "".join("<tr>" + "".join(f"<td>{e(c)}</td>" for c in r) + "</tr>" for r in rows) + "</tbody></table></div>")


def render_html(items, findings, skipped, source):
    e = html.escape
    per_sheet = Counter(i["_sheet"] for i in items)
    parts = []
    parts.append("<h2>Übersicht</h2>" + html_table(["Blatt", "Einträge"], list(per_sheet.items()) + [("Gesamt", len(items))]))
    if skipped:
        parts.append(f"<p>Übersprungene Blätter (keine Spalten Autoren und Titel): {e(', '.join(skipped))}</p>")
    parts.append("<h2>Auffälligkeiten</h2>" + html_table(["Prüfung", "Anzahl"], [(f["title"], f["count"]) for f in findings]))
    for f in findings:
        parts.append(f"<h3>{e(f['title'])}: {f['count']}</h3><p>{e(f['hint'])}</p>")
        for cap, head, rows in f["tables"]:
            parts.append(f"<details><summary>Liste anzeigen ({len(rows)})</summary>" +
                         (f"<p>{e(cap)}:</p>" if cap else "") + html_table(head, rows) + "</details>")
    now = dt.datetime.now().strftime("%d.%m.%Y %H:%M")
    return f"""<!doctype html>
<html lang="de"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex"><title>Prüfbericht – Medienliste Bücherei Hechendorf</title>
<style>
body{{margin:0;background:#F2F4F1;color:#16262B;font:1rem/1.5 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif}}
main{{width:min(100% - 2rem,62rem);margin:2rem auto 3rem}}
h1{{font:600 2rem/1.1 "Iowan Old Style","Palatino Linotype",Palatino,Georgia,serif;margin:0 0 .5rem}}
h2{{margin:2rem 0 .5rem;font-size:1.3rem}} h3{{margin:1.75rem 0 .25rem;font-size:1.05rem}}
p{{margin:.25rem 0 .6rem;color:#33464B}} a{{color:#1B5E73}}
.scroll{{overflow-x:auto;margin:.5rem 0}}
table{{border-collapse:collapse;background:#fff;font-size:.9rem;min-width:24rem}}
th,td{{padding:.4rem .7rem;border:1px solid #D3DAD7;text-align:left;vertical-align:top}} th{{background:#DDEBEF}}
details{{margin:.4rem 0}} summary{{cursor:pointer;color:#1B5E73;font-weight:600;padding:.3rem 0}}
</style></head><body><main>
<p><a href="./">← Zur Suche</a></p>
<h1>Prüfbericht Medienliste</h1>
<p>Quelle: {e(source)}, erstellt am {now}. Die Liste zeigt Auffälligkeiten in der Excel-Datei. Nichts davon verhindert die Suche. Zeile ist die Zeilennummer in Excel.</p>
{''.join(parts)}
</main></body></html>
"""


# ---------------------------------------------------------------- Ausgabe
def main():
    ap = argparse.ArgumentParser(description="Excel-Medienliste in Daten für die Suchseite umwandeln")
    ap.add_argument("xlsx", nargs="?", default=str(ROOT / "data" / "Medienliste.xlsx"))
    ap.add_argument("out", nargs="?", default=str(ROOT / "data"))
    ap.add_argument("--strict", action="store_true", help="mit Fehler abbrechen, wenn Spendernamen in der Excel stehen")
    args = ap.parse_args()
    xlsx, out = Path(args.xlsx), Path(args.out)

    if not xlsx.exists():
        sys.exit(f"Datei nicht gefunden: {xlsx}\nBitte die Excel-Liste als data/Medienliste.xlsx ablegen.")
    items, skipped = read_workbook(xlsx)
    if not items:
        sys.exit("Keine Einträge gefunden. Erwartet werden Blätter mit den Spalten 'Autoren' und 'Titel'.")

    findings, donors = analyse(items)
    if donors and args.strict:
        rows = ", ".join(f"{i['_sheet']} Zeile {i['_row']}" for i in donors[:10])
        sys.exit(f"ABGEBROCHEN: In der Spalte 'Spende' stehen bei {len(donors)} Einträgen Namen statt 'x' ({rows} ...). "
                 "Spendernamen dürfen nicht veröffentlicht werden. Bitte in der Excel durch 'x' ersetzen und neu hochladen.")

    public = [{k: v for k, v in i.items() if not k.startswith("_") and v not in (None, "")} for i in items]
    data = {
        "meta": {
            "generated": dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "count": len(public),
            "sheets": dict(Counter(i["_sheet"] for i in items)),
        },
        "items": public,
    }
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))

    out.mkdir(parents=True, exist_ok=True)
    (out / "medien.json").write_text(payload, encoding="utf-8")
    (out / "medien.js").write_text("window.MEDIEN=" + payload + ";\n", encoding="utf-8")
    (out / "pruefbericht.md").write_text(render_md(items, findings, skipped, xlsx.name), encoding="utf-8")
    (out / "pruefbericht.html").write_text(render_html(items, findings, skipped, xlsx.name), encoding="utf-8")

    print(f"{len(public)} Medien aus {len(data['meta']['sheets'])} Blatt/Blättern übernommen -> {out}")
    for f in findings:
        if f["count"]:
            print(f"  Hinweis: {f['title']}: {f['count']}")


if __name__ == "__main__":
    main()
