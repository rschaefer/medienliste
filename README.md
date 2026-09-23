# Medienliste Bücherei Hechendorf

Suchseite für die Medienliste: Suche nach Titel, Autor, ISBN und Medienart.
Reine statische Webseite, läuft auf GitHub Pages ohne Server.

## Inhalt

| Datei | Zweck |
|---|---|
| `index.html` | die Suchseite (alles in einer Datei, keine externen Dienste, keine Schriften von Google) |
| `data/medien.js` | die Daten, die die Seite lädt (wird aus der Excel erzeugt) |
| `data/medien.json` | dieselben Daten als JSON |
| `data/pruefbericht.md` | Auffälligkeiten in der Excel: doppelte Nummern, fehlende oder ungültige ISBN usw. |
| `tools/convert.py` | wandelt die Excel in die Datendateien um |

Die Spalte **Spende** wird nicht übernommen. Spendernamen tauchen weder in den Daten noch auf der Seite auf.

## Auf GitHub veröffentlichen

1. Auf github.com ein neues **öffentliches** Repository anlegen, z. B. `medienliste`.
2. Den Inhalt dieses Ordners hochladen (*Add file → Upload files*, alle Dateien und Ordner hineinziehen).
3. *Settings → Pages → Build and deployment → Source: Deploy from a branch*, Branch `main`, Ordner `/ (root)`, speichern.
4. Nach ein bis zwei Minuten ist die Seite erreichbar unter `https://<benutzername>.github.io/medienliste/`.

## Lokal ausprobieren

`index.html` per Doppelklick öffnen. Das funktioniert ohne Server, weil die Daten als `medien.js` geladen werden.

## Daten aktualisieren (bis Phase 2 fertig ist)

1. Aktuelle Excel als `data/Medienliste.xlsx` ablegen (die Datei ist in `.gitignore` eingetragen und wird nicht mit hochgeladen).
2. Einmalig `pip install openpyxl`, dann `python tools/convert.py`.
3. Geänderte Dateien in `data/` auf GitHub hochladen.

In Phase 2 übernimmt eine GitHub Action diese Schritte automatisch.

Die Zuordnung der Excel-Spalten erfolgt über die Überschriften (`Autoren`, `Titel`, `Medien-Nr.`, `Signatur`, `Publikationsdatum`, `Verlag`, `Seiten`, `ISBN`, `Standort`, `Anzahl CD`, `Medienart`, `Sprache`, `Zugang`, `Zugang (Datum)`). Die Reihenfolge der Spalten darf sich ändern, die Überschriften sollten gleich bleiben. Alle Blätter mit den Spalten `Autoren` und `Titel` werden eingelesen, andere (z. B. `Rückenschilder`) ignoriert.

## Anpassen

- **Farben der Rückenschilder:** in `index.html` das Objekt `SIGNATUR_FARBEN` ändern. Nicht aufgeführte Signaturen bekommen automatisch eine feste Farbe.
- **Trefferzahl pro Seite:** `PAGE` in `index.html`.

## So sucht die Seite

- Alle Suchwörter müssen vorkommen, die Reihenfolge ist egal (`Astrid Lindgren` findet `Lindgren, Astrid`).
- Umlaute und Groß-/Kleinschreibung sind egal (`Müller`, `Mueller` und `Muller` finden sich gegenseitig, ebenso `Straße` und `Strasse`).
- ISBN mit oder ohne Bindestriche, als ISBN-10 oder ISBN-13.
- Auch die Medien-Nr. (`1399`), die Medienart (`Tonie`), das Erscheinungsjahr und der Verlag werden bei "Suchen in: Allem" berücksichtigt.
- Die Adresszeile enthält die aktuelle Suche und lässt sich als Link weitergeben.
