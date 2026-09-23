# Medienliste Bücherei Hechendorf

Suchseite für die Medienliste (Titel, Autor, ISBN, Medienart) mit Verwaltungsseite zum Aktualisieren der Excel-Liste.
Läuft komplett auf GitHub Pages, ohne eigenen Server.

- **Suche:** `https://<benutzer>.github.io/<repository>/`
- **Verwaltung (Büchereiteam):** `https://<benutzer>.github.io/<repository>/verwaltung.html`
- **Prüfbericht** (Auffälligkeiten in der Excel): `.../data/pruefbericht.html`

Die Anleitung für das Team steht in [`ANLEITUNG.md`](ANLEITUNG.md).

## So funktioniert es

```
Excel hochladen ──► Verwaltungsseite ──► data/Medienliste.xlsx im Repository
(bereinigt, geprüft,                            │
 mit Vorschau)                                  ▼
                                    GitHub Action "Website aktualisieren"
                                    prüft die Excel, erzeugt die Suchdaten
                                                │
                                                ▼
                                    GitHub Pages: Suche + Excel-Download
```

- Die Excel-Datei im Repository ist die einzige Quelle. Alle laden dieselbe Datei herunter und wieder hoch.
- Jede Änderung wird von Git gespeichert. Frühere Versionen lassen sich auf der Verwaltungsseite herunterladen.
- Ist die Excel fehlerhaft (z. B. fehlen die Spalten *Autoren* und *Titel*), bricht die Action ab und die bisherige Suchseite bleibt unverändert online.

## Einmalige Einrichtung (Verwalter)

1. **GitHub-Konto:** Empfohlen ist ein gemeinsames Konto der Bücherei (z. B. `buecherei-hechendorf`). Grund: Ein Zugriffsschlüssel funktioniert nur für Repositories, die dem Konto gehören, das ihn erstellt hat. Wer den Schlüssel benutzt, wird im Änderungsverlauf trotzdem mit Namen genannt, weil die Verwaltungsseite den eingegebenen Namen in die Änderungsnachricht schreibt.
2. **Repository anlegen:** Neues **öffentliches** Repository, z. B. `medienliste`.
3. **Dateien hochladen:** *Add file → Upload files* und den gesamten Inhalt dieses Ordners hineinziehen (`index.html`, `verwaltung.html`, `data/`, `tools/`, `vendor/`, `README.md`, `ANLEITUNG.md`, `.gitignore`).
   Der versteckte Ordner `.github` wird beim Hochladen im Browser oft nicht mitgenommen. Dann so anlegen: *Add file → Create new file*, als Namen `.github/workflows/website.yml` eintippen (die Schrägstriche legen die Ordner an) und den Inhalt der mitgelieferten Datei `website.yml` einfügen.
4. **Pages einschalten:** *Settings → Pages → Build and deployment → Source: **GitHub Actions***.
5. **Erster Lauf:** Im Reiter *Actions* den Lauf „Website aktualisieren“ abwarten. Startet er nicht von selbst: *Run workflow*. Danach ist die Suche unter der Adresse oben erreichbar.
6. **Zugriffsschlüssel erstellen** (mit dem Konto aus Schritt 1): <https://github.com/settings/personal-access-tokens/new>
   - Name: `Medienliste`, Ablauf: 1 Jahr
   - *Repository access → Only select repositories* → das Repository auswählen
   - *Repository permissions → Contents → Read and write*
   - Schlüssel kopieren (`github_pat_…`) und dem Team geben (persönlich oder per Passwortmanager, nicht per E-Mail).
7. **Test:** Auf `verwaltung.html` verbinden, die Excel herunterladen, eine Kleinigkeit ändern, hochladen, veröffentlichen. Die Statusanzeige muss „Fertig“ melden und die Änderung in der Suche erscheinen.

Läuft der Schlüssel ab oder soll jemand keinen Zugriff mehr haben: Schlüssel unter *Settings → Developer settings → Fine-grained tokens* löschen und einen neuen erstellen.

## Datenschutz

Das Repository und die Excel-Datei sind öffentlich einsehbar. Deshalb gilt:

- **Spendernamen:** Die Spalte *Spende* darf nur „x“ oder nichts enthalten. Beim Hochladen ersetzt die Verwaltungsseite Namen automatisch durch „x“. Zusätzlich bricht die Action ab, falls doch Namen in der Excel stehen, und nennt nur Zeilennummern.
- **Dateieigenschaften:** Beim Hochladen werden Autor, letzter Bearbeiter, lokaler Ordnerpfad und Namen aus Kommentar-Verwaltung entfernt. Sonst würden die Namen der Bearbeiter mit veröffentlicht.
- **Kommentare in Zellen** bleiben erhalten und sind öffentlich. Keine vertraulichen Notizen als Excel-Kommentar ablegen.
- **Nur über die Verwaltungsseite hochladen.** Wer die Excel direkt über die GitHub-Oberfläche hochlädt, umgeht die automatische Bereinigung.
- Die Suchseite selbst lädt nichts von fremden Servern (keine Google-Schriften, kein CDN). Die zwei Bibliotheken der Verwaltungsseite liegen im Ordner `vendor/`.

## Dateien

| Datei | Zweck |
|---|---|
| `index.html` | Suchseite |
| `verwaltung.html` | Download, Upload mit Prüfung und Vorschau, Versionsverlauf |
| `data/Medienliste.xlsx` | die Excel-Liste (Quelle aller Daten) |
| `tools/convert.py` | Excel → Suchdaten und Prüfbericht (läuft in der Action, lokal mit `python tools/convert.py`) |
| `.github/workflows/website.yml` | die Action „Website aktualisieren“ |
| `vendor/` | SheetJS und JSZip für die Verwaltungsseite (Lizenzen in `vendor/LICENSES.md`) |

Erzeugt werden bei jedem Lauf (nicht im Repository): `medien.js`, `medien.json`, `pruefbericht.md`, `pruefbericht.html`.

## Lokal ausprobieren

`pip install openpyxl`, dann `python tools/convert.py` und danach `index.html` per Doppelklick öffnen.

## Excel-Spalten

Die Zuordnung erfolgt über die Überschriften in Zeile 1: `Autoren`, `Titel`, `Medien-Nr.`, `Signatur`, `Publikationsdatum`, `Verlag`, `Seiten`, `ISBN`, `Standort`, `Anzahl CD`, `Medienart`, `Sprache`, `Zugang`, `Zugang (Datum)`, `Spende`.
Die Reihenfolge der Spalten darf sich ändern, die Überschriften bitte nicht. Alle Blätter mit den Spalten `Autoren` und `Titel` werden gelesen, andere (z. B. `Rückenschilder`) ignoriert.

## Anpassen

- **Farben der Rückenschilder:** in `index.html` das Objekt `SIGNATUR_FARBEN`.
- **Trefferzahl pro Seite:** `PAGE` in `index.html`.
- **Zwischenspeicher:** GitHub Pages speichert Seiten bis zu 10 Minuten zwischen. Eine Änderung kann daher etwas verzögert sichtbar werden.
