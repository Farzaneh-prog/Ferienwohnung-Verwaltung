# Marktresidenz — Automatisierungs-Dashboard (Phase 1)

Privates Web-Dashboard zur Verwaltung der Reservierungen für **Marktresidenz Karlstraße**
und **Marktresidenz Eisenach**. Architektur und Zielbild: siehe
`سند-معماری-سیستم-مهمانخانه.md` und `نظافتچی_ها-و-فرمول_های-مالی.md`.

## Phase 1 (dieser Stand)

- Projektgrundgerüst (Flask)
- Einfacher Session-Login (ein Admin-Account)
- Lesen + Anzeigen der Reservierungen aus den echten Excel-Dateien (nur lesend,
  keine Berechnung/Bearbeitung — das kommt in Phase 3 laut Architektur-Dokument)

## Wichtiger Unterschied zur ursprünglichen Architektur-Annahme

Das Architektur-Dokument ging von **einer** Datei `reservations.xlsx` mit **zwei
Sheets** (ein Sheet pro Immobilie) aus. Die tatsächlichen Dateien sind stattdessen:

- `GästeListe_2026K.xlsx` → Karlstraße
- `GästeListe_2026_Pf.xlsx` → Eisenach

Jede Datei hat **vier Quartals-Sheets** (`1`, `2`, `3`, `4`) statt eines
Immobilien-Sheets, plus Hilfs-Sheets (`Übersicht`, `ÜbersichtSteuer`, `Muster`, …),
die keine Reservierungsdaten enthalten und ignoriert werden. Die Spaltenstruktur
A–BE entspricht ansonsten weitgehend dem Dokument. Diese Datei-Zuordnung liegt in
`app/config.py` (`PROPERTY_FILES`) — bei einer neuen Jahresdatei dort nur den
Dateinamen anpassen, kein Code-Umbau nötig.

## Setup

```bash
pip install -r requirements.txt
copy .env.example .env
```

In `.env` echten Benutzernamen setzen und einen Passwort-Hash erzeugen:

```bash
python -c "from werkzeug.security import generate_password_hash; print(generate_password_hash('dein-passwort'))"
```

Den Hash in `ADMIN_PASSWORD_HASH` eintragen.

## Starten

```bash
python run.py
```

Dann `http://127.0.0.1:5000` öffnen und einloggen.

## Demo-Daten

`sample_data/` enthält Beispieldateien mit derselben Struktur, aber frei
erfundenen Gästedaten — für Demo/GitHub-Veröffentlichung ohne echte Gästedaten.
Um damit zu testen, `DATA_DIR` in `.env` auf `sample_data` setzen und die
Dateinamen in `app/config.py` entsprechend anpassen (oder die Sample-Dateien
temporär in den Projekt-Root kopieren und umbenennen).

## Excel-Dateigröße (behoben 2026-09-14)

Die echten `GästeListe_*.xlsx`-Dateien waren ursprünglich 28-40 MB und brauchten
~2 Minuten zum Einlesen, obwohl die echten Reservierungsdaten pro Sheet nur
~30-95 Zeilen sind. Ursache: eine frühere Formatierung (Rahmen/Füllung) wurde
auf ganze Spalten angewendet, wodurch Excel leere, aber formatierte Zeilen bis
Zeile ~1.048.576 gespeichert hat. `tools/shrink_xlsx.py` entfernt diesen reinen
Formatierungs-Ballast (keine echten Werte/Formeln betroffen — vor dem Einsatz
gegen die Originaldateien wurde jede Zelle mit Inhalt verglichen, 0 Abweichungen).
Ladezeit danach: <0.5 Sekunden statt ~2 Minuten. Bei einer neuen Jahresdatei, die
sich wieder langsam anfühlt, dasselbe Skript erneut laufen lassen (Anleitung im
Skript-Docstring). Backups der Originaldateien vor dem Schrumpfen liegen in
`backups/` (gitignored, nicht versioniert).

## Struktur

```
app/
  __init__.py       Flask-App-Factory
  config.py         Immobilie → Dateiname-Zuordnung (kein Code-Umbau nötig)
  auth.py           Einfacher Session-Login
  excel_reader.py   Liest Reservierungen aus allen 4 Quartals-Sheets
  routes.py         /login, /logout, /reservations
  templates/        login.html, reservations.html
sample_data/        Fake-Demo-Dateien (Struktur wie echte Dateien)
```

## Nächste Phasen (laut Architektur-Dokument)

2. Cowork-Task zum automatischen Einlesen neuer Reservierungen
3. Finanzmodul (vollständige Formeln S–BE)
4. Automatische Gäste-E-Mails (Steuererinnerung, Schlüsselcode)
5. WhatsApp-Koordination der Putzkräfte (Twilio)
6. GitHub-Veröffentlichung (Screenshots, finaler Cleanup)
