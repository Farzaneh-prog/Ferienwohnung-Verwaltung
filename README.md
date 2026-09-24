# Marktresidenz — Automatisierungs-Dashboard

Privates Web-Dashboard zur Verwaltung der Reservierungen für **Marktresidenz Karlstraße**
und **Marktresidenz Eisenach**. Architektur und Zielbild: siehe
`سند-معماری-سیستم-مهمانخانه.md` und `نظافتچی_ها-و-فرمول_های-مالی.md` (lokal, nicht
im Repo — enthalten echte Namen/Sätze). `docs/STATUS.md` ist der laufend gepflegte
Übergabe-Stand — dort steht am genauesten, was gerade fertig ist und was als Nächstes
kommt.

## Aktueller Stand

Lesen und Schreiben der echten Reservierungsdaten ist fertig und in der Praxis
getestet:

- **Import** offizieller Booking.com-/Airbnb-Exportdateien, mit Vorschau vor jedem
  Schreiben, Duplikat-Erkennung und automatischer Backup-Sicherung
- **Manuelles Nachtragen** von Gästename/E-Mail/Personenzahl für Felder, die kein
  Export liefert
- **Stornierungen**, ohne Zeilen zu löschen (Formel-Sicherheit) — mit einem
  zusätzlichen, von der Excel-Datei unabhängigen Gedächtnis (`known_codes.json`),
  damit ein manuell gelöschter Stornierungs-Eintrag bei einem späteren erneuten
  Import nicht wieder als "neu" auftaucht
- Alle bestätigten Finanz-Formelspalten (S–BE) sind hartkodiert und mit Farzaneh
  Zelle für Zelle abgeglichen — keine geratenen Formeln
- **Putzplan-Automatisierung** (`Putzplan2026.xlsx`): jede neue Reservierung bekommt
  automatisch eine Zeile mit dem Auszugsdatum der *vorherigen* Buchung und der
  Personenzahl der *nächsten* Buchung (nicht der eigenen!) — die Putzkraft muss
  wissen, für wen sie vorbereitet, nicht wer gerade ausgezogen ist

Details, offene Fragen und die genaue Entscheidungshistorie: `docs/STATUS.md`.

## Kurswechsel — manueller Export statt Live-Scraping (2026-09-16)

**Sowohl Airbnb als auch Booking.com verbieten in ihren Nutzungsbedingungen
ausdrücklich automatisierten/KI-gestützten Zugriff** (Booking.com nennt
explizit "AI-powered assistants"; Airbnb schließt auch automatisierten
Zugriff auf den eigenen Account ein) — Konsequenzen reichen bis zur
Kontosperrung/Vertragskündigung. Ein früher geplanter Cloud-Automatisierungs-Ansatz
(Cowork, siehe `tools/apply_incoming.py` und `wordpress-plugin/` — Code bleibt im
Repo, wird aber nicht mehr benutzt) wurde deshalb für Airbnb/Booking.com
**eingestellt**. Stattdessen nutzen wir die **offiziellen Export-Funktionen** beider
Plattformen (Booking.com Extranet → Reservations → Export; Airbnb Host → Earnings →
Export CSV) — ein regulärer, für Kontoinhaber gedachtes Feature, kein Scraping.
Farzaneh lädt die Dateien manuell hoch, wann sie will.

**Beide Plattformen ändern ihr Exportformat gelegentlich ohne Ankündigung** —
Airbnbs CSV wechselte am 2026-09-22 von cp1252/Tab-getrennt zu UTF-8/Komma-getrennt
mit echtem Quoting; `app/import_parser.py` erkennt inzwischen beide Varianten
automatisch.

**Wichtige Erkenntnisse aus den echten Exportdateien:**
- Gästename und E-Mail sind in **keinem** der Formate enthalten — bleibt immer ein
  manueller Schritt (`/complete`).
- Booking.com "detailed" liefert Erwachsene/Kinder sauber; die "simple"-Variante
  und Airbnb nicht.
- Airbnbs `Servicegebühr` ist **brutto** (inkl. 19% MwSt.), nicht netto — wird vor
  dem Schreiben in die Nettobetrag-Spalte durch 1,19 geteilt, sonst würde die
  eigene Umsatzsteuer-Formel die MwSt. doppelt berechnen.
- Booking.coms "Unit type"-Text für Eisenach ist ein Marketingname ohne das Wort
  "Eisenach" — die Immobilien-Erkennung berücksichtigt das inzwischen.

## Struktur

```
app/
  __init__.py           Flask-App-Factory
  config.py              Immobilie → Dateiname-Zuordnung, Putzplan-Konfiguration
  auth.py                 Einfacher Session-Login
  excel_reader.py         Liest Reservierungen aus GästeListe (alle Quartals-Sheets)
  xlsx_writer.py           Sichere Schreiblogik: Anhängen, Stornieren, Nachtragen
  import_parser.py         Parser für alle 3 Export-Formate (Booking einfach/
                            detailliert, Airbnb CSV)
  putzplan_writer.py        Schreibt/synchronisiert Putzplan2026.xlsx
  known_codes.py            Persistentes Duplikat-Gedächtnis (überlebt manuelles
                            Löschen einer Zeile)
  quick_alerts.py           Von der Excel-Datei entkoppelte Mini-Warnliste
  routes.py / routes_workflow.py   /login, /reservations, /import, /complete, /alerts
  templates/               Alle HTML-Seiten
tools/
  shrink_xlsx.py            Entfernt Formatierungs-Ballast (siehe unten)
  apply_incoming.py         Cowork-Anbindung — nicht mehr in Benutzung (siehe oben)
sample_data/               Fake-Demo-Dateien (Struktur wie echte Dateien)
wordpress-plugin/          Nicht mehr benutzter Eingangs-Briefkasten (siehe oben)
docs/STATUS.md              Laufend gepflegter Übergabe-/Entscheidungs-Stand
```

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
Formatierungs-Ballast (keine echten Werte/Formeln betroffen). Ladezeit danach:
<0,5 Sekunden statt ~2 Minuten. Bei einer neuen Jahresdatei, die sich wieder
langsam anfühlt, dasselbe Skript erneut laufen lassen (Anleitung im
Skript-Docstring). Backups der Originaldateien vor dem Schrumpfen liegen in
`backups/` (gitignored, nicht versioniert).

## Sicherheitsprinzipien beim Schreiben

- **Nie mitten im Sheet einfügen/löschen** — `openpyxl` passt beim Einfügen/Löschen
  von Zeilen die Formel-Bezüge anderer Zeilen NICHT automatisch an (anders als
  Excel selbst). Neue Reservierungen landen in einer bereits leeren Zeile oder
  werden ans Ende angehängt; nur Excel selbst darf Zeilen wirklich verschieben.
  (`Putzplan2026.xlsx` ist die eine Ausnahme — dort gibt es keine einzige Formel,
  ein echtes Einfügen an der chronologisch richtigen Stelle ist dort also sicher.)
- **Stornierungen werden nie gelöscht**, sondern nur mit `Storniert` = `ja`
  markiert. Manuelles Löschen durch Farzaneh selbst in Excel bleibt möglich —
  `known_codes.json` sorgt dafür, dass das keine Duplikate beim nächsten Import
  erzeugt.
- **Jede berührte Datei wird vor dem Schreiben automatisch gesichert** (`backups/`,
  gitignored).
- `ws.cell(row, col, value=None)` schreibt NICHTS — `None` ist der Sentinel-Wert für
  "kein Wert übergeben". Zum Leeren einer Zelle immer `cell.value = None` direkt
  setzen.

## Nächste Phasen (Stand 2026-09-24, siehe docs/STATUS.md Abschnitt 10)

1. **Passendes Frontend** — die aktuelle Oberfläche ist bewusst minimal und war nur
   zum Testen gedacht.
2. **Erinnerung einen Tag vor Gästeankunft** an Farzaneh selbst.
3. **WhatsApp-Koordination der Putzkräfte** (Twilio) — Rotations-/Eskalations-Logik
   ist bereits dokumentiert, braucht aber noch eine Hosting-Entscheidung (öffentlich
   erreichbarer Webhook) und einen Verfügbarkeits-Kalender der Putzkräfte.
4. **Automatisches Eintragen von Putzkraft-Name und -Preis** in die Excel-Datei,
   sobald eine Reinigung über WhatsApp bestätigt wurde.

Phasen 2–4 hängen alle an derselben Hosting-Entscheidung — sinnvoll, die einmal
gemeinsam zu treffen statt dreimal einzeln.
