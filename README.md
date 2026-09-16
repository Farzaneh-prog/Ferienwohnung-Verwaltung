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

## Phase 2 — Cowork-Einlese-Pipeline (2026-09-14)

Cowork (Cloud-Automatisierung, läuft unabhängig vom eigenen Rechner) prüft täglich
Airbnb/Booking/Vrbo und schreibt alles, was es findet, an einen kleinen, isolierten
Eingangs-Briefkasten auf **marktresidenz-eisenach.de** selbst (nicht auf ein
fremdes Cloud-Konto) — siehe `wordpress-plugin/marktresidenz-incoming-api.php`.
Zwei getrennte Schlüssel (Schreiben für Cowork, Lesen für das lokale Skript;
in `.env`/`wp-content/mu-plugins/marktresidenz-secrets.php`, nie im Repo) sorgen
dafür, dass ein geleakter Schreib-Schlüssel höchstens Datenmüll erzeugen könnte,
aber niemals Gästedaten auslesen kann.

`tools/apply_incoming.py` holt die wartenden Einträge ab und überträgt sie sicher
in die echten Excel-Dateien:
- **Neue Reservierungen** werden nur an eine bereits leere Zeile im richtigen
  Quartals-Sheet **angehängt**, nie mittendrin eingefügt — `openpyxl` passt beim
  Einfügen/Löschen von Zeilen die Formel-Bezüge anderer Zeilen NICHT automatisch
  an (anders als Excel selbst), ein Insert/Delete mitten im Sheet würde also
  bestehende Formeln stillschweigend kaputt machen. Nur Excel selbst darf Zeilen
  wirklich löschen (macht das automatisch richtig).
- **Stornierungen** werden nie gelöscht, sondern nur mit dem neuen Flag-Feld
  `Storniert` (Spalte BB) markiert (`ja`) — Zuordnung über den
  `Bestätigungs-Code`. So bleiben Formeln unberührt und ein zukünftiges
  Reinigungs-Modul kann stornierte Zeilen einfach überspringen, ohne dass sie
  optisch/inhaltlich wie eine normale Buchung aussehen (frühere Idee, den
  Gästenamen umzubenennen, wurde verworfen — das würde mit der Zeilenzuordnung
  fürs Putzpersonal kollidieren). Manuelles Löschen dieser Zeilen durch dich
  selbst in Excel bleibt jederzeit sicher möglich.
- Jede berührte Datei wird vor dem Schreiben automatisch nach `backups/`
  gesichert; ein Eintrag wird erst nach erfolgreichem Schreiben als erledigt
  markiert (`/incoming/{id}/ack`) — ein Absturz mitten im Lauf führt einfach zu
  einem Retry beim nächsten Aufruf, nichts geht doppelt oder verloren.

**openpyxl-Falle (gefunden 2026-09-14):** `ws.cell(row, col, value=None)`
schreibt NICHTS — `None` ist der Sentinel-Wert für "kein Wert übergeben", die
Zelle bleibt unverändert. Zum Leeren einer Zelle immer `cell.value = None`
direkt setzen, nie über das `value=`-Keyword-Argument.

Welche Felder Cowork mitschicken soll (inkl. E-Mail und den vom Portal selbst
sichtbaren Finanzfeldern wie Zahlbetrag/Reinigungsgebühr/Plattformgebühr) —
bewusst großzügig ("alles was sichtbar ist"), die Zuordnung/Auswahl passiert
erst in `apply_incoming.py`, nichts wird geraten.

## Kurswechsel — manueller Export statt Live-Scraping (2026-09-16)

**Sowohl Airbnb als auch Booking.com verbieten in ihren Nutzungsbedingungen
ausdrücklich automatisierten/KI-gestützten Zugriff** (Booking.com nennt
explizit "AI-powered assistants"; Airbnb schließt auch automatisierten
Zugriff auf den eigenen Account ein) — Konsequenzen reichen bis zur
Kontosperrung/Vertragskündigung. Der Cowork-Ansatz (oben) wurde deshalb für
Airbnb/Booking.com **eingestellt**. Statt automatisiertem Scraping/API-Sync
nutzen wir jetzt die **offiziellen Export-Funktionen** beider Plattformen
(Booking.com Extranet → Reservations → Export; Airbnb Host → Earnings →
Export CSV) — das ist ein regulärer, für Kontoinhaber gedachtes Feature,
kein Scraping, und Farzaneh lädt die Dateien manuell hoch, wann sie will
(wöchentlich, alle paar Tage, …). `app/import_parser.py` liest beide
Booking.com-Export-Varianten (einfach/mit Kontaktdetails) sowie Airbnbs
CSV-Export.

**Wichtige Erkenntnis aus den echten Exportdateien:** Gästename und E-Mail
sind in **keinem** der drei Formate enthalten — bleibt immer ein manueller
Schritt (`/complete`, siehe unten). Booking.com "detailed" liefert dafür
Erwachsene/Kinder sauber; Airbnb nicht. Die Finanzfelder aus Airbnbs Export
(`Betrag`, `Servicegebühr`, `Reinigungsgebühr`) wurden gegen einen echten
Screenshot der Airbnb-"Einkünfte"-Aufschlüsselung verifiziert: `Betrag` =
Auszahlung an den Gastgeber, `Servicegebühr` = Plattformgebühr inkl. MwSt.,
`guest_paid_total` wird als `Betrag + Servicegebühr` rekonstruiert (exakt
207,00 € = 168,81 € + 38,19 € im Testfall).

Drei neue Web-Seiten (`/import`, `/complete`, `/alerts`):
- **`/import`**: Dateien hochladen → Vorschau (neu / bereits vorhanden,
  dedupliziert über `Bestätigungs-Code` gegen die echten Dateien UND
  zwischen mehreren hochgeladenen Dateien / dieselbe Reservierung kann in
  Airbnbs "letzter Monat"- und "nächste 2 Monate"-Export gleichzeitig
  auftauchen / / Stornierungen über Booking.coms `Status`-Spalte) → erst
  nach Bestätigen wird geschrieben. Fehlende Felder (Gästename etc.) werden
  als Platzhalter `Gast (Code <Code>)` gespeichert, nie leer gelassen (eine
  leere Spalte A würde mit der "erste leere Zeile"-Logik kollidieren, die
  neue Zeilen in Lücken alter, manuell gelöschter Reservierungen einfügt).
- **`/complete`**: mobilfreundliche Liste aller Platzhalter-Zeilen zum
  Nachtragen von Gästename/E-Mail/Personenzahl — kein Computer nötig.
- **`/alerts`**: bewusst **von der eigentlichen Excel-Tabelle entkoppelte**
  Mini-Liste ("Ferienwohnung X, Datum Y") für den Fall, dass eine neue
  Buchung sofort einer Putzkraft gemeldet werden muss, aber die vollständige
  Dateneingabe erst später (wöchentlich) passiert — landet in
  `data/quick_alerts.json` (gitignored), nicht in den Excel-Dateien.

Die gemeinsame sichere Schreiblogik (Anhängen statt Einfügen, Backup vor
jedem Schreiben, Storniert-Flag statt Löschen) wurde aus
`tools/apply_incoming.py` nach `app/xlsx_writer.py` ausgelagert — beide
Pfade (Cowork-Warteschlange und manueller Upload) nutzen jetzt exakt
dieselbe Funktion (`process_batch`).

## Struktur

```
app/
  __init__.py       Flask-App-Factory
  config.py         Immobilie → Dateiname-Zuordnung (kein Code-Umbau nötig)
  auth.py           Einfacher Session-Login
  excel_reader.py   Liest Reservierungen aus allen 4 Quartals-Sheets
  routes.py         /login, /logout, /reservations
  templates/        login.html, reservations.html
tools/
  shrink_xlsx.py       Entfernt Formatierungs-Ballast (siehe oben)
  apply_incoming.py    Holt Cowork-Daten ab, schreibt sicher in die echten Dateien
sample_data/        Fake-Demo-Dateien (Struktur wie echte Dateien)
wordpress-plugin/    Eingangs-Briefkasten-Plugin für marktresidenz-eisenach.de
```

## Nächste Phasen (laut Architektur-Dokument)

3. Finanzmodul (vollständige Formeln S–BE)
4. Automatische Gäste-E-Mails (Steuererinnerung, Schlüsselcode)
5. WhatsApp-Koordination der Putzkräfte (Twilio) — nutzt das `Storniert`-Flag
6. GitHub-Veröffentlichung (Screenshots, finaler Cleanup)
