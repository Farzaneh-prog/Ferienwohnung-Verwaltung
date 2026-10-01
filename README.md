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
  checkin_reminder.py        Check-in-Erinnerungs-E-Mail mit .ics-Kalenderalarm
                            (siehe unten)
  notified_checkins.py       Persistentes Gedächtnis, welche Reservierung schon
                            eine Erinnerung bekommen hat (verhindert Doppel-Mails)
  quick_alerts.py           Von der Excel-Datei entkoppelte Mini-Warnliste
  routes.py / routes_workflow.py   /login, /reservations, /import, /complete, /alerts
  templates/               Alle HTML-Seiten
tools/
  shrink_xlsx.py            Entfernt Formatierungs-Ballast (siehe unten)
  apply_incoming.py         Cowork-Anbindung — nicht mehr in Benutzung (siehe oben)
sample_data/               Fake-Demo-Dateien (Struktur wie echte Dateien)
wordpress-plugin/          Nicht mehr benutzter Eingangs-Briefkasten (siehe oben)
docs/STATUS.md              Laufend gepflegter Übergabe-/Entscheidungs-Stand
Dockerfile, docker-compose.yml, wsgi.py, .dockerignore
                            Für den Dauerbetrieb auf dem QNAP-NAS (siehe unten)
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

## Dauerbetrieb auf dem QNAP-NAS (seit 2026-09-25)

Die App läuft nicht mehr nur lokal auf Farzanehs Rechner, sondern dauerhaft
(24/7) als Docker-Container auf ihrem eigenen QNAP-NAS (TS-431P2, ARM) —
öffentlich erreichbar unter `https://<NAS-DOMAIN>:4844/`
(DDNS + Let's-Encrypt-Zertifikat + Fritz!Box-Portfreigabe + QNAP
Reverse-Proxy). Grund: Phasen 2–4 unten brauchen einen Server, der auch
erreichbar ist, wenn Farzanehs Laptop aus ist (z.B. auf Reise, nur Handy
dabei). Die echten Excel-Dateien liegen dabei weiterhin nur auf der
eigenen Hardware (NAS-Freigabeordner `Ferienwohnung-Data`, in den
Container als `/data` gemountet) — nicht bei einem fremden Hosting-Anbieter.

Produktions-Entrypoint ist `wsgi.py` (nutzt `waitress`, **kein**
`debug=True` wie `run.py` — der Werkzeug-Debugger wäre bei einer öffentlich
erreichbaren App ein Sicherheitsrisiko).

Der Code liegt auf dem NAS unter `/share/CACHEDEV1_DATA/ferienwohnung-app`
(per `scp`/tar kopiert, **nicht** per `git clone` — auf dem NAS ist kein
Git installiert). Nach jeder Code-Änderung muss der Ordner neu übertragen
und der Container neu gebaut werden:

```bash
# lokal
DOCKER=/share/CACHEDEV1_DATA/.qpkg/container-station/bin/docker
# via ssh admin@<NAS-IP> (SSH-Key liegt unter ~/.ssh/<ssh-key>):
export DOCKER_HOST=unix:///var/run/system-docker.sock
cd /share/CACHEDEV1_DATA/ferienwohnung-app
$DOCKER compose up -d --build
```

Volle Details (Netzwerk-Setup, SSH-Key-Einrichtung, der `$`-Escaping-Stolperstein
in `.env` für Docker Compose, warum Reverse-Proxy Port 8443 statt 443 nutzt)
stehen in `docs/STATUS.md`, Abschnitt 12.

## Check-in-Erinnerung per E-Mail (seit 2026-09-25)

Sobald eine neue Reservierung geschrieben wird (`/import`), verschickt
`app/checkin_reminder.py` sofort eine E-Mail an Farzaneh (Gmail) mit einer
`.ics`-Kalenderanhang. Öffnet sie den Anhang am iPhone, bietet iOS "Zum
Kalender hinzufügen" an — der Termin selbst liegt auf **16:00 Uhr am Tag
vor der Anreise** und trägt einen `VALARM`, sodass iOS von sich aus einen
echten Alarm auslöst (kein Push-Dienst, kein Twilio nötig).

- **Kein täglicher Check** — die Mail geht sofort beim Anlegen der
  Reservierung raus, nicht erst am Vortag (2026-09-25 bewusst so von
  Farzaneh entschieden).
- **Versand:** über Gmail mit einem App-Passwort (`GMAIL_USER` /
  `GMAIL_APP_PASSWORD` in `.env` — siehe `.env.example`). Fehlt eins von
  beiden, wird der Import trotzdem normal fortgesetzt, nur die Mail
  bleibt aus (Warnung im Log).
- **Empfänger bewusst NICHT das private Gmail-Konto** (`NOTIFY_TO_EMAIL`
  in `.env`, Default wäre `GMAIL_USER`) — die Gmail-App auf dem iPhone
  zeigt bei `.ics`-Anhängen keinen "Zum Kalender hinzufügen"-Button
  (bestätigtes 2026-09-25-Problem, siehe Git-Historie). Farzaneh empfängt
  die Erinnerungen stattdessen auf `<owner-mail>`,
  das über die native iOS-Mail-App eingerichtet ist — dort funktioniert
  der Kalender-Import automatisch.
- **Kein Doppelversand:** `data/notified_checkins.json` merkt sich jede
  bereits benachrichtigte Reservierung (Bestätigungscode, sonst
  Gastname+Anreisedatum als Fallback).
- **Backfill für bereits bestehende Buchungen:** einmalig
  `python -m app.checkin_reminder` ausführen — schickt die Erinnerung für
  alle aktuell noch in der Zukunft liegenden Reservierungen in beiden
  Dateien nach (bereits benachrichtigte werden übersprungen).
- **Stornierung:** Wenn eine Reservierung storniert wird, für die schon
  eine Erinnerung verschickt wurde, geht automatisch eine zweite Mail
  raus ("STORNIERT: ... — Kalendereintrag löschen") mit Datum/Uhrzeit des
  betroffenen Kalendertermins zum manuellen Löschen, plus einem
  Storno-`.ics` (`METHOD:CANCEL`, gleiche UID) als Best-Effort-Versuch,
  den Termin automatisch zu entfernen — das klappt aber nur zuverlässig,
  wenn die Kalender-App den ursprünglichen Termin noch mit dieser UID
  verknüpft; manuelles Löschen bleibt der garantierte Weg. War noch nie
  verschickt worden (z.B. weil die Reservierung storniert wurde, bevor
  eine Erinnerung fällig war), bleibt die Mail ganz aus — es gibt dann
  nichts zu löschen.

## WhatsApp-Koordination der Putzkräfte (seit 2026-10-01)

Vollständig gebaut und live getestet (Twilio-Sandbox, Testrollen auf Farzanehs
Telefon); läuft auf dem NAS im **Trockenlauf** (nur Log), bis die eigene
WhatsApp-Nummer und die Vorlagen freigegeben sind. Stand, Entscheidungen und
die Schritte bis zum Live-Gang: `docs/STATUS.md` Abschnitte 15–18.

**Ablauf:** > 10 Tage vor Abreise nichts Automatisches; ab 10 Tagen wird die
Prioritätskette nacheinander angeschrieben (Jennifer/Mehrnaz fair im Wechsel →
Manuela → Tahmine → Ramic, je nach Objekt/Wochentag/Vertrag/Urlaub); ab 3 Tagen
geht ein "dringend"-Broadcast an alle noch Infrage-Kommenden. Ja bestätigt
(Putzplan Spalte A, Vorname wie in der echten Datei), Nein/Vielleicht/1 h ohne
Antwort → sofort die nächste Person. Antworten sind jederzeit änderbar; ein
"Ja" auf eine schon vergebene Reinigung wird **Reserve**, springt die bestätigte
Person ab, rückt die erste Reserve nach (sonst wird neu angefragt + Alarm).
Keine Nachricht zwischen 22:00 und 08:00 (Ausnahme: Stornierung nach 19:00 am
Vorabend).

- `app/cleaner_roster.py` — Putzkräfte in `data/cleaners.json` (gitignored):
  Name, WhatsApp-Nummer, Objekte, Wochentage, Satz, Vertragsende, Prioritäts-
  Stufe, `sheet_name` (so steht die Person in Putzplan-Spalte A);
  `match_cleaner` ordnet Freitext-Werte zu ("ich", Notizen → keine Person).
- `app/cleaner_coordination.py` — Entscheidungslogik, Antworten, Reserve,
  Zeilenstatus; State in `data/cleaner_coordination_state.json`.
  Urlaub: **hart** (unterwegs, wird nie angeschrieben) oder **weich** (freie
  Woche, ganz ans Kettenende) — Pflege über die Dashboard-Seite `/leave`.
- `app/day_before.py` — 16:00 Erinnerung am Vortag (inkl. nächste Gäste:
  Erwachsene, Kinder < 18, Kinder < 3 aus Spalte I), Stornierungen
  (Putzplan-Zeile bleibt, Spalte A merkt die zuständige Person; 19:00
  "entfällt"), Ersatzbuchung (gleiche Person zuerst), Vertragsende-Warnung.
- `app/scheduler.py` — ereignisgesteuert statt Dauer-Polling: Import /
  manuelle Suche / Stornierung lösen sofort eine Prüfung aus; Einmal-Timer für
  "1 h ohne Antwort" und "08:00 nach der Ruhezeit"; Tagesjobs 16:00 und 19:00.
  `SCHEDULER_ENABLED` (Standard aus) und `SCHEDULER_DRY_RUN` (Standard an).
- `app/clock_guard.py` — vergleicht die Systemuhr mit mehreren Webservern;
  > 15 min Abweichung → Warnung (WhatsApp an Farzaneh) und im Live-Modus keine
  Nachrichten.
- `app/whatsapp_sender.py` / `tools/setup_whatsapp_templates.py` /
  `tools/submit_whatsapp_templates.py` — Twilio-Wrapper; alle
  business-initiierten Nachrichten nutzen deutsche Content Templates
  (müssen für den echten Absender von WhatsApp freigegeben werden).
- `app/whatsapp_webhook.py` — Eingangs-Webhook (`/webhooks/whatsapp`,
  Signaturprüfung gegen `TWILIO_WEBHOOK_URL`); antwortet der Putzkraft sofort
  auf Deutsch, leitet Freitext an Farzaneh weiter (Sprachnachrichten nur als
  Hinweis).
- `app/owner_alerts.py` — kurze Zeile an Farzaneh für jede Nachricht/Antwort
  und alle Alarme (`OWNER_WHATSAPP_NUMBER`).
- Dashboard "Putz-Alerts": Datum eintragen und "Suche starten", Statuszeile je
  Eintrag, Formular "Stornierung melden"; `/leave` für Urlaub.
- Test- und Sicherheitsmodus: `SCHEDULER_ONLY_CLEANERS=test1,test2` +
  `COORDINATION_COOLDOWN_MINUTES` — nur Testrollen und Testzeilen, echte
  Putzplan-Zeilen/Putzkräfte bleiben unberührt.
- Excel-Dateien werden **atomar** gespeichert (`save_workbook_atomic`), der
  Import-Bestätigungsschritt überspringt schon vorhandene Buchungen.

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

## Nächste Phasen (Stand 2026-09-25, siehe docs/STATUS.md Abschnitt 10 & 12)

1. **Passendes Frontend** — die aktuelle Oberfläche ist bewusst minimal und war nur
   zum Testen gedacht.
2. ~~**Erinnerung einen Tag vor Gästeankunft** an Farzaneh selbst.~~ **Erledigt
   (2026-09-25)** — siehe Abschnitt "Check-in-Erinnerung per E-Mail" oben.
3. **WhatsApp-Koordination der Putzkräfte** (Twilio) — **gebaut und getestet
   (2026-10-01)**, Scheduler im Trockenlauf. Fehlt bis zum Live-Gang: Freigabe
   der 9 Vorlagen für die eigene Nummer (die neue Nummer, siehe `.env` auf dem NAS, als WhatsApp-Sender
   registriert, Status Online), Umstellen von `TWILIO_WHATSAPP_FROM`, ein kurzer
   Echt-Test und die Entscheidung `SCHEDULER_DRY_RUN=0`. Siehe `docs/STATUS.md`
   Abschnitt 18.
4. ~~**Automatisches Eintragen von Putzkraft-Name** in die Excel-Datei,
   sobald eine Reinigung über WhatsApp bestätigt wurde.~~ **Im Rahmen von
   Stage 1 miterledigt** (`app/putzplan_writer.assign_cleaner`) — nur das
   automatische Eintragen des **Preises** in die GästeListe-Spalten T/U/V
   (aktuell Platzhalter) ist noch offen, siehe `docs/STATUS.md` Abschnitt
   15 ("Stage 2+").

Die Hosting-Entscheidung, an der Phasen 2–4 vorher hingen (ein öffentlich
erreichbarer, dauerhaft laufender Server), ist am 2026-09-25 gefallen und
umgesetzt (QNAP-NAS, siehe Abschnitt "Dauerbetrieb" oben) — diese drei Phasen
sind also jetzt infrastrukturell nicht mehr blockiert.
