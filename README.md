# LoopForgeLab — Marketing Teaser & Outreach

**Projekt:** Forge Engine · Pilot-Akquise Q2 2026  
**Einstieg:** Diese Datei — danach je nach Aufgabe die Ordner unten.

---

## In 30 Sekunden: Wo finde ich was?

| Ich will … | Ordner |
|------------|--------|
| **Den aktuellen Teaser** (PPTX, HTML) | [`02_Aktuell/Teaser/Variante_E_Interpack/`](02_Aktuell/Teaser/Variante_E_Interpack/) |
| **Design-Vorlage** (CSS, Bilder) | [`02_Aktuell/Teaser/Design_System_Option_B/`](02_Aktuell/Teaser/Design_System_Option_B/) |
| **LinkedIn / Outreach** (InMails, Playbooks) | [`04_Outreach/`](04_Outreach/) |
| **Textentwurf bearbeiten** | [`03_Entwuerfe/Teaser/`](03_Entwuerfe/Teaser/) |
| **Fakten & Briefing** (auch für KI) | [`01_Quellen/`](01_Quellen/) |
| **Video-Konzept** (Skript, Produktion) | [`02_Aktuell/Video/`](02_Aktuell/Video/) |
| **Alte Versionen** | [`06_Archiv/`](06_Archiv/) |
| **Skripte** (PPTX bauen, Extrakte) | [`99_Tools/`](99_Tools/) |

**Aktueller Teaser-Stand:** Variante E Interpack · freigegeben 02.06.2026 · PPTX **v02**

---

## Ordnerlogik (neue Nummerierung)

```text
01_Quellen     → Wahrheit für Fakten: Briefing, Extrakte, Rohdaten (PDF nur über Extrakt)
02_Aktuell     → Freigegeben & einsatzbereit (Teaser, Outreach-Kampagnen, Video)
03_Entwuerfe   → Arbeit in Progress (Markdown, nicht versenden)
04_Outreach    → LinkedIn, Playbooks, Kontaktlisten, Kampagnen
05_Referenz    → Inspiration, Layout-Research (nicht verbindlich)
06_Archiv      → Ersetzt, Duplikate, abgeschlossene Kampagnen
99_Tools       → Build-Skripte, Brand-Tokens, Migration
```

**Regel:** Nur `02_Aktuell` und aktive Dateien in `04_Outreach/Kampagnen/` sind „live“. Alles andere ist Quelle, Entwurf oder Archiv.

---

## Typischer Workflow

### Teaser aktualisieren

1. Copy in [`03_Entwuerfe/Teaser/`](03_Entwuerfe/Teaser/) ändern  
2. Briefing prüfen: [`01_Quellen/verbindlich/LFL_Teaser_Briefing_20260529.yaml`](01_Quellen/verbindlich/LFL_Teaser_Briefing_20260529.yaml)  
3. HTML/PPTX erzeugen (siehe [`02_Aktuell/Teaser/Variante_E_Interpack/00_README.md`](02_Aktuell/Teaser/Variante_E_Interpack/00_README.md))  
4. Nach Freigabe: Dateien in `02_Aktuell/Teaser/…` · alte Version nach `06_Archiv/Teaser/`

### LinkedIn / InMail

1. Playbooks: [`04_Outreach/Playbooks/`](04_Outreach/Playbooks/)  
2. Workflow & KI-Prompt: [`04_Outreach/Workflow/`](04_Outreach/Workflow/)  
3. Konkrete Kampagne: [`04_Outreach/Kampagnen/`](04_Outreach/Kampagnen/)  
4. Personalisierte Mails pro Kontakt unter `04_Outreach/Kampagnen/<datum>_<thema>/`

### KI / Cursor / Claude nutzen

1. Projektregel: [`CLAUDE.md`](CLAUDE.md)  
2. Pflichtquellen: [`01_Quellen/README.md`](01_Quellen/README.md)  
3. LinkedIn-Regel: [`.cursor/rules/lfl-linkedin-outreach.mdc`](.cursor/rules/lfl-linkedin-outreach.mdc)

**Wichtig:** Agenten lesen **keine** PDF/DOCX/XLSX direkt — nur Markdown unter `01_Quellen/extrahiert/`. Neue Rohdateien → Extrakt aktualisieren (`99_Tools/extract/`).

### Neues Kundengespräch / Input

1. Notiz oder Transkript nach `01_Quellen/roh/Kundengespraeche/YYYY-MM-DD_<thema>.md`  
2. Relevante Insights in Briefing oder Entwurf übernehmen  
3. Optional: Outreach-Kampagne unter `04_Outreach/Kampagnen/`

### Archivieren

Wenn etwas ersetzt wurde:

- Teaser → `06_Archiv/Teaser/<datum>_<name>/`  
- Kampagne abgeschlossen → `06_Archiv/Outreach/`  
- **Nicht löschen**, nur verschieben (Nachvollziehbarkeit)

---

## Was nicht in Git gehört

- `~$*.pptx` (Office-Lockdateien)  
- `.decompressed`-Reste von Handoff-Extrakten (nach Migration Archiv)  
- Leere Ordner `OUTPUTS/`, `Neuer Ordner/`

---

## Migration von alter Struktur

Abgeschlossen (2026-06). Nachzug: [`MIGRATION_ANLEITUNG_NUTZER.md`](MIGRATION_ANLEITUNG_NUTZER.md) · [`MIGRATION_PLAN.md`](MIGRATION_PLAN.md)

## Template für neue Projekte

**Fester Ort:** [`00_Projekt_Vorlage/`](00_Projekt_Vorlage/) — dort starten mit [`START_NEUES_PROJEKT.md`](00_Projekt_Vorlage/START_NEUES_PROJEKT.md)

| Inhalt | Datei/Ordner |
|--------|----------------|
| Schritt-für-Schritt | `00_Projekt_Vorlage/START_NEUES_PROJEKT.md` |
| Regeln & Baum | `00_Projekt_Vorlage/PROJEKT_VORLAGE.md` |
| **Kopierbare Ordner** | `00_Projekt_Vorlage/Vorlage_Ordnerstruktur/` |

---

## Kontakt (Teaser-CTA)

- Olaf Pick · olaf@loopforgelab.com  
- hello@loopforgelab.com  
- https://calendly.com/olaf-pick/olaf-15min
