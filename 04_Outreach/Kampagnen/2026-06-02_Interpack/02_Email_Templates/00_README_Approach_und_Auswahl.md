# E-Mail-Outreach — Approach & Template-Auswahl (Interpack Extended)

**Stand:** 2026-06-09  
**Zielgruppe:** 19 Firmen aus [`260609_Email_Kontakte_Interpack_Extended_v2.csv`](../260609_Email_Kontakte_Interpack_Extended_v2.csv)  
**Prio-A:** [`260609_Email_Kontakte_PrioA_gt200.csv`](../260609_Email_Kontakte_PrioA_gt200.csv)

---

## Wo liegen die finalen Recherche-Ergebnisse?

| Pfad | Inhalt |
|------|--------|
| [`04_Outreach/Kampagnen/2026-06-02_Interpack/`](../) | Kampagnen-Root |
| [`260609_Email_Kontakte_Interpack_Extended_v2.md`](../260609_Email_Kontakte_Interpack_Extended_v2.md) | **Finale Kontaktliste** (19 Firmen, Quellen, Status) |
| [`260609_Email_Kontakte_Interpack_Extended_v2.csv`](../260609_Email_Kontakte_Interpack_Extended_v2.csv) | CSV für Mail-Merge / CRM |
| [`260609_Email_Kontakte_PrioA_gt200.md`](../260609_Email_Kontakte_PrioA_gt200.md) | Prio A + >200 MA (Handtmann, Bausch+Ströbel) |
| [`260609_Email_Kontakte_PrioA_gt200.csv`](../260609_Email_Kontakte_PrioA_gt200.csv) | CSV Prio-A |
| [`00_README_Insights_und_InMails.md`](../00_README_Insights_und_InMails.md) | Kampagnen-Übersicht (Teaser, LinkedIn) |
| [`02_Email_Templates/`](./) | **Dieser Ordner** — Approach + E-Mail-Templates |

---

## Empfohlener Approach (E-Mail, nicht LinkedIn)

### Grundprinzip: Discovery vor Pitch

Aus [`Quellen/08_Kundensprache_Discovery_Playbook.md`](Quellen/08_Kundensprache_Discovery_Playbook.md):

| Touch | Ziel | E-Mail erlaubt | Noch zu früh |
|-------|------|----------------|--------------|
| **1** Erstkontakt | Dialog öffnen | 1 Fakt + 1 offene Frage | Calendly, Teaser-PDF, Pilot, Feature-Liste |
| **2** Follow-up (5–7 T) | Vertiefung / Soft-Ping | Eine Nachfrage | „Sie haben nicht geantwortet“ |
| **3** | Telefon-Einladung | „15–20 Min — Sie nennen Zeit“ | Harte Demo |
| **4** | Calendly optional | Nur nach Interesse / Touch 3 | Druck „nur noch X Plätze“ |
| **5** | Teaser | Nur wenn er Interesse signalisiert | Kalt-Anhang |

**Warum:** E-Mail an GF/CTO im Maschinenbau mit Calendly im Erstkontakt wirkt wie Sales-Automation (Playbook-Fall Balensiefer). LinkedIn-InMails in `01_InMails_Interpack_Personalisiert.md` enthielten Calendly — für **E-Mail Touch 1** bewusst **entfernt** (siehe `Einsatzbereit_Email/`).

### Segment-Sprache

Zielunternehmen = überwiegend **Verpackungsmaschinenbau** (interpack-Liste).  
→ Schmerzpunkte aus [`Quellen/09_Schmerzpunkte_Verpackung_vs_ETO.md`](Quellen/09_Schmerzpunkte_Verpackung_vs_ETO.md) Variante **(V)** nutzen: RFQ nach Messe, Formatwechsel, PPWR, Export-Tender.

### Rollen-Mapping (je 2 Kontakte pro Firma)

| `contact_type` in CSV | Template | Lens |
|----------------------|----------|------|
| CEO/GF | `Einsatzbereit_Email/01_Touch1_CEO_GF.md` | Kapazität vs. Marge |
| Engineering / CTO / VP Tech | `Einsatzbereit_Email/02_Touch1_Engineering.md` | Freigabe, Varianten, Nacharbeit |
| Konzern ohne persönliche Mail | `Einsatzbereit_Email/05_Konzern_ueber_Presse.md` | Routing über `press@` / `presse@` |

### Priorisierung ICP (Reihenfolge Versand)

1. **Prio A + ETO-Fit:** Handtmann, Bausch+Ströbel, KÖRA-PACKMAT, Schubert, W+D, Höfliger, KHS  
2. **Mittelstand OEM:** Syntegon, Romaco, Multivac, Uhlmann  
3. **Konzern** (nur über Presse oder nach SMTP-Check): Krones, GEA, Sidel, IMA, TOMRA, Beumer  
4. **Aussetzen:** PolTech (Firma nicht identifiziert), SMI (CRM-Personen korrigiert → Paolo Nava)

### Technik vor Versand

- SMTP-Check aller `abgeleitet`-Adressen  
- Betreff: konkret, kein „Partnership“ / kein „KI-Lösung“  
- Signatur: Olaf Pick · LoopForgeLab · olaf@loopforgelab.com  
- Teaser (`Variante E Interpack.html`) **erst Touch 5** oder auf explizite Anfrage

---

## Ausgewählte Templates — und warum

### ✅ Übernommen (in `Quellen/` kopiert)

| Datei | Warum ausgewählt |
|-------|------------------|
| [`10_LinkedIn_Templates_Touch0-4_nach_Rolle.md`](Quellen/10_LinkedIn_Templates_Touch0-4_nach_Rolle.md) | **Hauptquelle.** Touch 1 GF + Produktion/Engineering ohne Pitch; Touch 2–3 Follow-up; explizit **kein Calendly vor Touch 3**. Struktur 1 Fakt + 1 Frage passt zu E-Mail-Erstkontakt. |
| [`08_Kundensprache_Discovery_Playbook.md`](Quellen/08_Kundensprache_Discovery_Playbook.md) | **Approach-Regeln** (Touch-Stufen, was Interesse auslöst vs. überrumpelt). Verbindliche Leitplanke für alle Mails. |
| [`09_Schmerzpunkte_Verpackung_vs_ETO.md`](Quellen/09_Schmerzpunkte_Verpackung_vs_ETO.md) | **Segment-Sprache (V)** für interpack-Firmen — RFQ-Flut nach Messe, PPWR, FMCG-Spezifikationen. |
| [`01_InMails_Interpack_Personalisiert.md`](Quellen/01_InMails_Interpack_Personalisiert.md) | **Branchen-Fakten** (80 % Kosten im Design, Datenschatz, interpack-Bezug). Als Ideenpool für `[Fakt]` — Calendly in E-Mail-Versionen gestrichen. |
| [`10b_LinkedIn_CopyPaste_Kurz.md`](Quellen/10b_LinkedIn_CopyPaste_Kurz.md) | Kurzvarianten Touch 2/3 zum schnellen Follow-up. |
| [`07_Gespraechsleitfaden_Telefon.md`](Quellen/07_Gespraechsleitfaden_Telefon.md) | Nach positivem Touch 3 — Discovery-Fragen am Telefon. |

### ❌ Nicht übernommen (und warum)

| Datei | Grund |
|-------|-------|
| `04/05/06_FollowUp_Vorschlag_A/B/C.md` | Meypack-spezifisch, LinkedIn-DM, **Calendly in Touch 2** — widerspricht E-Mail-Approach |
| `KW22/01_Templates_InMail_FollowUp.md` | InMail-Credits, **Calendly im Erstkontakt**, zu pitch-lastig für kalte E-Mail |
| `260521_LI_InMail_Templates.md` | Rohe InMail-Varianten mit Pilot/Calendly — für E-Mail Touch 1 ungeeignet |
| `01_InMails` unverändert als Versandtext | Enthält Calendly — nur als Fakt-Quelle, nicht 1:1 für E-Mail |

---

## Einsatzbereite E-Mail-Texte

### Warm — interpack / Mouaz (empfohlen für Extended-Liste)

| Ordner | Wann |
|--------|------|
| [`Einsatzbereit_Email_Interpack_Warm/`](Einsatzbereit_Email_Interpack_Warm/) | **Messe-Follow-up:** Mouaz-Bezug, Modell vorstellen, Termin in Touch 1 |

### Cold — ohne vorherigen Messekontakt

| Datei | Wann |
|-------|------|
| [`Einsatzbereit_Email/01_Touch1_CEO_GF.md`](Einsatzbereit_Email/01_Touch1_CEO_GF.md) | Erstkontakt CEO/GF (19 Firmen × Kontakt 1) |
| [`Einsatzbereit_Email/02_Touch1_Engineering.md`](Einsatzbereit_Email/02_Touch1_Engineering.md) | Erstkontakt Engineering-Leitung (19 Firmen × Kontakt 2) |
| [`Einsatzbereit_Email/03_Touch2_FollowUp_Stille.md`](Einsatzbereit_Email/03_Touch2_FollowUp_Stille.md) | 5–7 Tage ohne Antwort |
| [`Einsatzbereit_Email/04_Touch3_Telefon_Einladung.md`](Einsatzbereit_Email/04_Touch3_Telefon_Einladung.md) | Nach sachlicher Antwort oder bereitem Dialog |
| [`Einsatzbereit_Email/05_Konzern_ueber_Presse.md`](Einsatzbereit_Email/05_Konzern_ueber_Presse.md) | Krones, GEA, Multivac, TOMRA, IMA, Sidel — Routing-Adressen |

---

*LoopForgeLab Outreach · Forge Engine Design-Partner Q2 2026*
