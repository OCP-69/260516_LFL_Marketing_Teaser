# Kampagne 2 — E-Mail-Outreach (14 Firmen)

**Stand:** 2026-06-15 (vollständig recherchiert)  
**Master:** [`260615_Kampagne_2.xlsx`](../../../260615_Kampagne_2.xlsx) · Blatt `Kontakte_Email`  
**Kopie:** [`260615_Kampagne_2_mit_Kontakte.xlsx`](260615_Kampagne_2_mit_Kontakte.xlsx)

## Export-Dateien (28 Zeilen = 14 × 2)

| Datei | Format |
|-------|--------|
| [`260615_Kontakte_Email.xlsx`](260615_Kontakte_Email.xlsx) | Excel-Tabelle (TableStyleMedium2) — **empfohlen** |
| [`260615_Kontakte_Email.csv`](260615_Kontakte_Email.csv) | CSV UTF-8-BOM |

Neue Spalte: `email_variants_tested` — getestete Adress-Varianten + SMTP-Ergebnis

---

## Neu generieren

```bash
python 99_Tools/finalize_kampagne2_outreach.py
```

---

## SMTP-Stand

| Kategorie | Anzahl | Firmen |
|-----------|--------|--------|
| **smtp OK** | 10 | AUMUND, STIWA, W&H, Optima, Schubert |
| **IP-blockiert (plausibel)** | 16 | Microsoft/Outlook-Domains — kein Nachweis ungültig |
| **Routing** | 1 | Bobst CEO → gudrun.alex@bobst.com |
| **MX-Timeout** | 1 | Bielomatik (T-Systems) |

---

## Wichtige Korrekturen (Deep Research)

| Firma | Änderung |
|-------|----------|
| **Theegarten** | Domain `.de`; CEO Markus Rustler; Eng Dr. Egbert Röhm; Robert Berge ab Jun 2026 |
| **W&H** | CEO Dr. Falco Paepenmüller (nicht Nienkemper) |
| **Optima** | CEO Dr. Stefan König (seit Apr 2024) |
| **GEBHARDT** | Eng Andreas Hooge; Domain gebhardt-group.com |
| **Bielomatik** | Paper → BW Papersystems; heute Plastic Joining Römerstein |
| **EREMA** | Eng Markus Huber-Lindinger; Domain erema.at |
| **Fill** | Sondermaschinenbau, kein Verpackungs-OEM |

---

## Versand-Priorität

1. **Sofort:** AUMUND, STIWA, W&H, Optima, Schubert (smtp OK)
2. **Plausibel:** Kolbus, Theegarten, EREMA, Fill, DS, GEBHARDT, Senning, Bobst Eng
3. **Routing:** Bobst CEO — Spalte `email_body_routing`
