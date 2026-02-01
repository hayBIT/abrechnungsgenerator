# Provisionsabrechnungssystem

Dieses Skript matcht Abrechnungstabellen der Versicherer KRAVAG und R+V mit der Vertragsliste aus dem CRM Ameise. Grundlage ist die Versicherungsscheinnummer (VSN):

- In der Ameise-Liste steht sie als `VSN` (Format `32-123456789` oder `32 123456789`, numerische Teile werden auf 9 Stellen aufgefüllt).
- In den Abrechnungen wird sie aus `ag` + `-` + `vsnr` gebildet.
- Falls `vsnr` weniger als 9 Stellen hat, wird links mit Nullen aufgefüllt.
- Die Zuordnung zu Vermittlern erfolgt über die Spalte `VMT`.

## Nutzung

```bash
python provision.py \
  --ameise path/to/ameise.csv \
  --kravag path/to/kravag_1.csv path/to/kravag_2.csv \
  --rv path/to/rv_1.csv path/to/rv_2.csv \
  --output-dir output
```

## Ausgabe

- `output/matched_rows.csv`: Alle Abrechnungszeilen mit angereicherter `VSN`, `VMT`, `match_status` und `insurer`.
- `output/summary_by_vmt.csv`: Zusammenfassung der Treffer nach `VMT` (inklusive `UNMATCHED`).
- `output/vmt/<VMT>.ods`: Je Vermittler eine ODS-Datei mit den Spalten `VSN`, `Vorname / Ansprechpartner`, `Nachname / Firma`, `Gesellschaft`, `Sparte`, `beg_wirk_dat`, `abrechnungsbetrag`.

## Hinweise

- Das Skript erkennt das CSV-Trennzeichen automatisch.
- Mindestens eine Abrechnung (`--kravag` oder `--rv`) muss übergeben werden (jeweils mit einem oder mehreren CSVs).
- `Vorname / Ansprechpartner`, `Nachname / Firma`, `Gesellschaft` und `Sparte` stammen aus der Ameise-CSV; `abrechnungsbetrag` wird mit `€` ausgegeben.
