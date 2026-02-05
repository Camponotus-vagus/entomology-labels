# Entomology Labels Generator

[![Python 3.9+](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Genera etichette professionali per specimen entomologici con supporto per molteplici formati di input e output.

## Caratteristiche

- **Interfaccia Grafica (GUI)**: Interfaccia intuitiva per creare e gestire etichette
- **Interfaccia Linea di Comando (CLI)**: Per automazione e scripting
- **Formati di Input Multipli**: Excel (.xlsx, .xls), CSV, TXT, Word (.docx), JSON, YAML
- **Formati di Output Multipli**: HTML, PDF, Word (.docx)
- **Layout Configurabile**: 10x13 etichette per pagina (default A4), completamente personalizzabile
- **Generazione Sequenziale**: Crea serie di etichette con codici incrementali (N1, N2, N3...)

## Formato Etichetta

Ogni etichetta contiene:
```
Italia, Trentino Alto Adige,        ← Riga 1: Località principale
Giustino (TN), Vedretta d'Amola     ← Riga 2: Località secondaria
                                     ← Riga vuota
N1                                   ← Codice specimen
15.vi.2024                          ← Data raccolta
```

## Installazione

### Opzione 1: pip (consigliato)

```bash
# Installazione base
pip install entomology-labels

# Con supporto completo (tutti i formati)
pip install entomology-labels[all]
```

### Opzione 2: Da sorgente

```bash
git clone https://github.com/Camponotus-vagus/entomology-labels.git
cd entomology-labels
pip install -e .[all]
```

### Dipendenze Opzionali

| Feature | Pacchetti | Installazione |
|---------|-----------|---------------|
| Excel | pandas, openpyxl | `pip install entomology-labels[excel]` |
| Word | python-docx | `pip install entomology-labels[docx]` |
| PDF | weasyprint | `pip install entomology-labels[pdf]` |
| YAML | pyyaml | `pip install entomology-labels[yaml]` |
| Tutto | - | `pip install entomology-labels[all]` |

> **Nota**: Per la generazione PDF, weasyprint richiede dipendenze di sistema aggiuntive. Consulta la [documentazione weasyprint](https://doc.courtbouillon.org/weasyprint/stable/first_steps.html).

## Utilizzo

### Interfaccia Grafica (GUI)

```bash
entomology-labels-gui
```

O da Python:

```python
from entomology_labels.gui import main
main()
```

La GUI permette di:
- Aggiungere etichette manualmente
- Importare dati da file
- Configurare layout e dimensioni
- Visualizzare anteprima
- Esportare in HTML, PDF, DOCX

### Linea di Comando (CLI)

```bash
# Genera etichette da file Excel a HTML
entomology-labels generate dati.xlsx -o etichette.html

# Genera etichette da CSV a PDF
entomology-labels generate dati.csv -o etichette.pdf

# Genera etichette da JSON a Word
entomology-labels generate dati.json -o etichette.docx

# Con layout personalizzato
entomology-labels generate dati.xlsx -o etichette.html --rows 12 --cols 15

# Apri il file dopo la generazione
entomology-labels generate dati.xlsx -o etichette.html --open
```

#### Generazione Sequenziale

```bash
entomology-labels sequence \
  --location1 "Italia, Trentino Alto Adige," \
  --location2 "Giustino (TN), Vedretta d'Amola" \
  --prefix N --start 1 --end 50 \
  --date "15.vi.2024" \
  -o etichette.html
```

#### Creare Template

```bash
# Crea template JSON
entomology-labels template miei_dati.json

# Crea template Excel
entomology-labels template miei_dati.xlsx --format excel

# Crea template CSV
entomology-labels template miei_dati.csv --format csv
```

### API Python

```python
from entomology_labels import LabelGenerator, Label, LabelConfig
from entomology_labels import load_data, generate_html, generate_pdf, generate_docx

# Configurazione layout
config = LabelConfig(
    labels_per_row=10,
    labels_per_column=13,
    font_size_pt=6,
)

# Crea generatore
generator = LabelGenerator(config)

# Aggiungi etichette manualmente
label = Label(
    location_line1="Italia, Trentino Alto Adige,",
    location_line2="Giustino (TN), Vedretta d'Amola",
    code="N1",
    date="15.vi.2024"
)
generator.add_label(label)

# Oppure carica da file
labels = load_data("dati.xlsx")
generator.add_labels(labels)

# Genera output
generate_html(generator, "etichette.html", open_in_browser=True)
generate_pdf(generator, "etichette.pdf")
generate_docx(generator, "etichette.docx")
```

#### Generazione Sequenziale

```python
from entomology_labels import LabelGenerator

generator = LabelGenerator()

# Genera N1 fino a N50
labels = generator.generate_sequential_labels(
    location_line1="Italia, Trentino Alto Adige,",
    location_line2="Giustino (TN), Vedretta d'Amola",
    code_prefix="N",
    start_number=1,
    end_number=50,
    date="15.vi.2024"
)
generator.add_labels(labels)
```

## Formati File di Input

### Excel (.xlsx, .xls)

Crea un foglio con le colonne:

| location_line1 | location_line2 | code | date | count |
|----------------|----------------|------|------|-------|
| Italia, Trentino Alto Adige, | Giustino (TN), Vedretta d'Amola | N1 | 15.vi.2024 | 5 |
| Italia, Lombardia, | Sondrio, Valmalenco | O1 | 20.vii.2024 | 3 |

La colonna `count` è opzionale e serve per duplicare le etichette.

### CSV

```csv
location_line1,location_line2,code,date,count
"Italia, Trentino Alto Adige,","Giustino (TN), Vedretta d'Amola",N1,15.vi.2024,5
"Italia, Lombardia,","Sondrio, Valmalenco",O1,20.vii.2024,3
```

### JSON

```json
{
  "labels": [
    {
      "location_line1": "Italia, Trentino Alto Adige,",
      "location_line2": "Giustino (TN), Vedretta d'Amola",
      "code": "N1",
      "date": "15.vi.2024",
      "count": 5
    }
  ]
}
```

### TXT (formato chiave-valore)

```
location1: Italia, Trentino Alto Adige,
location2: Giustino (TN), Vedretta d'Amola
code: N1
date: 15.vi.2024
count: 5

location1: Italia, Lombardia,
location2: Sondrio, Valmalenco
code: O1
date: 20.vii.2024
count: 3
```

### Nomi Colonne Alternativi

Il software riconosce vari nomi per le colonne:

| Campo | Nomi accettati |
|-------|----------------|
| location_line1 | location1, località1, location, loc1 |
| location_line2 | location2, località2, loc2 |
| code | specimen_code, codice, id |
| date | collection_date, data, data_raccolta |
| count | quantity, quantità, n, copies |

## Configurazione Layout

### Parametri Disponibili

| Parametro | Default | Descrizione |
|-----------|---------|-------------|
| labels_per_row | 10 | Etichette per riga |
| labels_per_column | 13 | Etichette per colonna |
| label_width_mm | 21.0 | Larghezza etichetta (mm) |
| label_height_mm | 22.85 | Altezza etichetta (mm) |
| page_width_mm | 210.0 | Larghezza pagina (mm) |
| page_height_mm | 297.0 | Altezza pagina (mm) |
| font_size_pt | 6.0 | Dimensione font (pt) |
| font_family | Arial | Famiglia font |

### Preimpostazioni

- **A4 Standard**: 10x13 etichette (130 per pagina)
- **A4 Compatto**: 12x15 etichette (180 per pagina)
- **Letter US**: 10x12 etichette (120 per pagina)

## Output HTML e Stampa PDF

L'output HTML include un pulsante "Stampa" che apre la finestra di stampa del browser. Per salvare come PDF:

1. Genera il file HTML
2. Aprilo nel browser
3. Clicca "Stampa" o usa Ctrl+P (Cmd+P su Mac)
4. Seleziona "Salva come PDF" come destinazione
5. Assicurati che i margini siano impostati su "Nessuno"

## Esempi

La cartella `examples/` contiene file di esempio:

- `example_labels.json` - Formato JSON
- `example_labels.csv` - Formato CSV
- `example_labels.txt` - Formato TXT

## Risoluzione Problemi

### Errore "weasyprint not found"

```bash
# Ubuntu/Debian
sudo apt-get install libpango-1.0-0 libpangocairo-1.0-0 libgdk-pixbuf2.0-0

# macOS
brew install pango

# Poi installa weasyprint
pip install weasyprint
```

### Errore "pandas not found"

```bash
pip install pandas openpyxl
```

### Le etichette non si allineano correttamente

- Verifica che i margini della stampante siano impostati su 0
- Usa la modalità "Adatta alla pagina" se necessario
- Prova a regolare i parametri `margin_*` nella configurazione

## Contribuire

Contributi sono benvenuti! Per contribuire:

1. Fork del repository
2. Crea un branch per la feature (`git checkout -b feature/NuovaFeature`)
3. Commit delle modifiche (`git commit -m 'Aggiunge NuovaFeature'`)
4. Push al branch (`git push origin feature/NuovaFeature`)
5. Apri una Pull Request

## Licenza

Questo progetto è rilasciato sotto licenza MIT. Vedi il file [LICENSE](LICENSE) per i dettagli.

## Crediti

Ispirato da [insect-labels](https://github.com/tracyyao27/insect-labels) di Tracy Yao.
