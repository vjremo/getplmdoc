# techpack-rtm-sync

Convert Windchill Techpack product specification `.properties` files to a Requirements Traceability Matrix Excel file (`Techpack_RTM.xlsx`).

## Requirements

- Python 3.10+
- openpyxl

```bash
pip install -r requirements.txt
```

## Usage

Run from the project root (the script resolves input paths relative to a `properties/` directory in the working directory):

```bash
python techpack-rtm-sync/techpack.py
```

The script reads both `ProductSpecification2.properties` and `ProductSpecificationBOM2.properties` in a single pass and writes all entries to one sheet.

## Input files

| File | Description |
|---|---|
| `properties/ProductSpecification2.properties` | Techpack PDF section definitions |
| `properties/ProductSpecificationBOM2.properties` | BOM-related section definitions |

See `properties/sample/*.properties.example` for the expected format.

## Properties file format

Standard Java `.properties` key-value format:

```
# This is a comment — ignored
HEADER=com.xyz.wc.product.XYZPDFProductSpecificationHeader
Construction=com.lcs.wc.product.PDFProductSpecificationConstruction2|orientation=LANDSCAPE
```

- Lines starting with `#` are treated as comments and skipped
- Entries whose value contains `com.lcs` are filtered out (LCS-internal references are excluded from the RTM)

## Output

Writes `Techpack_RTM.xlsx` in the working directory. If the file already exists, the data rows are cleared and rewritten while preserving the header row.

### Output columns

| Column | Description |
|---|---|
| Property Key | The key from the `.properties` file (e.g. `Construction`) |
| Class | The value from the `.properties` file (e.g. `com.xyz.wc.product...`) |

### Formatting

- Blue header row (`4472C4`), white bold Arial text
- Alternating row shading (white / light blue `DCE6F1`)
- Thin black borders on all cells
- Column A: 45 width, Column B: 65 width

## License

MIT
