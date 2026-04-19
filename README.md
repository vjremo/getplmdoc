# getplmdoc

Automation suite to Report enhancements on PTC Retail PLM/FlexPLM Application. Parses `.properties` configuration files and generates styled Excel file report for SSP plugins, JSP/URL mappings, CSP client-side plugins, and Techpack.

## Prerequisites

- Python 3.10+

```bash
pip install -r requirements.txt
```

## Quickstart

1. Copy the sample properties files into `properties/` and remove the `.example` suffix:

```bash
cp properties/sample/custom.lcs.plugins.properties.example properties/custom.lcs.plugins.properties
# repeat for whichever modules you need
```

2. Run all modules:

```bash
python run_all_rtm.py
```

This generates four individual RTM files and a combined workbook `CustomizationReport.xlsx` with one sheet per module.

3. Or run a specific module:

```bash
python run_all_rtm.py --only ssp
python run_all_rtm.py --only jsp csp
```

## Modules

| Module | Script | Input | Output |
|---|---|---|---|
| SSP | `ssp-rtm-sync/ssp.py` | `custom.lcs.plugins.properties` | `SSP.xlsx` |
| JSP | `jsp-rtm-sync/scripts/jsp.py` | `custom.urlMappings.properties`, `custom.activityControllerMappings.properties`, `custom.controllerAliases.properties` | `JSP.xlsx` |
| CSP | `csp-rtm-sync/scripts/csp.py` | `custom.clientSidePluginManagerMappings.properties` | `CSP.xlsx` |
| Techpack | `techpack-rtm-sync/techpack.py` | `ProductSpecification2.properties`, `ProductSpecificationBOM2.properties` , `ProductSpecificationMeasure2.properties` | `Techpack.xlsx` |

Each module has its own `README.md` with output column definitions and properties file format details.

## Properties files

Place your `.properties` files in the `properties/` directory. That directory is gitignored — your configs will not be committed.

Template files in `properties/sample/*.properties.example` show the expected format for each module. Copy and rename them to get started.

## CLI options

```
python run_all_rtm.py [options]
```

| Option | Default | Description |
|---|---|---|
| `--only ssp jsp csp techpack` | all | Run only the specified module(s) |
| `--combined-output FILE` | `CustomizationReport.xlsx` | Path for the merged all-in-one workbook |
| `--no-combine` | — | Skip generating the combined workbook |
| `--ssp-input FILE [FILE ...]` | `properties/*.properties` | Explicit input file(s) for SSP |
| `--ssp-output FILE` | `SSP.xlsx` | SSP output path |
| `--ssp-template FILE` | — | Existing `SSP.xlsx` to reuse for formatting |
| `--csp-properties FILE` | `properties/custom.clientSidePluginManagerMappings.properties` | CSP input file |
| `--csp-rtm FILE` | `CSP.xlsx` | CSP output path |
| `--work-dir DIR` | project root | Working directory for JSP and Techpack |

## License

MIT
