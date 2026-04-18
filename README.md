# helpmedoc

RTM (Requirements Traceability Matrix) automation suite for Windchill/PTC PLM systems. Parses `.properties` configuration files and generates styled Excel RTM files for LCS plugins, JSP/URL mappings, CSP client-side plugins, and Techpack product specs.

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

3. Or run a specific module:

```bash
python run_all_rtm.py --only lcs
python run_all_rtm.py --only jsp csp
```

## Modules

| Module | Script | Input | Output |
|---|---|---|---|
| LCS | `ssp-rtm-sync/ssp.py` | `custom.lcs.plugins.properties` | `SSP_RTM.xlsx` |
| JSP | `jsp-rtm-sync/scripts/jsp.py` | `custom.urlMappings.properties`, `custom.activityControllerMappings.properties`, `custom.controllerAliases.properties` | `JSP_RTM.xlsx` |
| CSP | `csp-rtm-sync/scripts/csp.py` | `custom.clientSidePluginManagerMappings.properties` | `CSP_RTM.xlsx` |
| Techpack | `techpack-rtm-sync/techpack.py` | `ProductSpecification2.properties`, `ProductSpecificationBOM2.properties` | `Techpack_RTM.xlsx` |

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
| `--only lcs jsp csp techpack` | all | Run only the specified module(s) |
| `--lcs-input FILE [FILE ...]` | `properties/*.properties` | Explicit input file(s) for LCS |
| `--lcs-output FILE` | `RTM.xlsx` | LCS output path |
| `--lcs-template FILE` | — | Existing `RTM.xlsx` to reuse for formatting |
| `--csp-properties FILE` | `properties/custom.clientSidePluginManagerMappings.properties` | CSP input file |
| `--csp-rtm FILE` | `CSP_RTM.xlsx` | CSP RTM path |
| `--work-dir DIR` | project root | Working directory for JSP and Techpack |

## License

MIT
