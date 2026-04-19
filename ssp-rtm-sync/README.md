# ssp-rtm-sync

Convert `custom.lcs.plugins.properties` to a Requirement Traceability Matrix Excel file (`RTM.xlsx`).

## Requirements

- Python 3.10+
- openpyxl

```bash
pip install -r requirements.txt
```

## Usage

```bash
# Basic — reads custom.lcs.plugins.properties in current directory, writes RTM.xlsx
python ssp.py

# Custom input/output paths
python ssp.py path/to/custom.lcs.plugins.properties -o output/RTM.xlsx

# Use an existing RTM.xlsx as a template (preserves formatting/headers)
python ssp.py -t RTM.xlsx -o RTM_updated.xlsx
```

## Output columns

| Column | Description |
|---|---|
| Plugin Number | Numeric plugin ID (e.g. 1001) |
| Target Class | Short class name of the target object |
| Target Type | Target type scope (e.g. ALL) |
| Plugin Class | Fully qualified plugin class name |
| Plugin Method | Method called on the plugin class |
| Event | Lifecycle event trigger (e.g. POST_PERSIST) |
| Priority | Execution priority |
| By Pass Security | `True` / `False` — defaults to `False` if not specified |

## Properties file format

Each plugin entry follows this pattern:

```
com.lcs.wc.foundation.LCSPluginManager.eventPlugin.<ID>=targetClass|<class>^targetType|<type>^pluginClass|<class>^pluginMethod|<method>^event|<event>^priority|<n>[^bypassSecurity|true]
```

Lines starting with `#` are treated as comments and ignored.
