---
name: csp-rtm-sync
description: >
  Synchronize a Client Side Plugin (CSP) Requirements Traceability Matrix (RTM) Excel file with
  the active entries in a `.properties` file used for Windchill/PTC FlexType CSP mappings.
  Use this skill whenever the user asks to sync, update, reconcile, or document CSP mappings,
  mentions "update documentation" alongside a .properties or RTM file, asks what's in the RTM vs
  the properties file, or wants to add missing entries to CSP_RTM.xlsx. Also trigger when the user
  says things like "keep the RTM in sync", "document my plugins", or "what's missing from the RTM".
---

# CSP Sync

## What this skill does

Given a `.properties` file of Windchill Client Side Plugin mappings and an RTM Excel file, this
skill:

1. Parses all **active** (non-commented) entries from the `.properties` file
2. Reads the RTM Excel file
3. Adds any entries present in `.properties` but **missing** from the RTM
4. Reports what was added, what already existed, and any RTM rows with **no matching** properties
   entry (orphans — these may indicate stale or removed mappings)
5. Saves the updated RTM file

## Properties file format

Each active line follows this pattern:

```
<FlextypePath>.<Activity>.<Action>.<clientSidePluginType> = <jspFilePath>
```

- **FlextypePath**: dot-separated type hierarchy with spaces removed (e.g. `Sample.Product.TestingRequest`)
- **Activity**: ALL_CAPS action verb (e.g. `CREATE_TESTING_REQUEST`, `CREATE_DOCUMENT`)
- **Action**: e.g. `CLASSIFY`
- **clientSidePluginType**: one of `handleWidgetEvent`, `handleSubmitEvent`, `onLoadEvent`
- Lines starting with `#` or `##` are comments — skip them entirely

## RTM Excel columns (in order)

| Column | Header |
|--------|--------|
| A | FlextypePath |
| B | Activity |
| C | Action |
| D | Client Side Plugin Type |
| E | Client Side Plugin JSP File |

Row 1 is the header. Data starts at row 2.

## How to run

Use the bundled script `scripts/sync.py`. It handles all parsing, comparison, and Excel writing.

```bash
<python_exe> <skill_dir>/scripts/sync.py \
  --properties <path_to_.properties_file> \
  --rtm <path_to_RTM.xlsx>
```

Find `<python_exe>` by checking these candidates in order until one works:
- `python`
- `python3`
- `C:/Users/vjrem/AppData/Local/Python/bin/python.exe`

Find `<skill_dir>` — it is the directory containing this SKILL.md file.

The script prints a summary and exits 0 on success.

## After running

Read the script output and present a clear summary to the user:

- **Added** (N rows): list each new FlextypePath + clientSidePluginType
- **Already existed** (N rows): brief confirmation, no need to list all
- **Orphans** (N rows): list each — these are RTM rows with no matching active properties entry.
  Mention that the user may want to verify whether these should be removed.

If the script fails (e.g. missing openpyxl, file not found), diagnose the error and tell the user
what to fix before retrying.

## Edge cases

- Entries with surrounding whitespace around `=` should be trimmed
- The key to match on is the combination of FlextypePath + Activity + Action + clientSidePluginType
  (all four fields must match — same JSP path is not required for "already exists" check)
- Treat the comparison as case-sensitive
- If the RTM file does not exist yet, create a new one with just the header row, then add all
  properties entries
