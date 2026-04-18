---
name: jsp-rtm-sync
description: "Synchronize a JSP Requirements Traceability Matrix (RTM) Excel file with URL mappings and activity controller mappings from Windchill `.properties` files. Use this skill whenever the user asks to sync, update, reconcile, or document JSP mappings, mentions 'update documentation' alongside a .properties or RTM file, asks what's in the RTM vs the properties files, or wants to add missing entries to JSP_RTM.xlsx. Also trigger when the user says things like 'keep the RTM in sync', 'document my JSP mappings', or 'what's missing from the RTM'."
license: MIT
---

# JSP RTM Sync Skill

## Overview

This skill synchronizes a JSP Requirements Traceability Matrix (RTM) Excel file (`JSP_RTM.xlsx`) from three Windchill `.properties` source files:

| Properties File | Content |
|---|---|
| `custom.urlMappings.properties` | Page Key → JSP/resource file path |
| `custom.activityControllerMappings.properties` | Activity Key → Controller Alias |
| `custom.controllerAliases.properties` | Controller Alias → Controller JSP path |

## Output Format

All entries are written to **a single sheet** in `JSP_RTM.xlsx` with two columns:

- **Page Key** — the mapping key (URL mapping keys as-is; controller entries formatted as `ActivityKey,ControllerAlias`)
- **JSP File** — the resolved file path

## How to Run

Use the provided `scripts/update_rtm.py` helper to sync the RTM:

```bash
python scripts/update_rtm.py
```

The script must be run from the directory containing the `.properties` files and `JSP_RTM.xlsx`.

## Workflow

1. Parse all three `.properties` files (skip comment lines starting with `//` or `#`)
2. Resolve activity controller entries: chain `activityControllerMappings` through `controllerAliases` to get the JSP path
3. Merge URL mapping entries and resolved controller entries into one list
4. Write all entries to `JSP_RTM.xlsx` Sheet1, replacing previous data
5. Apply consistent formatting: blue header row, alternating row shading, Arial font, bordered cells

## Formatting Standards

- Header: Arial Bold, white text, blue fill (`4472C4`)
- Even rows: light blue fill (`DCE6F1`)
- Odd rows: white fill (`FFFFFF`)
- Column A width: 55, Column B width: 65
- All cells have a thin black border

## When Properties Files Are Missing

- If `custom.activityControllerMappings.properties` or `custom.controllerAliases.properties` are absent, skip controller entries and only write URL mappings.
- If `JSP_RTM.xlsx` does not exist, create it from scratch with the correct headers.
- If a controller alias has no matching entry in `controllerAliases.properties`, write an empty string for the JSP File cell.
