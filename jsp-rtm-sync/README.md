# jsp-rtm-sync

A Claude Code skill that synchronizes a JSP Requirements Traceability Matrix (RTM) Excel file from Windchill `.properties` configuration files.

## What it does

Reads three `.properties` files and writes all resolved mappings to a single sheet in `JSP_RTM.xlsx`:

| Source File | Maps |
|---|---|
| `custom.urlMappings.properties` | Page Key → JSP/resource path |
| `custom.activityControllerMappings.properties` | Activity Key → Controller Alias |
| `custom.controllerAliases.properties` | Controller Alias → Controller JSP path |

Controller entries are resolved end-to-end (`Activity Key → Alias → JSP path`) and written as `ActivityKey,ControllerAlias` in the Page Key column.

## Usage

### As a Claude Code skill

Ask Claude:
> "Read the properties files and update documentation"
> "Sync the JSP RTM"
> "What's in the RTM vs the properties files?"

Claude will invoke this skill automatically.

### As a standalone script

Run from the directory containing your `.properties` files and `JSP_RTM.xlsx`:

```bash
python scripts/update_rtm.py
```

If `JSP_RTM.xlsx` does not exist, it will be created automatically.

## Requirements

```bash
pip install openpyxl
```

## Installing as a Claude Code skill

1. Clone this repo into your Claude skills directory:
   ```bash
   cd ~/.claude/skills   # or your configured skills path
   git clone https://github.com/YOUR_USERNAME/jsp-rtm-sync
   ```
2. The skill will be available as `jsp-rtm-sync` in Claude Code.

## Output format

- Blue header row (`4472C4`), white bold Arial text
- Alternating row shading (white / light blue `DCE6F1`)
- Thin black borders on all cells
- Column A: 55 width, Column B: 65 width

## License

MIT
