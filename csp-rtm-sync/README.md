# csp-rtm-sync

A Claude Code skill that synchronizes a Client Side Plugin (CSP) Requirements Traceability Matrix (RTM) Excel file with the active entries in a Windchill/PTC `.properties` file.

## What it does

- Parses all active (non-commented) CSP mappings from a `.properties` file
- Compares against an RTM Excel file (`CSP_RTM.xlsx`)
- Adds missing entries to the RTM, preserving row formatting
- Reports what was added, what already existed, and any orphaned RTM rows (entries with no matching active mapping)

## Install

Download `csp-rtm-sync.skill` and install it in Claude Code via **Settings → Skills**.

## Requirements

- Python with `openpyxl` installed (`pip install openpyxl`)

## Usage

Say things like:
- "Sync the CSP RTM with the properties file"
- "Update documentation for our client side plugins"
- "What entries in CSP_RTM.xlsx are missing from the properties file?"

## Properties file format

```
Sample.Product.TestingRequest.CREATE_TESTING_REQUEST.CLASSIFY.onLoadEvent = /cabelas/jsp/samples/CPSRequestedByTestingSample_OnLoad.jsp
```

Format: `<FlextypePath>.<Activity>.<Action>.<clientSidePluginType> = <jspFilePath>`

## RTM format

| FlextypePath | Activity | Action | Client Side Plugin Type | Client Side Plugin JSP File |
|---|---|---|---|---|
| Sample.Product.TestingRequest | CREATE_TESTING_REQUEST | CLASSIFY | onLoadEvent | /cabelas/jsp/... |
