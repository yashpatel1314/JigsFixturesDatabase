# Jigs & Fixtures Database

An Excel-based lookup tool for retrieving fixture records and photos from a centralized manufacturing database.

## Overview

This tool gives quick access to fixture metadata stored in a shared Excel workbook. Enter a fixture number and the macro pulls the associated name, owning department, usage notes, supply details, and operational status — then loads the corresponding photo directly from the network archive.

The database tracks 500+ jigs and fixtures across manufacturing operations.

## Prerequisites

- Microsoft Excel 2016 or later with macros enabled
- Network access to the company share at `\\clcrs010\CLCOvensGroup\Manufacturing\...`
- Fixture photos stored as `.HEIC` files in the archive directory (see [Photo Archive](#photo-archive))

## Workbook Structure

The workbook uses two sheets:

| Sheet | Purpose |
|-------|---------|
| `Input` | User-facing interface — enter a fixture number, view results |
| `Fixtures` | Master data table — all fixture records |

### Fixtures Sheet — Column Schema

| Column | Field | Description |
|--------|-------|-------------|
| B | Fixture Number | Unique identifier (lookup key) |
| C | Name | Fixture name or description |
| D | Department | Owning department |
| E | Used | Where or how the fixture is used |
| F | Supply | Sourcing / supply chain details |
| G | Status | Current operational status |

### Input Sheet — Cell Layout

| Cell | Role |
|------|------|
| C3 | Fixture number (input) |
| E11 | Name (output) |
| E13 | Department (output) |
| E15 | Used (output) |
| E17 | Supply (output) |
| E19 | Status (output) |
| E26 | Fixture photo (output) |

## Usage

1. Open the workbook and go to the **Input** sheet.
2. Type a fixture number into cell **C3**.
3. Click the **Appear** button — details and photo populate automatically.

If the fixture number is not found, or the photo file is missing from the archive, a message box will indicate the issue.

## Photo Archive

Photos are stored on the network at:

```
\\clcrs010\CLCOvensGroup\Manufacturing\Pub\BTP 2020\ME Projects\Jigs and Fixtures\Photo Archive\
```

Each fixture expects a subfolder named by its fixture number, with the image at:

```
{Fixture Number}\Pictures\{Fixture Number}.HEIC
```

## Configuration

To update the archive path, open the VBA editor (`Alt + F11`) and change the `Destination1` / `Destination2` variables in `main.vb`.
