# SCT VBA Project

## Purpose

This repository contains the VBA source code of the SCT Excel tool.

The project is used for:

- SCT calculations
- EHB processing
- Quickchecks
- File generation
- Mail generation
- Data transfer between workbooks

## Structure

### Workbook
Contains workbook events and startup logic.

### Worksheets
Contains worksheet-specific event handlers and logic.

### Modules
Contains standard VBA modules and business logic.

### Forms
Contains VBA user forms.

## Important Notes

The repository only contains exported VBA source files.

The following Excel components are NOT included:

- Named Ranges
- Formulas
- Data Validation
- Sheet Protection Settings
- Hidden Sheets
- Workbook Structure

These elements remain inside the original Excel workbook.

## Instructions for AI Analysis

Before making any changes:

1. Analyze the complete project architecture.
2. Identify all dependencies.
3. Identify workbook and worksheet events.
4. Create a call hierarchy.
5. Explain the purpose of each module.
6. Do not modify code before understanding the complete flow.
