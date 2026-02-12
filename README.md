# R PowerPoint Report Generator

This project automates the creation of professional PowerPoint reports using R and the `officer` package.
It supports advanced features like formatted tables (`flextable`), editable vector graphics (`rvg`), and custom layouts (`patchwork`).

## Files in this Repository

| File | Description |
| :--- | :--- |
| **`create_template.R`** | Generates `template.pptx`. Can take an existing `.pptx` as an argument to strip its slides and create a "pure" template from its layouts. |
| **`generate_report.R`** | The main script. Sources `report_functions.R` and orchestrates the report generation. |
| **`report_functions.R`** | Contains generalized helper functions (`add_title_slide`, `get_slide_inventory`, etc.) to handle custom layouts and masters. |
| **`inspect_template.R`** | A helper script to inspect any `.pptx` file to find Layout Names, Master Names, and a Slide-by-Slide Inventory. |

## Usage

### 1. Generate or Customize the Template
Run the template creation script to get a starter file:
```bash
Rscript create_template.R
```
Or create a template from an existing corporate presentation:
```bash
Rscript create_template.R corporate_slides.pptx
```
This creates a slide-free `template.pptx` while preserving all branding and layouts.

### 2. Inspect a Template
Use `inspect_template.R` to identify the correct layout names and see what's inside a file:
```bash
Rscript inspect_template.R
# Or specify a file:
# Rscript inspect_template.R my_presentation.pptx
```
This outputs a layout summary and a detailed **Slide-by-Slide Inventory** (Index, Layout, Master, and Title).

### 3. Generate the Report
Run the main report generation script:
```bash
Rscript generate_report.R
```
This produces `report.pptx` with automated plots, tables, and a demonstration of generalized function parameters.

## Advanced Features
*   **Generalized Functions:** All `add_*_slide` functions in `report_functions.R` allow passing custom `layout` and `master` names.
*   **Editable Graphics:** Plots are inserted as vector graphics (`rvg`), allowing you to ungroup and edit them in PowerPoint.
*   **Clean Output:** All scripts use `suppressPackageStartupMessages()` for a clean console experience.
*   **Inventory Logic:** Reusable functions in `report_functions.R` (`get_layout_summary`, `get_slide_inventory`) for programmatic inspection.

