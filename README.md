# R PowerPoint Report Generator

This project automates the creation of professional PowerPoint reports using R and the `officer` package.
It supports advanced features like formatted tables (`flextable`), editable vector graphics (`rvg`), and custom layouts (`patchwork`).

## Files in this Repository

| File | Description |
| :--- | :--- |
| **`create_template.R`** | Generates the base `template.pptx` using `officer` defaults. Use this to create a starter file you can customize. |
| **`generate_report.R`** | The main script. Generates the final `report.pptx` with plots, tables, and text. |
| **`inspect_template.R`** | A helper script to inspect any `.pptx` file to find Layout Names and Placeholder Labels. |

## Usage

### 1. Generate or Customize the Template
Run the template creation script to get a starter file:
```bash
Rscript create_template.R
```
This creates `template.pptx`.
**Tip:** Open this file in PowerPoint, edit the Master Slides (fonts, logos, colors), and save it. The report script will use this file as its design base.

### 2. Generate the Report
Run the main report generation script:
```bash
Rscript generate_report.R
```
This will produce `report.pptx` containing:
*   **Basic Plots:** Histogram, Heatmap, Pie Chart, Line Plot.
*   **Formatted Tables:** Financial data table with conditional formatting (using `flextable`).
*   **Side-by-Side Plots:** Two heatmaps combined into one slide (using `patchwork`).
*   **Dynamic Text:** Executive summary slide with calculated values.

## Advanced Features
The `generate_report.R` script demonstrates:
*   **Editable Graphics:** Plots are inserted as vector graphics (`rvg`), allowing you to ungroup and edit them in PowerPoint.
*   **Layout Management:** Combining multiple plots into a single placeholder using `patchwork` + `ggplot2`.
*   **Data-Driven Text:** Inserting R-calculated variables directly into text boxes.

## Using Your Own Template
If you want to use an existing corporate presentation as your template:

1.  Place your `.pptx` file in the project folder.
2.  Use `inspect_template.R` to identify the correct layout names and placeholder labels:
    ```bash
    Rscript inspect_template.R
    # (Edit the script first to set: target_pptx <- "my_presentation.pptx")
    ```
3.  Update `generate_report.R` to use your filename and layout names:
    ```r
    doc <- read_pptx("my_presentation.pptx")
    # ...
    doc |> add_slide(layout = "Your Layout Name", master = "Your Master Name")
    # ...
    ph_with(..., location = ph_location_label(ph_label = "Your Placeholder Label"))
    ```
