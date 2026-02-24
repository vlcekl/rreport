# RReport: Presentation Design & Generation

This project automates the creation of professional PowerPoint reports and presentations. It combines a Markdown-based authoring workflow via Quarto (`.qmd`) with robust R-based utility scripts (`officer` package) for template manipulation and data-heavy slide generation.

## Design Principles

We prioritize lean presentation design. For full guidelines, see the automated rules in the `.agent/rules/` directory. Core principles include:
- **Clarity over embellishment**: Keep slides lean.
- **Assertion-Evidence Structure**: Slide titles must be assertions (complete sentences making a point) supported by visual evidence.
- **Visual hierarchy**: Direct the viewer's eye to the critical information.

## Project Structure

- `present/`: Contains the Quarto authoring files (`.qmd`), Quarto configurations (`_quarto.yml`), and presentation-specific assets.
- `import/pptx/`: A drop-in folder for raw `.pptx` files that need to be converted to Quarto Markdown.
- `tools/`: Contains R utility scripts for generating templates, and bash scripts for file conversion/migration.
- `templates/`: Contains `.pptx` template files used as the foundation by both Quarto and the R utilities.

## Workflows and Usage

### 1. Presentation Authoring (Quarto)
To quickly draft and preview the main presentation in `present/presentation.qmd`:
```bash
quarto preview present/presentation.qmd
```
*Note*: This relies on `templates/template.pptx` for structural formatting.

### 2. Template Utilities (R)
The scripts in `tools/` manage the underlying templates used across this project.

**Generate a Clean Template**
Create a base `templates/template.pptx` from scratch, or strip an existing corporate presentation to create a layout-only template:
```bash
Rscript tools/create_template.R
# Or with an existing file:
# Rscript tools/create_template.R my_corporate_slides.pptx
```

**Inspect a Template**
Identify the correct Layout Names and Master Names within any template. This is crucial for both Quarto YAML configs and R generation scripts.
```bash
Rscript tools/inspect_template.R templates/template.pptx
```

### 3. Programmatic Report Generation (R)
For data-heavy deliverables, run the main report generation script. This produces `templates/report.pptx` with automated plots and tables:
```bash
Rscript tools/generate_report.R
```

### 4. Converting PPTX to Quarto
If you have an existing PowerPoint that you want to move into Quarto:
1. Drop the file into `import/pptx/`.
2. Run the conversion script (it will automatically find the newest file, set up a python `.venv`, extract images, and write the `.qmd` to the `present/` folder):
   ```bash
   ./tools/convert_pptx.sh
   ```
3. After editing the `.qmd` file, you can easily clean up any unused extracted images by running:
   ```bash
   python3 tools/clean_assets.py
   ```

*Authoring Tip:* If editing generated `.qmd` files in VS Code, utilize the included workspace snippets! Type `qcol2`, `qnote`, or `qinc` to quickly scaffold layouts.

## Advanced R Features
* **Generalized Functions:** All `add_*_slide` functions in `tools/report_functions.R` allow passing custom `layout` and `master` names.
* **Editable Graphics:** Plots are inserted as vector graphics (`rvg`), allowing you to ungroup and edit them in natively via PowerPoint.

---

## AI Integration (Antigravity & Copilot)
This repository includes heavily defined AI context rules to maintain design standards.

**If using Antigravity:**
The rules, skills, and slash commands are located in `.agent/`. You can use the `/convert_pptx` workflow directly in the chat to automate the conversion and formatting cleanup of PPTX files.

**If using GitHub Copilot:**
Copilot requires a different file structure. To sync the AI instructions so Copilot understands them:
```bash
./tools/migrate_agent_to_copilot.sh
```
This safely compiles the rules into `.github/copilot-instructions.md` and copies the relevant skills to `.github/skills/` for easy referencing in your Copilot Chat.
