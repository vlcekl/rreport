---
description: Convert a PPTX file to a Quarto-compatible QMD file with AI assistance
---
This workflow converts a PowerPoint presentation into a Quarto RevealJS presentation. 

1. List the files available in `import/pptx/` using the `list_dir` tool.
2. If there are multiple files, ask the user which one they want to convert. If there's only one, automatically select it.
// turbo
3. Run the conversion script on the selected file: `tools/convert_pptx.sh "import/pptx/<filename>"`
4. Use the `view_file` tool to inspect the generated `.qmd` file in `present/`.
5. Proactively invoke the `Refine Quarto Presentation` skill (located inside the `.agent/skills/refine_quarto_presentation/` directory) to read the skill instructions and help the user clean up the resulting markdown formatting.
