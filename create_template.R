suppressPackageStartupMessages(library(officer))

# --- Argument Handling ---
args <- commandArgs(trailingOnly = TRUE)
source_pptx <- if (length(args) > 0) args[1] else NULL

# --- Template Creation ---

if (!is.null(source_pptx)) {
  if (!file.exists(source_pptx)) {
    stop(paste("Source file not found:", source_pptx))
  }
  cat(paste("Creating template from existing file:", source_pptx, "\n"))
  doc <- read_pptx(source_pptx)
  
  # Remove all existing slides to make it a "pure" template
  num_slides <- length(doc)
  if (num_slides > 0) {
    cat(paste("Removing", num_slides, "slides...\n"))
    while (length(doc) > 0) {
      doc <- remove_slide(doc, index = 1)
    }
  }
} else {
  cat("Creating template from default PowerPoint theme.\n")
  doc <- read_pptx()
}

# Print available layouts to verify
cat("\nAvailable Layouts in Template:\n")
print(layout_summary(doc))

# Save the template
print(doc, target = "template.pptx")
cat("\nTemplate saved as 'template.pptx'.\n")

