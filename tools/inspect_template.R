# Source the modular functions (provides officer, dplyr)
source("tools/report_functions.R")

# Set console width to prevent table wrapping
options(width = 200)

# --- Argument Handling ---
args <- commandArgs(trailingOnly = TRUE)
target_pptx <- if (length(args) > 0) args[1] else "templates/template.pptx"

# --- Inspection Logic ---

if (!file.exists(target_pptx)) {
  stop(paste("File not found:", target_pptx))
}

doc <- read_pptx(target_pptx)

cat(paste("Inspecting:", target_pptx, "\n\n"))

# 1. List all available Layouts and Masters
cat("=== Available Layouts and Masters ===\n")
summary <- get_layout_summary(doc)
print(summary)

# 3. Slide-by-Slide Inventory
cat("\n=== Slide-by-Slide Inventory ===\n")

inventory_clean <- get_slide_inventory(doc)

if (!is.null(inventory_clean)) {
  # Ensure all columns are printed in a single row without row numbers
  if (inherits(inventory_clean, "tbl_df")) {
    print(as.data.frame(inventory_clean), row.names = FALSE, right = FALSE)
  } else {
    print(inventory_clean, row.names = FALSE, right = FALSE)
  }
} else {
  cat("No slides found in this presentation.\n")
}

cat("\nTip: Use 'inspect_template.R' to identify exactly which layouts are being used in an existing presentation.\n")
