library(officer)
library(dplyr)

# Set console width to prevent table wrapping
options(width = 200)

# Source the modular functions
source("report_functions.R")

# --- Configuration ---
# Change this to the path of your existing PowerPoint file
target_pptx <- "report.pptx" 

# --- Inspection Script ---

if (!file.exists(target_pptx)) {
  stop(paste("File not found:", target_pptx))
}

doc <- read_pptx(target_pptx)

cat(paste("Inspecting:", target_pptx, "\n\n"))

# 1. List all available Layouts and Masters
cat("=== Available Layouts and Masters ===\n")
summary <- get_layout_summary(doc)
print(summary)

cat("\n------------------------------------------------\n")
cat("To use a layout in add_slide(), you need the 'layout' and 'master' names from above.\n")
cat("Example: add_slide(layout = '", summary$layout[1], "', master = '", summary$master[1], "')\n", sep = "")
cat("------------------------------------------------\n\n")

# 3. Slide-by-Slide Inventory
cat("\n=== Slide-by-Slide Inventory ===\n")

inventory_clean <- get_slide_inventory(doc)

if (!is.null(inventory_clean)) {
  # Ensure all columns are printed in a single row without row numbers
  if (inherits(inventory_clean, "tbl_df")) {
    # Tibbles don't have row names, but they show row numbers in some environments.
    # Converting to data.frame and setting row.names to FALSE is a reliable way.
    print(as.data.frame(inventory_clean), row.names = FALSE, right = FALSE)
  } else {
    print(inventory_clean, row.names = FALSE, right = FALSE)
  }
} else {
  cat("No slides found in this presentation.\n")
}

cat("\nTip: Use 'inspect_template.R' to identify exactly which layouts are being used in an existing presentation.\n")

