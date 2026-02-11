library(officer)
library(dplyr)

# --- Configuration ---
# Change this to the path of your existing PowerPoint file
target_pptx <- "template.pptx" 

# --- Inspection Script ---

if (!file.exists(target_pptx)) {
  stop(paste("File not found:", target_pptx))
}

doc <- read_pptx(target_pptx)

cat(paste("Inspecting:", target_pptx, "\n\n"))

# 1. List all available Layouts and Masters
cat("=== Available Layouts and Masters ===\n")
summary <- layout_summary(doc)
print(summary)

cat("\n------------------------------------------------\n")
cat("To use a layout in add_slide(), you need the 'layout' and 'master' names from above.\n")
cat("Example: add_slide(layout = '", summary$layout[1], "', master = '", summary$master[1], "')\n", sep = "")
cat("------------------------------------------------\n\n")

# 2. Detailed Placeholder View (Optional)
# If you know the layout name you want to use, change it here to see its placeholders (index, label, type)
target_layout <- "Two Content" # Change this to your desired layout name

if (target_layout %in% summary$layout) {
  cat(paste("=== Placeholders for layout:", target_layout, "===\n"))
  props <- layout_properties(doc, layout = target_layout)
  
  # Select relevant columns for readability
  props_clean <- props |> 
    select(label = ph_label, type, index = id, master_name)
    
  print(props_clean)
  
  cat("\nTip: Use 'ph_location_label(ph_label = ...)' to target specific placeholders.\n")
} else {
  cat(paste("Layout '", target_layout, "' not found in this presentation.\n", sep=""))
}
