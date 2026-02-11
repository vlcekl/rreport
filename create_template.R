library(officer)
library(dplyr)

# Create a new PowerPoint presentation (loads default template)
doc <- read_pptx()

# Print available layouts to verify "Two Content" exists
cat("Available Layouts:\n")
print(layout_summary(doc))

# Check properties for "Two Content" layout to identify placeholders
# We need to know which placeholder is on the left and which is on the right.
cat("\n'Two Content' Layout Properties:\n")
layout_props <- layout_properties(doc, layout = "Two Content")
print(layout_props)

# Save the template
print(doc, target = "template.pptx")
cat("\nTemplate saved as 'template.pptx'.\n")
