library(officer)
library(ggplot2)
library(dplyr)
# Advanced libraries
library(flextable)
library(rvg)
library(patchwork)

# Standard styling for better readability
theme_set(theme_bw(base_size = 14) + 
          theme(plot.title = element_text(size = 16, face = "bold"),
                axis.title = element_text(size = 14),
                axis.text = element_text(size = 12)))

# --- Data Generation ---
set.seed(42) # Unified seed

# 1. Histogram Data
hist_data <- data.frame(value = rnorm(1000, mean = 50, sd = 10))

# 2. Heatmap Data (Basic)
heatmap_data <- expand.grid(x = 1:10, y = 1:10) |>
  mutate(z = rnorm(n(), mean = x + y, sd = 2))

# 3. Piechart Data
pie_data <- data.frame(
  Category = c("Product A", "Product B", "Product C", "Product D"),
  Share = c(30, 20, 15, 35)
) |>
  mutate(Label = paste0(Category, "\n", Share, "%"))

# 4. Line Plot Data
line_data <- data.frame(
  Time = 1:50,
  Value = cumsum(rnorm(50))
)

# 5. Advanced Product Data (for Flextable)
product_data <- data.frame(
  Product = c("Alpha", "Beta", "Gamma", "Delta", "Epsilon"),
  Revenue = c(15000, 23000, 8500, 31000, 12000),
  Growth = c(0.12, 0.05, -0.08, 0.22, -0.02) * 100
)

# 6. Advanced Heatmap Data (Side-by-Side)
heat_data1 <- expand.grid(Region = LETTERS[1:5], Quarter = 1:4) |>
  mutate(Sales = rnorm(n(), mean = 100, sd = 20))

heat_data2 <- expand.grid(Region = LETTERS[1:5], Quarter = 1:4) |>
  mutate(Score = runif(n(), 3, 5))


# --- Plot / Object Creation ---

# 1. Histogram
p_hist <- ggplot(hist_data, aes(x = value)) +
  geom_histogram(binwidth = 2, fill = "steelblue", color = "black") +
  labs(title = "Distribution Analysis", x = "Value", y = "Frequency")

# 2. Heatmap (Basic)
p_heat <- ggplot(heatmap_data, aes(x = x, y = y, fill = z)) +
  geom_tile() +
  scale_fill_viridis_c() +
  labs(title = "Correlation Heatmap", x = "Dimension X", y = "Dimension Y")

# 3. Piechart
p_pie <- ggplot(pie_data, aes(x = "", y = Share, fill = Category)) +
  geom_col(color = "white") +
  coord_polar(theta = "y") +
  geom_text(aes(label = Label), position = position_stack(vjust = 0.5), color = "white", fontface = "bold") +
  labs(title = "Market Share Split") +
  theme_void(base_size = 14) +
  theme(legend.position = "right", plot.title = element_text(hjust = 0.5, face = "bold", size=16))

# 4. Line Plot
p_line <- ggplot(line_data, aes(x = Time, y = Value)) +
  geom_line(color = "darkred", linewidth = 1) +
  geom_point(size = 2) +
  labs(title = "Time Series Trend", x = "Time Period", y = "Cumulative Value")

# 5. Flextable 
ft <- flextable(product_data) |>
  colformat_double(j = "Revenue", digits = 0, prefix = "$") |>
  colformat_double(j = "Growth", digits = 1, suffix = "%") |>
  bold(part = "header") |>
  bg(part = "header", bg = "#4472C4") |>
  color(part = "header", color = "white") |>
  color(i = ~ Growth < 0, j = "Growth", color = "red") |>
  autofit()

# 6. Side-by-Side Heatmaps (Patchwork)
p_adv1 <- ggplot(heat_data1, aes(x = Quarter, y = Region, fill = Sales)) +
  geom_tile() +
  scale_fill_viridis_c(option = "magma") +
  labs(title = "Regional Sales", fill = "Sales ($)") +
  theme_minimal(base_size = 12)

p_adv2 <- ggplot(heat_data2, aes(x = Quarter, y = Region, fill = Score)) +
  geom_tile() +
  scale_fill_viridis_c(option = "viridis") +
  labs(title = "Satisfaction Score", fill = "Score") +
  theme_minimal(base_size = 12)

combined_plot <- p_adv1 + p_adv2 + plot_annotation(title = "Q1-Q4 Performance Overview")

# 7. Dynamic Summary Text
total_revenue <- sum(product_data$Revenue)
top_performer <- product_data$Product[which.max(product_data$Growth)]
summary_text <- sprintf("Total Revenue: $%s\nTop Performer: %s (%.1f%% Growth)", 
                        format(total_revenue, big.mark=","), 
                        top_performer, 
                        max(product_data$Growth))


# --- Report Generation ---

# Load the template created by create_template.R
if (!file.exists("template.pptx")) {
  stop("template.pptx not found! Please run create_template.R first.")
}
doc <- read_pptx("template.pptx")

# Helper Function: Adds a slide with "Two Content" layout
# Accepts any 'left_content' (ggplot, flextable, dml, etc.)
add_content_slide <- function(doc, title, left_content, right_bullets) {
  doc |>
    # add_slide creates a new slide based on a specific layout and master theme.
    add_slide(layout = "Two Content", master = "Office Theme") |>
    
    # Add Title
    ph_with(value = title, location = ph_location_type(type = "title")) |>
    
    # Add Left Content (Index 1) - Can be plot, table, editable graphic
    ph_with(value = left_content, location = ph_location_label(ph_label = "Content Placeholder 2")) |>
    
    # Add Right Content (Index 2) - Bullet points
    ph_with(value = right_bullets, location = ph_location_label(ph_label = "Content Placeholder 3"))
}

# --- Basic Slides ---

# Slide 1: Histogram
doc <- add_content_slide(doc, 
  title = "Histogram Analysis (Basic)", 
  left_content = p_hist, 
  right_bullets = c(
    "Normal distribution observed",
    "Mean value is approximately 50",
    "Standard deviation is around 10"
  )
)

# Slide 2: Heatmap
doc <- add_content_slide(doc, 
  title = "Correlation Overview (Basic)", 
  left_content = p_heat, 
  right_bullets = c(
    "Visual representation of matrix data",
    "Higher values observed in top-right quadrant",
    "Correlation pattern suggests linear relationship"
  )
)

# Slide 3: Pie Chart
doc <- add_content_slide(doc, 
  title = "Category Breakdown (Basic)", 
  left_content = p_pie, 
  right_bullets = c(
    "Product D holds the largest share (35%)",
    "Combined share of A and B is 50%",
    "Market segmentation is diversified"
  )
)

# Slide 4: Line Plot
doc <- add_content_slide(doc, 
  title = "Trend Over Time (Basic)", 
  left_content = p_line, 
  right_bullets = c(
    "Cumulative growth over 50 time periods",
    "Dark red line indicates the trajectory",
    "Overall positive trend maintained"
  )
)

# --- Advanced Slides ---

# Slide 5: Financial Table (Flextable)
doc <- add_content_slide(doc, 
  title = "Financial Performance (Flextable)", 
  left_content = ft, 
  right_bullets = c(
    "Delta leads revenue with $31k",
    "Gamma and Epsilon show negative growth (highlighted in red)",
    "Table dynamically formatted in R"
  )
)

# Slide 6: Side-by-Side Editable Heatmaps (Patchwork + RVG)
doc <- add_content_slide(doc, 
  title = "Regional Analysis (Patchwork + RVG)", 
  # Use dml(ggobj = ...) to make it editable in PPT
  left_content = dml(ggobj = combined_plot), 
  right_bullets = c(
    "Left: Sales volume by region",
    "Right: Customer satisfaction scores",
    "Combined using layout package 'patchwork'",
    "Fully editable vector graphics (try ungrouping!)"
  )
)

# Slide 7: Dynamic Text Summary (Title Only Layout)
doc |>
  add_slide(layout = "Title Only", master = "Office Theme") |> 
  ph_with(value = paste("Executive Summary -", format(Sys.Date(), "%Y")), location = ph_location_type(type = "title")) |>
  ph_with(value = summary_text, location = ph_location(left = 1, top = 2, width = 8, height = 2))

# Save the unified report
print(doc, target = "report.pptx")
cat("Consolidated report generated successfully as 'report.pptx'.\n")
