library(officer)
library(ggplot2)
library(dplyr)
library(flextable)
library(rvg)
library(patchwork)

# Source the modular functions
source("report_functions.R")

# Standard styling
theme_set(theme_bw(base_size = 14) + 
          theme(plot.title = element_text(size = 16, face = "bold"),
                axis.title = element_text(size = 14),
                axis.text = element_text(size = 12)))

# --- Data Generation ---
set.seed(42)

# 1. Basic Data
hist_data <- data.frame(value = rnorm(1000, mean = 50, sd = 10))
line_data <- data.frame(Time = 1:50, Value = cumsum(rnorm(50)))

# 2. Comparison Data
comp_data <- data.frame(
  Metric = c("Revenue", "Cost", "Profit", "Margin"),
  Value = c(100000, 60000, 40000, "40%")
)
heat_data <- expand.grid(x = 1:5, y = 1:5) |> mutate(z = rnorm(n()))

# 3. Product Data (Flextable)
product_data <- data.frame(
  Product = c("Alpha", "Beta", "Gamma", "Delta", "Epsilon", "Zeta", "Eta"),
  Revenue = c(15000, 23000, 8500, 31000, 12000, 18000, 22000),
  Growth = c(0.12, 0.05, -0.08, 0.22, -0.02, 0.04, 0.10) * 100
)

# --- Objects Creation ---

p_hist <- ggplot(hist_data, aes(x = value)) +
  geom_histogram(binwidth = 2, fill = "steelblue", color = "black") +
  labs(title = "Distribution", x = "Value")

p_line <- ggplot(line_data, aes(x = Time, y = Value)) +
  geom_line(color = "darkred", linewidth = 1) +
  labs(title = "Trend", x = "Period")

p_heat <- ggplot(heat_data, aes(x = x, y = y, fill = z)) +
  geom_tile() + scale_fill_viridis_c() + theme_minimal()

ft_comp <- flextable(comp_data) |> autofit()

ft_full <- flextable(product_data) |>
  colformat_double(j = "Revenue", digits = 0, prefix = "$") |>
  colformat_double(j = "Growth", digits = 1, suffix = "%") |>
  bold(part = "header") |>
  bg(part = "header", bg = "#4472C4") |>
  color(part = "header", color = "white") |>
  color(i = ~ Growth < 0, j = "Growth", color = "red") |>
  autofit()

# Dynamic Text
summary_text <- sprintf("Report Generated: %s\nTotal Revenue: $%s", 
                        Sys.Date(), format(sum(product_data$Revenue), big.mark=","))

# --- Report Generation ---

if (!file.exists("template.pptx")) {
  stop("template.pptx not found! Please run create_template.R first.")
}
doc <- read_pptx("template.pptx")

# 1. Title Slide
doc <- add_title_slide(doc, 
  title = "Automated Business Report", 
  subtitle = "Generated via R & officer"
)

# 2. Executive Summary (Dynamic Text)
doc <- add_summary_slide(doc, 
  title = "Executive Summary", 
  text = summary_text
)

# 3. Section Header
doc <- add_section_header(doc, title = "Basic Analysis")

# 4. Two Content Slide (Histogram)
doc <- add_two_content_slide(doc,
  title = "Distribution Analysis",
  left_content = p_hist,
  right_content = c("Normal distribution", "Centered at 50", "No outliers")
)

# 5. Section Header
doc <- add_section_header(doc, title = "Advanced Analysis")

# 6. Comparison Slide (Heatmap vs Table)
doc <- add_comparison_slide(doc,
  title = "Performance vs Metrics",
  left_header = "Visual Heatmap",
  right_header = "Key Metrics Table",
  left_content = p_heat,
  right_content = ft_comp
)

# 7. Full Content Slide (Large Flextable)
doc <- add_full_content_slide(doc,
  title = "Detailed Product Performance",
  content = ft_full
)

# Save Report
print(doc, target = "report.pptx")
cat("Report generated successfully as 'report.pptx'.\n")
