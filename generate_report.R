suppressPackageStartupMessages(library(officer))
suppressPackageStartupMessages(library(ggplot2))
suppressPackageStartupMessages(library(dplyr))
suppressPackageStartupMessages(library(flextable))
suppressPackageStartupMessages(library(rvg))
suppressPackageStartupMessages(library(patchwork))

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
pie_data <- data.frame(
  Category = c("Product A", "Product B", "Product C", "Product D"),
  Share = c(30, 20, 15, 35)
) |>
  mutate(Label = paste0(Category, "\n", Share, "%"))

# 2. Comparison Data
comp_data <- data.frame(
  Metric = c("Revenue", "Cost", "Profit", "Margin"),
  Value = c(100000, 60000, 40000, "40%")
)
heat_data_basic <- expand.grid(x = 1:10, y = 1:10) |> mutate(z = rnorm(n(), mean = x+y))

# 3. Product Data (Advanced Table)
product_data <- data.frame(
  Product = c("Alpha", "Beta", "Gamma", "Delta", "Epsilon", "Zeta", "Eta"),
  Revenue = c(15000, 23000, 8500, 31000, 12000, 18000, 22000),
  Growth = c(0.12, 0.05, -0.08, 0.22, -0.02, 0.04, 0.10) * 100
)

# 4. regional Heatmap Data (Side-by-Side)
heat_data1 <- expand.grid(Region = LETTERS[1:5], Quarter = 1:4) |>
  mutate(Sales = rnorm(n(), mean = 100, sd = 20))
heat_data2 <- expand.grid(Region = LETTERS[1:5], Quarter = 1:4) |>
  mutate(Score = runif(n(), 3, 5))


# --- Objects Creation ---

# Basic Plots
p_hist <- ggplot(hist_data, aes(x = value)) +
  geom_histogram(binwidth = 2, fill = "steelblue", color = "black") +
  labs(title = "Distribution Analysis", x = "Value")

p_heat_basic <- ggplot(heat_data_basic, aes(x = x, y = y, fill = z)) +
  geom_tile() + scale_fill_viridis_c() + labs(title = "Correlation Map")

p_pie <- ggplot(pie_data, aes(x = "", y = Share, fill = Category)) +
  geom_col(color = "white") +
  coord_polar(theta = "y") +
  geom_text(aes(label = Label), position = position_stack(vjust = 0.5), color = "white", fontface = "bold") +
  labs(title = "Market Share Split") +
  theme_void(base_size = 14) +
  theme(legend.position = "right", plot.title = element_text(hjust = 0.5, face = "bold"))

p_line <- ggplot(line_data, aes(x = Time, y = Value)) +
  geom_line(color = "darkred", linewidth = 1) +
  labs(title = "Temporal Trend", x = "Period")

# Advanced Objects
ft_comp <- flextable(comp_data) |> autofit()

ft_full <- flextable(product_data) |>
  colformat_double(j = "Revenue", digits = 0, prefix = "$") |>
  colformat_double(j = "Growth", digits = 1, suffix = "%") |>
  bold(part = "header") |>
  bg(part = "header", bg = "#4472C4") |>
  color(part = "header", color = "white") |>
  color(i = ~ Growth < 0, j = "Growth", color = "red") |>
  autofit()

p_adv1 <- ggplot(heat_data1, aes(x = Quarter, y = Region, fill = Sales)) +
  geom_tile() + scale_fill_viridis_c(option = "magma") + theme_minimal()
p_adv2 <- ggplot(heat_data2, aes(x = Quarter, y = Region, fill = Score)) +
  geom_tile() + scale_fill_viridis_c(option = "viridis") + theme_minimal()

combined_heatmaps <- p_adv1 + p_adv2 + plot_annotation(title = "Regional Sales vs Satisfaction")

# Dynamic Text
summary_text <- sprintf("Report Generated: %s\nTotal Revenue Analyzed: $%s\nTop Growth Product: %s", 
                        Sys.Date(), 
                        format(sum(product_data$Revenue), big.mark=","),
                        product_data$Product[which.max(product_data$Growth)])

# --- Report Generation ---

if (!file.exists("template.pptx")) {
  stop("template.pptx not found! Please run create_template.R first.")
}
doc <- read_pptx("template.pptx")

# 1. Title Slide
doc <- add_title_slide(doc, 
  title = "Comprehensive Business Intelligence Report", 
  subtitle = paste("Generated on", Sys.Date())
)

# 2. Executive Summary
doc <- add_summary_slide(doc, 
  title = "Executive Summary", 
  text = summary_text
)

# 3. SECTION: BASIC ANALYSIS
doc <- add_section_header(doc, title = "Basic Exploratory Analysis")

doc <- add_two_content_slide(doc,
  title = "Distribution Analysis",
  left_content = p_hist,
  right_content = c("Bell-curve distribution", "Mean centered at 50", "High statistical significance")
)

doc <- add_two_content_slide(doc,
  title = "Correlation Heatmap",
  left_content = p_heat_basic,
  right_content = c("Linear correlation observed", "Strong positive trend in diagonal", "Interactive heatmap visualization")
)

doc <- add_two_content_slide(doc,
  title = "Market Share Split",
  left_content = p_pie,
  right_content = c("Product D leads market", "Diversified product portfolio", "Quarterly growth metrics included")
)

doc <- add_two_content_slide(doc,
  title = "Trend Analysis",
  left_content = p_line,
  right_content = c("Overall positive growth", "Cumulative trend tracking", "50 time-periods analyzed")
)

# 4. SECTION: ADVANCED ANALYSIS
doc <- add_section_header(doc, title = "Advanced Performance Metrics")

doc <- add_comparison_slide(doc,
  title = "KPI vs Regional Distribution",
  left_header = "Regional Comparison (Editable)",
  right_header = "Key Performance Indicators",
  left_content = dml(ggobj = combined_heatmaps),
  right_content = ft_comp
)

  doc <- add_full_content_slide(doc,
  title = "Detailed Product Revenue & Growth",
  content = ft_full
)

# 5. Demonstration of generalized parameters
# Here we explicitly pass the layout and master, which allows using custom templates easily.
doc <- add_summary_slide(doc,
  title = "Generalized Function Demo",
  text = "This slide was added by explicitly passing layout and master names to the function.",
  layout = "Title Only",
  master = "Office Theme"
)

# Save Report
print(doc, target = "report.pptx")
cat("Consolidated comprehensive report generated successfully as 'report.pptx'.\n")
