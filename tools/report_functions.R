suppressPackageStartupMessages(library(officer))
suppressPackageStartupMessages(library(dplyr))

#' Add a Title Slide
#' @param doc officer::rdocx object
#' @param title Main title string
#' @param subtitle Subtitle string
#' @param layout Name of the layout to use (default: "Title Slide")
#' @param master Name of the master to use (default: "Office Theme")
add_title_slide <- function(doc, title, subtitle = "", layout = "Title Slide", master = "Office Theme") {
  doc |>
    add_slide(layout = layout, master = master) |>
    ph_with(value = title, location = ph_location_type(type = "ctrTitle")) |>
    ph_with(value = subtitle, location = ph_location_type(type = "subTitle"))
}

#' Add a Section Header Slide
#' @param doc officer::rdocx object
#' @param title Section title string
#' @param layout Name of the layout to use (default: "Section Header")
#' @param master Name of the master to use (default: "Office Theme")
add_section_header <- function(doc, title, layout = "Section Header", master = "Office Theme") {
  doc |>
    add_slide(layout = layout, master = master) |>
    ph_with(value = title, location = ph_location_type(type = "title"))
}

#' Add a Two Content Slide (Chart Left, Text Right)
#' @param doc officer::rdocx object
#' @param title Slide title
#' @param left_content Object for left placeholder (plot, table, text)
#' @param right_content Object for right placeholder (plot, table, text)
#' @param layout Name of the layout to use (default: "Two Content")
#' @param master Name of the master to use (default: "Office Theme")
#' @param left_ph_label Label for left content placeholder
#' @param right_ph_label Label for right content placeholder
add_two_content_slide <- function(doc, title, left_content, right_content, 
                                 layout = "Two Content", master = "Office Theme",
                                 left_ph_label = "Content Placeholder 2",
                                 right_ph_label = "Content Placeholder 3") {
  doc |>
    add_slide(layout = layout, master = master) |>
    ph_with(value = title, location = ph_location_type(type = "title")) |>
    ph_with(value = left_content, location = ph_location_label(ph_label = left_ph_label)) |>
    ph_with(value = right_content, location = ph_location_label(ph_label = right_ph_label))
}

#' Add a Comparison Slide (Title + Left/Right Headers + Left/Right Content)
#' @param doc officer::rdocx object
#' @param title Slide title
#' @param left_header Text for left header
#' @param right_header Text for right header
#' @param left_content Object for left body
#' @param right_content Object for right body
#' @param layout Name of the layout to use (default: "Comparison")
#' @param master Name of the master to use (default: "Office Theme")
#' @param left_head_ph_label Label for left header placeholder
#' @param right_head_ph_label Label for right header placeholder
#' @param left_cont_ph_label Label for left content placeholder
#' @param right_cont_ph_label Label for right content placeholder
add_comparison_slide <- function(doc, title, left_header, right_header, left_content, right_content,
                                layout = "Comparison", master = "Office Theme",
                                left_head_ph_label = "Text Placeholder 2",
                                right_head_ph_label = "Text Placeholder 4",
                                left_cont_ph_label = "Content Placeholder 3",
                                right_cont_ph_label = "Content Placeholder 5") {
  doc |>
    add_slide(layout = layout, master = master) |>
    ph_with(value = title, location = ph_location_type(type = "title")) |>
    # Headers
    ph_with(value = left_header, location = ph_location_label(ph_label = left_head_ph_label)) |>
    ph_with(value = right_header, location = ph_location_label(ph_label = right_head_ph_label)) |>
    # Bodies
    ph_with(value = left_content, location = ph_location_label(ph_label = left_cont_ph_label)) |>
    ph_with(value = right_content, location = ph_location_label(ph_label = right_cont_ph_label))
}

#' Add a Full Content Slide (Title + Single Large Content)
#' @param doc officer::rdocx object
#' @param title Slide title
#' @param content Object for the main placeholder
#' @param layout Name of the layout to use (default: "Title and Content")
#' @param master Name of the master to use (default: "Office Theme")
add_full_content_slide <- function(doc, title, content, layout = "Title and Content", master = "Office Theme") {
  doc |>
    add_slide(layout = layout, master = master) |>
    ph_with(value = title, location = ph_location_type(type = "title")) |>
    ph_with(value = content, location = ph_location_type(type = "body"))
}

#' Add a Summary Slide
#' @param doc officer::rdocx object
#' @param title Slide title
#' @param text Large text content
#' @param layout Name of the layout to use (default: "Title Only")
#' @param master Name of the master to use (default: "Office Theme")
add_summary_slide <- function(doc, title, text, layout = "Title Only", master = "Office Theme") {
  doc |>
    add_slide(layout = layout, master = master) |>
    ph_with(value = title, location = ph_location_type(type = "title")) |>
    ph_with(value = text, location = ph_location(left = 1, top = 2, width = 8, height = 2))
}

#' Get Layout Summary from a PowerPoint object
#' @param doc officer::rdocx object
#' @return A data frame with layout and master names
get_layout_summary <- function(doc) {
  layout_summary(doc)
}

#' Get a detailed Slide Inventory from a PowerPoint object
#' @param doc officer::rdocx object
#' @return A data frame with slide index, layout, master, and title
get_slide_inventory <- function(doc) {
  # Get slide and layout metadata
  slide_md <- doc$slide$get_metadata()
  
  if (is.null(slide_md) || nrow(slide_md) == 0) {
    return(NULL)
  }
  
  layout_md <- doc$slideLayouts$get_metadata()
  
  # Prepare for merging
  slide_md$layout_filename <- basename(slide_md$layout_file)
  
  # Merge to get layout names
  inventory <- slide_md |>
    left_join(layout_md |> select(layout_name = name, filename, master_name), 
              by = c("layout_filename" = "filename")) |>
    select(slide_name = name, layout_name, master_name) |>
    mutate(index = row_number())
  
  # Function to safely get title from a slide
  get_slide_title <- function(doc, index) {
    sm <- slide_summary(doc, index = index)
    title_ph <- sm |> filter(type %in% c("title", "ctrTitle"))
    if (nrow(title_ph) > 0) {
      return(title_ph$text[1])
    } else {
      return("(No Title)")
    }
  }
  
  # Collect titles
  inventory$title <- sapply(1:nrow(inventory), function(i) get_slide_title(doc, i))
  
  # Select clean columns for output
  inventory |>
    select(index, layout_name, master_name, title)
}
