library(officer)
library(dplyr)

#' Add a Title Slide
#' Uses 'Title Slide' layout
#' @param doc officer::rdocx object
#' @param title Main title string
#' @param subtitle Subtitle string
add_title_slide <- function(doc, title, subtitle = "") {
  doc |>
    add_slide(layout = "Title Slide", master = "Office Theme") |>
    ph_with(value = title, location = ph_location_type(type = "ctrTitle")) |>
    ph_with(value = subtitle, location = ph_location_type(type = "subTitle"))
}

#' Add a Section Header Slide
#' Uses 'Section Header' layout
#' @param doc officer::rdocx object
#' @param title Section title string
add_section_header <- function(doc, title) {
  doc |>
    add_slide(layout = "Section Header", master = "Office Theme") |>
    ph_with(value = title, location = ph_location_type(type = "title"))
}

#' Add a Two Content Slide (Chart Left, Text Right)
#' Uses 'Two Content' layout
#' @param doc officer::rdocx object
#' @param title Slide title
#' @param left_content Object for left placeholder (plot, table, text)
#' @param right_content Object for right placeholder (plot, table, text)
add_two_content_slide <- function(doc, title, left_content, right_content) {
  doc |>
    add_slide(layout = "Two Content", master = "Office Theme") |>
    ph_with(value = title, location = ph_location_type(type = "title")) |>
    ph_with(value = left_content, location = ph_location_label(ph_label = "Content Placeholder 2")) |>
    ph_with(value = right_content, location = ph_location_label(ph_label = "Content Placeholder 3"))
}

#' Add a Comparison Slide (Title + Left/Right Headers + Left/Right Content)
#' Uses 'Comparison' layout
#' @param doc officer::rdocx object
#' @param title Slide title
#' @param left_header Text for left header
#' @param right_header Text for right header
#' @param left_content Object for left body
#' @param right_content Object for right body
add_comparison_slide <- function(doc, title, left_header, right_header, left_content, right_content) {
  doc |>
    add_slide(layout = "Comparison", master = "Office Theme") |>
    ph_with(value = title, location = ph_location_type(type = "title")) |>
    # Headers
    ph_with(value = left_header, location = ph_location_label(ph_label = "Text Placeholder 2")) |>
    ph_with(value = right_header, location = ph_location_label(ph_label = "Text Placeholder 4")) |>
    # Bodies
    ph_with(value = left_content, location = ph_location_label(ph_label = "Content Placeholder 3")) |>
    ph_with(value = right_content, location = ph_location_label(ph_label = "Content Placeholder 5"))
}

#' Add a Full Content Slide (Title + Single Large Content)
#' Uses 'Title and Content' layout
#' @param doc officer::rdocx object
#' @param title Slide title
#' @param content Object for the main placeholder
add_full_content_slide <- function(doc, title, content) {
  doc |>
    add_slide(layout = "Title and Content", master = "Office Theme") |>
    ph_with(value = title, location = ph_location_type(type = "title")) |>
    ph_with(value = content, location = ph_location_type(type = "body"))
}

#' Add a Summary Slide
#' Uses 'Title Only' layout and adds a large text body
#' @param doc officer::rdocx object
#' @param title Slide title
#' @param text Large text content
add_summary_slide <- function(doc, title, text) {
  doc |>
    add_slide(layout = "Title Only", master = "Office Theme") |>
    ph_with(value = title, location = ph_location_type(type = "title")) |>
    ph_with(value = text, location = ph_location(left = 1, top = 2, width = 8, height = 2))
}
