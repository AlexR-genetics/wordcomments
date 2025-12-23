#' Export comments to various formats
#'
#' Unified export function supporting Excel, Word, CSV, Markdown, and HTML
#' output formats.
#'
#' @param comments A data frame from \code{\link{extract_comments}}, or a path
#'   to a .docx file.
#' @param output_path Path for the output file. Format is auto-detected from
#'   the file extension.
#' @param format Output format. One of "excel", "word", "csv", "markdown", "html".
#'   If NULL (default), auto-detected from \code{output_path} extension.
#' @param include_resolved Logical. If FALSE, only exports unresolved comments.
#'   Default TRUE.
#' @param columns Character vector of column names to include. If NULL (default),
#'   includes all columns except internal IDs.
#'
#' @return Invisibly returns the output path.
#'
#' @examples
#' \dontrun{
#' comments <- extract_comments("manuscript.docx")
#'
#' # Export to Excel
#' export_comments(comments, "comments.xlsx")
#'
#' # Export only unresolved to Markdown
#' export_comments(comments, "open_comments.md", include_resolved = FALSE)
#'
#' # Export specific columns
#' export_comments(comments, "simple.csv",
#'                 columns = c("Author", "Comment", "Resolved"))
#' }
#'
#' @export
export_comments <- function(comments,
                            output_path,
                            format = NULL,
                            include_resolved = TRUE,
                            columns = NULL) {
  
  # Handle file path input
  if (is.character(comments) && length(comments) == 1 && file.exists(comments)) {
    source_file <- comments
    comments <- extract_comments(comments, include_resolved = include_resolved)
  } else {
    source_file <- "comments data"
    if (!include_resolved) {
      comments <- dplyr::filter(comments, !.data$Resolved)
    }
  }
  
  if (nrow(comments) == 0) {
    message("No comments to export.")
    return(invisible(NULL))
  }
  
  # Auto-detect format from extension
  if (is.null(format)) {
    ext <- tolower(tools::file_ext(output_path))
    format <- switch(ext,
                     "xlsx" = "excel",
                     "xls" = "excel",
                     "docx" = "word",
                     "csv" = "csv",
                     "md" = "markdown",
                     "html" = "html",
                     "excel"  # default
    )
  }
  
  # Select columns
  export_df <- comments
  if (!is.null(columns)) {
    valid_cols <- intersect(columns, names(comments))
    if (length(valid_cols) == 0) {
      stop("No valid columns specified. Available: ", paste(names(comments), collapse = ", "))
    }
    export_df <- dplyr::select(comments, dplyr::all_of(valid_cols))
  } else {
    # Default: exclude internal IDs
    export_df <- dplyr::select(comments, -"comment_id", -"para_id", -"parent_id")
  }
  
  # Format dates for display
  if ("Date" %in% names(export_df)) {
    export_df <- dplyr::mutate(export_df, Date = format(.data$Date, "%Y-%m-%d %H:%M"))
  }
  
  # Export based on format
  switch(format,
         "excel" = .export_excel(export_df, output_path, source_file, include_resolved),
         "word" = .export_word(export_df, output_path, source_file, include_resolved),
         "csv" = .export_csv(export_df, output_path),
         "markdown" = .export_markdown(export_df, output_path, source_file),
         "html" = .export_html(export_df, output_path, source_file),
         stop("Unsupported format: ", format)
  )
  
  return(invisible(output_path))
}

#' @keywords internal
#' @noRd
.export_excel <- function(df, path, source, include_resolved) {
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "Comments")
  
  openxlsx::writeData(wb, 1, paste("Source:", source), startRow = 1)
  openxlsx::writeData(wb, 1, paste("Extracted:", Sys.time()), startRow = 2)
  openxlsx::writeData(wb, 1, paste("Filter:", ifelse(include_resolved, "All", "Unresolved only")), startRow = 3)
  
  openxlsx::writeData(wb, 1, df, startRow = 5)
  
  headerStyle <- openxlsx::createStyle(
    textDecoration = "bold", fgFill = "#4472C4",
    fontColour = "#FFFFFF", halign = "center"
  )
  openxlsx::addStyle(wb, 1, headerStyle, rows = 5, cols = 1:ncol(df))
  openxlsx::setColWidths(wb, 1, cols = 1:ncol(df), widths = "auto")
  
  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
  message("Exported to Excel: ", path)
}

#' @keywords internal
#' @noRd
.export_word <- function(df, path, source, include_resolved) {
  doc <- officer::read_docx()
  doc <- officer::body_add_par(doc, "Extracted Comments", style = "heading 1")
  doc <- officer::body_add_par(doc, paste("Source:", source), style = "Normal")
  doc <- officer::body_add_par(doc, paste("Extracted:", Sys.time()), style = "Normal")
  doc <- officer::body_add_par(doc, paste("Filter:", ifelse(include_resolved, "All", "Unresolved only")), style = "Normal")
  doc <- officer::body_add_par(doc, "", style = "Normal")
  doc <- officer::body_add_table(doc, df, style = "table_template")
  
  print(doc, target = path)
  message("Exported to Word: ", path)
}

#' @keywords internal
#' @noRd
.export_csv <- function(df, path) {
  utils::write.csv(df, path, row.names = FALSE)
  message("Exported to CSV: ", path)
}

#' @keywords internal
#' @noRd
.export_markdown <- function(df, path, source) {
  lines <- c(
    "# Extracted Comments",
    "",
    paste("**Source:**", source),
    paste("**Extracted:**", Sys.time()),
    "",
    "---",
    ""
  )
  
  # Create markdown table
  header <- paste("|", paste(names(df), collapse = " | "), "|")
  separator <- paste("|", paste(rep("---", ncol(df)), collapse = " | "), "|")
  
  rows <- apply(df, 1, function(row) {
    row <- gsub("\\|", "\\\\|", row)
    paste("|", paste(row, collapse = " | "), "|")
  })
  
  lines <- c(lines, header, separator, rows)
  
  writeLines(lines, path)
  message("Exported to Markdown: ", path)
}

#' @keywords internal
#' @noRd
.export_html <- function(df, path, source) {
  html <- c(
    "<!DOCTYPE html>",
    "<html><head>",
    "<meta charset='UTF-8'>",
    "<title>Extracted Comments</title>",
    "<style>",
    "body { font-family: Arial, sans-serif; margin: 20px; }",
    "table { border-collapse: collapse; width: 100%; }",
    "th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }",
    "th { background-color: #4472C4; color: white; }",
    "tr:nth-child(even) { background-color: #f2f2f2; }",
    ".resolved { background-color: #e0e0e0; }",
    "</style>",
    "</head><body>",
    "<h1>Extracted Comments</h1>",
    paste("<p><strong>Source:</strong>", source, "</p>"),
    paste("<p><strong>Extracted:</strong>", Sys.time(), "</p>"),
    "<table>",
    paste("<tr>", paste0("<th>", names(df), "</th>", collapse = ""), "</tr>")
  )
  
  for (i in seq_len(nrow(df))) {
    row_class <- if ("Resolved" %in% names(df) && df$Resolved[i]) " class='resolved'" else ""
    cells <- paste0("<td>", sapply(df[i, ], as.character), "</td>", collapse = "")
    html <- c(html, paste0("<tr", row_class, ">", cells, "</tr>"))
  }
  
  html <- c(html, "</table>", "</body></html>")
  
  writeLines(html, path)
  message("Exported to HTML: ", path)
}


#' Generate a response table for reviewer comments
#'
#' Creates a structured table for responding to reviewer comments, commonly
#' needed for journal manuscript revisions. The output includes columns for
#' documenting your response and actions taken.
#'
#' @param comments A data frame from \code{\link{extract_comments}}, or a path
#'   to a .docx file.
#' @param output_path Optional path to save the output (.xlsx, .docx, or .csv).
#'   If NULL, only returns the data frame.
#' @param group_by_reviewer Logical. If TRUE (default), groups comments by
#'   reviewer with section headers in the output.
#' @param include_resolved Logical. If FALSE (default), excludes resolved
#'   comments from the response table.
#'
#' @return A data frame with columns:
#' \describe{
#'   \item{#}{Comment number}
#'   \item{Reviewer}{Author of the comment}
#'   \item{Location}{Page and line reference}
#'   \item{Commented Text}{The text that was commented on}
#'   \item{Reviewer Comment}{The comment content}
#'   \item{Author Response}{Empty column for your response}
#'   \item{Action Taken}{Empty column to describe changes made}
#'   \item{Line in Revision}{Empty column for new line references}
#' }
#'
#' @examples
#' \dontrun{
#' # Generate and export response table
#' response <- generate_response_table("manuscript.docx",
#'                                     "response_to_reviewers.xlsx")
#'
#' # Without export, just get the data frame
#' response <- generate_response_table("manuscript.docx")
#' }
#'
#' @export
generate_response_table <- function(comments,
                                    output_path = NULL,
                                    group_by_reviewer = TRUE,
                                    include_resolved = FALSE) {
  
  # Handle file path input
  if (is.character(comments) && length(comments) == 1 && file.exists(comments)) {
    comments <- extract_comments(comments, include_resolved = include_resolved)
  } else if (!include_resolved) {
    comments <- dplyr::filter(comments, !.data$Resolved)
  }
  
  if (nrow(comments) == 0) {
    message("No comments to process.")
    return(data.frame())
  }
  
  # Build response table
  if (group_by_reviewer) {
    comments <- dplyr::arrange(comments, .data$Author, .data$Line, .data$Date)
  } else {
    comments <- dplyr::arrange(comments, .data$Line, .data$Date)
  }
  
  response_table <- comments %>%
    dplyr::mutate(
      `#` = dplyr::row_number(),
      Reviewer = .data$Author,
      Location = paste0("Page ", .data$Page, ", Line ", .data$Line),
      `Commented Text` = .data$Text,
      `Reviewer Comment` = .data$Comment,
      `Author Response` = "",
      `Action Taken` = "",
      `Line in Revision` = ""
    ) %>%
    dplyr::select(
      "#", "Reviewer", "Location", "Commented Text", "Reviewer Comment",
      "Author Response", "Action Taken", "Line in Revision"
    )
  
  # Export if path provided
  if (!is.null(output_path)) {
    ext <- tolower(tools::file_ext(output_path))
    
    if (ext == "xlsx") {
      .export_response_excel(response_table, output_path, group_by_reviewer)
    } else if (ext == "docx") {
      .export_response_word(response_table, output_path, group_by_reviewer)
    } else if (ext == "csv") {
      utils::write.csv(response_table, output_path, row.names = FALSE)
      message("Response table exported to: ", output_path)
    } else {
      warning("Unsupported format. Use .xlsx, .docx, or .csv")
    }
  }
  
  return(response_table)
}

#' @keywords internal
#' @noRd
.export_response_excel <- function(response_table, output_path, group_by_reviewer) {
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "Response to Reviewers")
  
  # Title
  openxlsx::writeData(wb, 1, "Response to Reviewer Comments", startRow = 1, startCol = 1)
  openxlsx::mergeCells(wb, 1, cols = 1:8, rows = 1)
  openxlsx::addStyle(wb, 1, openxlsx::createStyle(fontSize = 14, textDecoration = "bold"),
                     rows = 1, cols = 1)
  
  openxlsx::writeData(wb, 1, paste("Generated:", Sys.time()), startRow = 2, startCol = 1)
  
  # Write data
  openxlsx::writeData(wb, 1, response_table, startRow = 4)
  
  # Header style
  headerStyle <- openxlsx::createStyle(
    textDecoration = "bold",
    fgFill = "#4472C4",
    fontColour = "#FFFFFF",
    halign = "center",
    border = "TopBottomLeftRight"
  )
  openxlsx::addStyle(wb, 1, headerStyle, rows = 4, cols = 1:8)
  
  # Alternating row colors
  for (i in seq_len(nrow(response_table))) {
    row <- i + 4
    if (i %% 2 == 0) {
      openxlsx::addStyle(wb, 1, openxlsx::createStyle(fgFill = "#F2F2F2"),
                         rows = row, cols = 1:8, stack = TRUE)
    }
  }
  
  # Column widths
  openxlsx::setColWidths(wb, 1, cols = 1, widths = 5)
  openxlsx::setColWidths(wb, 1, cols = 2, widths = 15)
  openxlsx::setColWidths(wb, 1, cols = 3, widths = 15)
  openxlsx::setColWidths(wb, 1, cols = 4, widths = 30)
  openxlsx::setColWidths(wb, 1, cols = 5, widths = 40)
  openxlsx::setColWidths(wb, 1, cols = 6, widths = 40)
  openxlsx::setColWidths(wb, 1, cols = 7, widths = 20)
  openxlsx::setColWidths(wb, 1, cols = 8, widths = 15)
  
  # Text wrapping
  wrapStyle <- openxlsx::createStyle(wrapText = TRUE, valign = "top")
  openxlsx::addStyle(wb, 1, wrapStyle, rows = 5:(nrow(response_table) + 4),
                     cols = 1:8, gridExpand = TRUE, stack = TRUE)
  
  openxlsx::saveWorkbook(wb, output_path, overwrite = TRUE)
  message("Response table exported to: ", output_path)
}

#' @keywords internal
#' @noRd
.export_response_word <- function(response_table, output_path, group_by_reviewer) {
  doc <- officer::read_docx()
  doc <- officer::body_add_par(doc, "Response to Reviewer Comments", style = "heading 1")
  doc <- officer::body_add_par(doc, paste("Generated:", Sys.time()), style = "Normal")
  doc <- officer::body_add_par(doc, "", style = "Normal")
  
  if (group_by_reviewer) {
    reviewers <- unique(response_table$Reviewer)
    
    for (rev in reviewers) {
      rev_comments <- dplyr::filter(response_table, .data$Reviewer == rev)
      
      doc <- officer::body_add_par(doc, paste("Reviewer:", rev), style = "heading 2")
      doc <- officer::body_add_par(doc, "", style = "Normal")
      
      rev_table <- dplyr::select(rev_comments, -"Reviewer")
      doc <- officer::body_add_table(doc, rev_table, style = "table_template")
      doc <- officer::body_add_par(doc, "", style = "Normal")
    }
  } else {
    doc <- officer::body_add_table(doc, response_table, style = "table_template")
  }
  
  print(doc, target = output_path)
  message("Response table exported to: ", output_path)
}
