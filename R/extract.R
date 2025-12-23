#' Extract comments from a Word document
#'
#' Extracts all comments from a Microsoft Word (.docx) document, including
#' metadata such as author, date, position, resolution status, and thread
#' information.
#'
#' @param docx_path Path to the .docx file
#' @param include_resolved Logical. If TRUE (default), includes resolved comments.
#'   If FALSE, only returns unresolved comments.
#'
#' @return A data frame with the following columns:
#' \describe{
#'   \item{Text}{The text that was commented on}
#'   \item{Comment}{The comment content}
#'   \item{Author}{Comment author name}
#'   \item{Date}{Date and time the comment was made}
#'   \item{Line}{Paragraph number in document (approximate line reference)}
#'   \item{Page}{Estimated page number}
#'   \item{Resolved}{Logical indicating if comment thread is resolved}
#'   \item{comment_id}{Internal comment ID}
#'   \item{para_id}{Paragraph ID for thread matching}
#'   \item{parent_id}{Parent paragraph ID (for replies)}
#' }
#'
#' @details
#' The function parses the internal XML structure of Word documents to extract
#' comment data. Line numbers refer to paragraph positions since Word doesn't
#' store true line numbers. Page numbers are estimated based on paragraph count
#' (~25 paragraphs per page).
#'
#' Comments are returned sorted by position (Line) then by date.
#'
#' @examples
#' \dontrun{
#' # Extract all comments
#' comments <- extract_comments("manuscript.docx")
#'
#' # Extract only unresolved comments
#' open_comments <- extract_comments("manuscript.docx", include_resolved = FALSE)
#' }
#'
#' @export
extract_comments <- function(docx_path, include_resolved = TRUE) {
  
  temp_dir <- .unpack_docx(docx_path)
  on.exit(unlink(temp_dir, recursive = TRUE))
  
  ns <- .get_namespaces()
  
 # --- Parse document.xml for positions ---
  document_path <- file.path(temp_dir, "word", "document.xml")
  document_xml <- xml2::read_xml(document_path)
  
  all_paragraphs <- xml2::xml_find_all(document_xml, "//w:p", ns)
  paragraphs_per_page <- 25
  
  # --- Parse comments.xml ---
  comments_path <- file.path(temp_dir, "word", "comments.xml")
  
  if (!file.exists(comments_path)) {
    message("No comments found in document.")
    return(.empty_comments_df())
  }
  
  comments_xml <- xml2::read_xml(comments_path)
  comment_nodes <- xml2::xml_find_all(comments_xml, "//w:comment", ns)
  
  if (length(comment_nodes) == 0) {
    message("No comments found in document.")
    return(.empty_comments_df())
  }
  
  # Extract comment metadata including parent references for threads
  comments_list <- lapply(comment_nodes, function(node) {
    text_nodes <- xml2::xml_find_all(node, ".//w:t", ns)
    comment_text <- paste(xml2::xml_text(text_nodes), collapse = " ")
    
    # Get paraId from first paragraph
    first_para <- xml2::xml_find_first(node, ".//w:p", ns)
    para_id <- NA
    if (!is.na(first_para)) {
      all_attrs <- xml2::xml_attrs(first_para)
      para_keys <- grep("paraId", names(all_attrs), value = TRUE, ignore.case = TRUE)
      if (length(para_keys) > 0) para_id <- all_attrs[[para_keys[1]]]
    }
    
    list(
      comment_id = xml2::xml_attr(node, "id"),
      author = xml2::xml_attr(node, "author"),
      date = xml2::xml_attr(node, "date"),
      comment_text = comment_text,
      para_id = para_id
    )
  })
  
  comments_df <- dplyr::bind_rows(lapply(comments_list, as.data.frame, stringsAsFactors = FALSE))
  
  # --- Get parent IDs for threading from commentsExtended.xml ---
  parent_map <- list()
  comments_ext_path <- file.path(temp_dir, "word", "commentsExtended.xml")
  
  if (file.exists(comments_ext_path)) {
    comments_ext_xml <- xml2::read_xml(comments_ext_path)
    comment_ex_nodes <- xml2::xml_find_all(comments_ext_xml, "//*[local-name()='commentEx']")
    
    for (node in comment_ex_nodes) {
      all_attrs <- xml2::xml_attrs(node)
      
      para_keys <- grep("paraId", names(all_attrs), value = TRUE, ignore.case = TRUE)
      para_id <- if (length(para_keys) > 0) all_attrs[[para_keys[1]]] else NA
      
      parent_keys <- grep("paraIdParent", names(all_attrs), value = TRUE, ignore.case = TRUE)
      parent_id <- if (length(parent_keys) > 0) all_attrs[[parent_keys[1]]] else NA
      
      if (!is.na(para_id)) {
        parent_map[[para_id]] <- parent_id
      }
    }
  }
  
  comments_df$parent_id <- sapply(comments_df$para_id, function(pid) {
    if (is.na(pid) || is.null(parent_map[[pid]])) NA else parent_map[[pid]]
  })
  
  # --- Find comment positions ---
  comment_range_starts <- xml2::xml_find_all(document_xml, "//w:commentRangeStart", ns)
  comment_positions <- list()
  
  for (start_node in comment_range_starts) {
    comment_id <- xml2::xml_attr(start_node, "id")
    
    para_index <- NA
    for (i in seq_along(all_paragraphs)) {
      descendants <- xml2::xml_find_all(all_paragraphs[[i]], ".//w:commentRangeStart", ns)
      for (d in descendants) {
        if (xml2::xml_attr(d, "id") == comment_id) {
          para_index <- i
          break
        }
      }
      if (!is.na(para_index)) break
    }
    
    if (!is.na(para_index)) {
      text_nodes <- xml2::xml_find_all(all_paragraphs[[para_index]], ".//w:t", ns)
      para_text <- paste(xml2::xml_text(text_nodes), collapse = "")
      
      comment_positions[[comment_id]] <- list(
        line = para_index,
        page = ceiling(para_index / paragraphs_per_page),
        text = para_text
      )
    }
  }
  
  # Fallback with commentReference
  comment_refs <- xml2::xml_find_all(document_xml, "//w:commentReference", ns)
  for (ref in comment_refs) {
    comment_id <- xml2::xml_attr(ref, "id")
    if (is.null(comment_positions[[comment_id]])) {
      for (i in seq_along(all_paragraphs)) {
        descendants <- xml2::xml_find_all(all_paragraphs[[i]], ".//w:commentReference", ns)
        for (d in descendants) {
          if (xml2::xml_attr(d, "id") == comment_id) {
            text_nodes <- xml2::xml_find_all(all_paragraphs[[i]], ".//w:t", ns)
            comment_positions[[comment_id]] <- list(
              line = i,
              page = ceiling(i / paragraphs_per_page),
              text = paste(xml2::xml_text(text_nodes), collapse = "")
            )
            break
          }
        }
      }
    }
  }
  
  # --- Determine resolved status ---
  resolved_para_ids <- character(0)
  
  if (file.exists(comments_ext_path)) {
    comments_ext_xml <- xml2::read_xml(comments_ext_path)
    comment_ex_nodes <- xml2::xml_find_all(comments_ext_xml, "//*[local-name()='commentEx']")
    
    for (node in comment_ex_nodes) {
      all_attrs <- xml2::xml_attrs(node)
      done_keys <- grep("done", names(all_attrs), value = TRUE, ignore.case = TRUE)
      
      if (length(done_keys) > 0 && all_attrs[[done_keys[1]]] == "1") {
        para_keys <- grep("paraId", names(all_attrs), value = TRUE, ignore.case = TRUE)
        if (length(para_keys) > 0) {
          resolved_para_ids <- c(resolved_para_ids, all_attrs[[para_keys[1]]])
        }
      }
    }
  }
  
  # Check commentsExtensible.xml
  comments_extensible_path <- file.path(temp_dir, "word", "commentsExtensible.xml")
  if (file.exists(comments_extensible_path)) {
    comments_extensible_xml <- xml2::read_xml(comments_extensible_path)
    all_nodes <- xml2::xml_find_all(comments_extensible_xml, "//*")
    
    for (node in all_nodes) {
      all_attrs <- xml2::xml_attrs(node)
      done_keys <- grep("done", names(all_attrs), value = TRUE, ignore.case = TRUE)
      
      if (length(done_keys) > 0 && all_attrs[[done_keys[1]]] == "1") {
        para_keys <- grep("paraId", names(all_attrs), value = TRUE, ignore.case = TRUE)
        if (length(para_keys) > 0) {
          resolved_para_ids <- c(resolved_para_ids, all_attrs[[para_keys[1]]])
        }
      }
    }
  }
  
  # --- Build final results ---
  results <- data.frame(
    Text = character(nrow(comments_df)),
    Comment = comments_df$comment_text,
    Author = comments_df$author,
    Date = as.POSIXct(comments_df$date, format = "%Y-%m-%dT%H:%M:%S", tz = "UTC"),
    Line = integer(nrow(comments_df)),
    Page = integer(nrow(comments_df)),
    Resolved = logical(nrow(comments_df)),
    comment_id = comments_df$comment_id,
    para_id = comments_df$para_id,
    parent_id = comments_df$parent_id,
    stringsAsFactors = FALSE
  )
  
  for (i in seq_len(nrow(results))) {
    cid <- comments_df$comment_id[i]
    pid <- comments_df$para_id[i]
    
    if (!is.null(comment_positions[[cid]])) {
      results$Text[i] <- comment_positions[[cid]]$text
      results$Line[i] <- comment_positions[[cid]]$line
      results$Page[i] <- comment_positions[[cid]]$page
    } else {
      results$Text[i] <- ""
      results$Line[i] <- NA
      results$Page[i] <- NA
    }
    
    results$Resolved[i] <- !is.na(pid) && pid %in% resolved_para_ids
  }
  
  # Sort by position then date
  results <- dplyr::arrange(results, .data$Line, .data$Date)
  
  # Filter if requested
  if (!include_resolved) {
    results <- dplyr::filter(results, !.data$Resolved)
  }
  
  return(results)
}


#' Check if a Word document contains comments
#'
#' Performs a fast check for the presence of comments without full parsing.
#' Useful for filtering files before batch processing.
#'
#' @param docx_path Path to the .docx file
#'
#' @return Logical. TRUE if the document contains at least one comment,
#'   FALSE otherwise.
#'
#' @examples
#' \dontrun{
#' # Quick check before processing
#' if (has_comments("document.docx")) {
#'   comments <- extract_comments("document.docx")
#' }
#'
#' # Filter a list of files
#' files <- list.files(pattern = "\\.docx$")
#' files_with_comments <- files[sapply(files, has_comments)]
#' }
#'
#' @export
has_comments <- function(docx_path) {
  if (!file.exists(docx_path)) {
    stop("File not found: ", docx_path)
  }
  
  # List files in the docx without fully extracting
  files_in_docx <- utils::unzip(docx_path, list = TRUE)$Name
  
  if (!"word/comments.xml" %in% files_in_docx) {
    return(FALSE)
  }
  
  # Extract just comments.xml to check content
  temp_dir <- tempfile()
  dir.create(temp_dir)
  on.exit(unlink(temp_dir, recursive = TRUE))
  
  utils::unzip(docx_path, files = "word/comments.xml", exdir = temp_dir)
  
  comments_path <- file.path(temp_dir, "word", "comments.xml")
  comments_xml <- xml2::read_xml(comments_path)
  
  ns <- .get_namespaces()
  comment_nodes <- xml2::xml_find_all(comments_xml, "//w:comment", ns)
  
  return(length(comment_nodes) > 0)
}
