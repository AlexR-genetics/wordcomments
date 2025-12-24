#' Extract and merge comments from multiple document versions
#'
#' Extracts comments from multiple versions of a document (or related documents)
#' and merges them into a single data frame. Comment IDs are renumbered to avoid
#' conflicts. The output format is identical to \code{\link{extract_comments}},
#' making it compatible with all other wordcomments functions.
#'
#' @param docx_paths Character vector of paths to .docx files
#' @param include_resolved Logical. If TRUE (default), includes resolved comments.
#'   If FALSE, only returns unresolved comments.
#' @param add_source Logical. If TRUE, adds a Source column indicating which
#'   document each comment came from. Default FALSE for compatibility with
#'   other functions.
#'
#' @return A data frame with the same structure as \code{\link{extract_comments}}:
#' \describe{
#'   \item{Text}{The text that was commented on}
#'   \item{Comment}{The comment content}
#'   \item{Author}{Comment author name}
#'   \item{Date}{Date and time the comment was made}
#'   \item{Line}{Paragraph number in document}
#'   \item{Page}{Estimated page number}
#'   \item{Resolved}{Logical indicating if comment thread is resolved}
#'   \item{comment_id}{Internal comment ID (renumbered to avoid conflicts)}
#'   \item{para_id}{Paragraph ID for thread matching (renumbered)}
#'   \item{parent_id}{Parent paragraph ID for replies (renumbered)}
#'   \item{Source}{(Optional) Source document filename, if add_source = TRUE}
#' }
#'
#' @details
#' This function is useful when you have multiple reviewers working on separate
#' copies of a document and you want to consolidate all their comments.
#'
#' Comments are sorted by Line (position) then by Date, regardless of which
#' document they came from. This gives a unified view of all feedback organized
#' by where it appears in the document.
#'
#' Since comment IDs are renumbered, thread relationships (replies) are preserved
#' within each document but not across documents.
#'
#' @examples
#' \dontrun{
#' # Extract comments from multiple reviewer versions
#' all_comments <- extract_comments_multiple(
#'   c("manuscript_reviewer1.docx",
#'     "manuscript_reviewer2.docx",
#'     "manuscript_editor.docx")
#' )
#'
#' # Use with other wordcomments functions
#' comment_summary(all_comments)
#' comments_by_reviewer(all_comments, split = TRUE)
#' generate_response_table(all_comments, "consolidated_response.xlsx")
#'
#' # Include source document info
#' all_comments <- extract_comments_multiple(
#'   c("reviewer1.docx", "reviewer2.docx"),
#'   add_source = TRUE
#' )
#' table(all_comments$Source)  # See comment counts per document
#' }
#'
#' @export
extract_comments_multiple <- function(docx_paths, 
                                       include_resolved = TRUE,
                                       add_source = FALSE) {
  

 # Validate inputs
  if (!is.character(docx_paths) || length(docx_paths) == 0) {
    stop("docx_paths must be a character vector of file paths")
  }
  
  # Check all files exist
  missing_files <- docx_paths[!file.exists(docx_paths)]
  if (length(missing_files) > 0) {
    stop("Files not found: ", paste(missing_files, collapse = ", "))
  }
  
  # Initialize list to store comments from each document
  all_comments <- list()
  
  # Track ID offsets to avoid conflicts
  comment_id_offset <- 0
  para_id_offset <- 0
  
  for (i in seq_along(docx_paths)) {
    docx_path <- docx_paths[i]
    doc_name <- basename(docx_path)
    
    message(sprintf("Processing [%d/%d]: %s", i, length(docx_paths), doc_name))
    
    # Extract comments from this document
    comments <- extract_comments(docx_path, include_resolved = TRUE)
    
    if (nrow(comments) == 0) {
      message(sprintf("  No comments found in %s", doc_name))
      next
    }
    
    message(sprintf("  Found %d comments", nrow(comments)))
    
    # Renumber comment_id to avoid conflicts
    # Original IDs are typically numeric strings like "0", "1", "2"
    comments$original_comment_id <- comments$comment_id
    comments$comment_id <- paste0("doc", i, "_", comments$comment_id)
    
    # Renumber para_id (these are hex strings like "2B5E7F8A")
    # Create a mapping for this document's para_ids
    unique_para_ids <- unique(c(comments$para_id, comments$parent_id))
    unique_para_ids <- unique_para_ids[!is.na(unique_para_ids)]
    
    if (length(unique_para_ids) > 0) {
      para_id_map <- setNames(
        paste0("doc", i, "_", unique_para_ids),
        unique_para_ids
      )
      
      # Apply mapping to para_id
      comments$para_id <- ifelse(
        is.na(comments$para_id),
        NA,
        para_id_map[comments$para_id]
      )
      
      # Apply mapping to parent_id
      comments$parent_id <- ifelse(
        is.na(comments$parent_id),
        NA,
        para_id_map[comments$parent_id]
      )
    }
    
    # Add source column
    comments$Source <- doc_name
    
    # Remove temporary column
    comments$original_comment_id <- NULL
    
    all_comments[[i]] <- comments
  }
  
  # Combine all comments
  if (length(all_comments) == 0) {
    message("No comments found in any document.")
    result <- .empty_comments_df()
    if (add_source) {
      result$Source <- character(0)
    }
    return(result)
  }
  
  combined <- dplyr::bind_rows(all_comments)
  
  # Sort by Line then Date (unified view across all documents)
  combined <- dplyr::arrange(combined, .data$Line, .data$Date)
  
  # Filter resolved if requested
  if (!include_resolved) {
    combined <- dplyr::filter(combined, !.data$Resolved)
  }
  
  # Remove Source column if not requested (for compatibility)
  if (!add_source) {
    combined$Source <- NULL
  }
  
  message(sprintf("\nTotal: %d comments from %d documents", 
                  nrow(combined), length(docx_paths)))
  
  return(combined)
}


#' Quick check for comments in multiple documents
#'
#' Checks multiple documents for the presence of comments and returns a summary.
#'
#' @param docx_paths Character vector of paths to .docx files
#'
#' @return A data frame with columns:
#' \describe{
#'   \item{file}{Document filename}
#'   \item{path}{Full path to document}
#'   \item{has_comments}{Logical indicating if document has comments}
#' }
#'
#' @examples
#' \dontrun{
#' # Check which documents have comments
#' status <- has_comments_multiple(
#'   c("doc1.docx", "doc2.docx", "doc3.docx")
#' )
#' print(status)
#'
#' # Filter to only documents with comments
#' docs_with_comments <- status$path[status$has_comments]
#' }
#'
#' @export
has_comments_multiple <- function(docx_paths) {
  
  if (!is.character(docx_paths) || length(docx_paths) == 0) {
    stop("docx_paths must be a character vector of file paths")
  }
  
  results <- data.frame(
    file = basename(docx_paths),
    path = docx_paths,
    has_comments = NA,
    stringsAsFactors = FALSE
  )
  
  for (i in seq_along(docx_paths)) {
    if (file.exists(docx_paths[i])) {
      results$has_comments[i] <- has_comments(docx_paths[i])
    } else {
      results$has_comments[i] <- NA
      warning("File not found: ", docx_paths[i])
    }
  }
  
  return(results)
}


# =============================================================================
# TESTING / USAGE EXAMPLES
# =============================================================================

# # Example usage:
# 
# # 1. Extract from multiple documents
# all_comments <- extract_comments_multiple(
#   c("manuscript_v1_reviewer1.docx",
#     "manuscript_v1_reviewer2.docx", 
#     "manuscript_v1_editor.docx")
# )
# 
# # 2. Works with all existing functions
# comment_summary(all_comments)
# 
# by_reviewer <- comments_by_reviewer(all_comments, split = TRUE)
# 
# threads <- find_comment_threads(all_comments)
# 
# generate_response_table(all_comments, "consolidated_response.xlsx")
# 
# export_comments(all_comments, "all_comments.xlsx")
#
# # 3. With source tracking
# all_comments <- extract_comments_multiple(
#   c("reviewer1.docx", "reviewer2.docx"),
#   add_source = TRUE
# )
# table(all_comments$Source)
