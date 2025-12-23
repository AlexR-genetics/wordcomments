#' Generate summary statistics for comments
#'
#' Produces a comprehensive summary of comment data including counts by author,
#' resolution rates, date ranges, and comment length statistics.
#'
#' @param comments A data frame from \code{\link{extract_comments}}, or a path
#'   to a .docx file (which will be processed automatically).
#'
#' @return A list of class "comment_summary" containing:
#' \describe{
#'   \item{total_comments}{Total number of comments}
#'   \item{resolved}{Number of resolved comments}
#'   \item{unresolved}{Number of unresolved comments}
#'   \item{resolution_rate}{Percentage of comments resolved}
#'   \item{unique_authors}{Number of unique comment authors}
#'   \item{by_author}{Data frame with counts per author}
#'   \item{by_page}{Data frame with counts per page}
#'   \item{date_range}{List with earliest, latest dates and span in days}
#'   \item{comment_length}{Statistics on comment text length}
#'   \item{threads}{Information about comment threading}
#' }
#'
#' @examples
#' \dontrun{
#' # From file
#' summary <- comment_summary("manuscript.docx")
#' print(summary)
#'
#' # From existing comments data frame
#' comments <- extract_comments("manuscript.docx")
#' summary <- comment_summary(comments)
#' }
#'
#' @export
comment_summary <- function(comments) {
  
  # Handle file path input
  if (is.character(comments) && length(comments) == 1 && file.exists(comments)) {
    comments <- extract_comments(comments)
  }
  
  if (nrow(comments) == 0) {
    return(list(
      total_comments = 0,
      message = "No comments found"
    ))
  }
  
  # Basic counts
  total <- nrow(comments)
  resolved_count <- sum(comments$Resolved, na.rm = TRUE)
  unresolved_count <- total - resolved_count
  
  # By author
  by_author <- comments %>%
    dplyr::group_by(.data$Author) %>%
    dplyr::summarise(
      count = dplyr::n(),
      resolved = sum(.data$Resolved, na.rm = TRUE),
      unresolved = dplyr::n() - sum(.data$Resolved, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    dplyr::arrange(dplyr::desc(.data$count))
  
  # By page
  by_page <- comments %>%
    dplyr::filter(!is.na(.data$Page)) %>%
    dplyr::group_by(.data$Page) %>%
    dplyr::summarise(count = dplyr::n(), .groups = "drop") %>%
    dplyr::arrange(.data$Page)
  
  # Date range
  valid_dates <- comments$Date[!is.na(comments$Date)]
  date_range <- if (length(valid_dates) > 0) {
    list(
      earliest = min(valid_dates),
      latest = max(valid_dates),
      span_days = as.numeric(difftime(max(valid_dates), min(valid_dates), units = "days"))
    )
  } else {
    list(earliest = NA, latest = NA, span_days = NA)
  }
  
  # Comment length stats
  comment_lengths <- nchar(comments$Comment)
  length_stats <- list(
    mean_chars = round(mean(comment_lengths, na.rm = TRUE), 1),
    median_chars = stats::median(comment_lengths, na.rm = TRUE),
    max_chars = max(comment_lengths, na.rm = TRUE),
    min_chars = min(comment_lengths, na.rm = TRUE)
  )
  
  # Thread stats
  thread_count <- sum(!is.na(comments$parent_id))
  root_comments <- sum(is.na(comments$parent_id))
  
  # Build summary list
  summary_list <- list(
    total_comments = total,
    resolved = resolved_count,
    unresolved = unresolved_count,
    resolution_rate = round(resolved_count / total * 100, 1),
    unique_authors = length(unique(comments$Author)),
    by_author = by_author,
    by_page = by_page,
    date_range = date_range,
    comment_length = length_stats,
    threads = list(
      root_comments = root_comments,
      replies = thread_count,
      has_threading = thread_count > 0
    )
  )
  
  class(summary_list) <- c("comment_summary", "list")
  return(summary_list)
}

#' Print method for comment_summary
#'
#' @param x A comment_summary object
#' @param ... Additional arguments (unused)
#'
#' @return Invisibly returns the input object
#' @export
print.comment_summary <- function(x, ...) {
  cat("=== COMMENT SUMMARY ===\n\n")
  
  cat("OVERVIEW\n")
  cat(sprintf("  Total comments:    %d\n", x$total_comments))
  cat(sprintf("  Resolved:          %d (%.1f%%)\n", x$resolved, x$resolution_rate))
  cat(sprintf("  Unresolved:        %d\n", x$unresolved))
  cat(sprintf("  Unique authors:    %d\n", x$unique_authors))
  
  if (x$threads$has_threading) {
    cat(sprintf("  Comment threads:   %d root + %d replies\n",
                x$threads$root_comments, x$threads$replies))
  }
  
  cat("\nBY AUTHOR\n")
  for (i in seq_len(nrow(x$by_author))) {
    row <- x$by_author[i, ]
    cat(sprintf("  %-20s %3d comments (%d resolved, %d open)\n",
                row$Author, row$count, row$resolved, row$unresolved))
  }
  
  if (nrow(x$by_page) > 0) {
    cat("\nBY PAGE\n")
    pages_str <- paste(sprintf("p%d:%d", x$by_page$Page, x$by_page$count), collapse = ", ")
    cat(sprintf("  %s\n", pages_str))
  }
  
  if (!is.na(x$date_range$earliest)) {
    cat("\nDATE RANGE\n")
    cat(sprintf("  From: %s\n", format(x$date_range$earliest, "%Y-%m-%d %H:%M")))
    cat(sprintf("  To:   %s\n", format(x$date_range$latest, "%Y-%m-%d %H:%M")))
    cat(sprintf("  Span: %.1f days\n", x$date_range$span_days))
  }
  
  cat("\nCOMMENT LENGTH\n")
  cat(sprintf("  Mean: %.0f chars | Median: %d | Range: %d-%d\n",
              x$comment_length$mean_chars, x$comment_length$median_chars,
              x$comment_length$min_chars, x$comment_length$max_chars))
  
  invisible(x)
}


#' Filter or split comments by reviewer/author
#'
#' Filters comments to a specific reviewer or splits all comments into
#' separate data frames by author.
#'
#' @param comments A data frame from \code{\link{extract_comments}}, or a path
#'   to a .docx file.
#' @param reviewer Optional character string. If provided, filters to comments
#'   from this reviewer (case-insensitive partial matching).
#' @param split Logical. If TRUE and no reviewer specified, returns a named
#'   list of data frames split by author. Default FALSE.
#'
#' @return If \code{reviewer} is specified or \code{split = FALSE}: a data frame.
#'   If \code{split = TRUE}: a named list of data frames, one per author.
#'
#' @examples
#' \dontrun{
#' comments <- extract_comments("manuscript.docx")
#'
#' # Filter to one reviewer
#' smith_comments <- comments_by_reviewer(comments, reviewer = "Smith")
#'
#' # Split by all reviewers
#' by_reviewer <- comments_by_reviewer(comments, split = TRUE)
#' names(by_reviewer)  # Shows all reviewer names
#' }
#'
#' @export
comments_by_reviewer <- function(comments, reviewer = NULL, split = FALSE) {
  
  # Handle file path input
  if (is.character(comments) && length(comments) == 1 && file.exists(comments)) {
    comments <- extract_comments(comments)
  }
  
  if (nrow(comments) == 0) {
    if (split) return(list())
    return(comments)
  }
  
  # Filter for specific reviewer
  if (!is.null(reviewer)) {
    matched <- comments %>%
      dplyr::filter(grepl(reviewer, .data$Author, ignore.case = TRUE))
    
    if (nrow(matched) == 0) {
      message("No comments found for reviewer matching: ", reviewer)
      message("Available authors: ", paste(unique(comments$Author), collapse = ", "))
    }
    
    return(matched)
  }
  
  # Split by all reviewers
  if (split) {
    authors <- unique(comments$Author)
    result <- lapply(authors, function(auth) {
      dplyr::filter(comments, .data$Author == auth)
    })
    names(result) <- authors
    return(result)
  }
  
  # Default: return all, sorted by author then line
  return(dplyr::arrange(comments, .data$Author, .data$Line, .data$Date))
}


#' Group comments into conversation threads
#'
#' Identifies and groups comments that are part of the same conversation thread,
#' including replies to comments.
#'
#' @param comments A data frame from \code{\link{extract_comments}}, or a path
#'   to a .docx file.
#' @param flatten Logical. If TRUE (default), returns a flat data frame with
#'   thread_id column added. If FALSE, returns a nested list structure.
#'
#' @return If \code{flatten = TRUE}: a data frame with additional columns:
#' \describe{
#'   \item{thread_id}{Integer identifying the conversation thread}
#'   \item{thread_size}{Number of comments in this thread}
#'   \item{reply_depth}{Nesting depth (0 for root comments)}
#' }
#'
#' If \code{flatten = FALSE}: a list of thread objects, each containing:
#' \describe{
#'   \item{thread_id}{Thread identifier}
#'   \item{root_comment}{Data frame with the original comment}
#'   \item{replies}{Data frame with reply comments}
#'   \item{size}{Total comments in thread}
#'   \item{authors}{Character vector of participating authors}
#'   \item{resolved}{Logical indicating if thread is resolved}
#' }
#'
#' @examples
#' \dontrun{
#' comments <- extract_comments("manuscript.docx")
#'
#' # Flat format with thread IDs
#' threaded <- find_comment_threads(comments)
#'
#' # Nested list format
#' threads <- find_comment_threads(comments, flatten = FALSE)
#' }
#'
#' @export
find_comment_threads <- function(comments, flatten = TRUE) {
  
  # Handle file path input
  if (is.character(comments) && length(comments) == 1 && file.exists(comments)) {
    comments <- extract_comments(comments)
  }
  
  if (nrow(comments) == 0) {
    if (flatten) return(dplyr::mutate(comments, thread_id = integer(0)))
    return(list())
  }
  
  # Build thread structure
  comments <- comments %>%
    dplyr::mutate(
      is_root = is.na(.data$parent_id),
      thread_id = NA_integer_
    )
  
  # Assign thread IDs
  thread_counter <- 0
  para_to_thread <- list()
  
  # First pass: assign thread IDs to root comments
  for (i in seq_len(nrow(comments))) {
    if (comments$is_root[i]) {
      thread_counter <- thread_counter + 1
      comments$thread_id[i] <- thread_counter
      
      pid <- comments$para_id[i]
      if (!is.na(pid)) {
        para_to_thread[[pid]] <- thread_counter
      }
    }
  }
  
  # Second pass: assign thread IDs to replies
  max_iterations <- 10
  for (iter in seq_len(max_iterations)) {
    changes <- 0
    
    for (i in seq_len(nrow(comments))) {
      if (is.na(comments$thread_id[i]) && !is.na(comments$parent_id[i])) {
        parent_pid <- comments$parent_id[i]
        
        if (!is.null(para_to_thread[[parent_pid]])) {
          comments$thread_id[i] <- para_to_thread[[parent_pid]]
          
          pid <- comments$para_id[i]
          if (!is.na(pid)) {
            para_to_thread[[pid]] <- para_to_thread[[parent_pid]]
          }
          changes <- changes + 1
        }
      }
    }
    
    if (changes == 0) break
  }
  
  # Handle orphaned replies
  for (i in seq_len(nrow(comments))) {
    if (is.na(comments$thread_id[i])) {
      thread_counter <- thread_counter + 1
      comments$thread_id[i] <- thread_counter
    }
  }
  
  # Calculate thread depth
  comments <- comments %>%
    dplyr::group_by(.data$thread_id) %>%
    dplyr::mutate(
      thread_size = dplyr::n(),
      reply_depth = ifelse(.data$is_root, 0, dplyr::row_number() - 1)
    ) %>%
    dplyr::ungroup() %>%
    dplyr::arrange(.data$thread_id, .data$Date)
  
  if (flatten) {
    return(dplyr::select(comments, -"is_root"))
  }
  
  # Nested list structure
  threads <- split(comments, comments$thread_id)
  threads <- lapply(threads, function(thread_df) {
    list(
      thread_id = thread_df$thread_id[1],
      root_comment = dplyr::select(dplyr::filter(thread_df, .data$is_root),
                                   -"is_root", -"thread_id"),
      replies = dplyr::select(dplyr::filter(thread_df, !.data$is_root),
                              -"is_root", -"thread_id"),
      size = nrow(thread_df),
      authors = unique(thread_df$Author),
      resolved = all(thread_df$Resolved)
    )
  })
  
  return(threads)
}
