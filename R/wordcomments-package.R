#' wordcomments: Extract and Analyze Comments from Word Documents
#'
#' A toolkit for extracting, analyzing, and managing comments from Microsoft
#' Word (.docx) documents. Particularly useful for academic manuscript revision
#' workflows and collaborative document review.
#'
#' @section Main Functions:
#' \describe{
#'   \item{\code{\link{extract_comments}}}{Extract all comments with full metadata}
#'   \item{\code{\link{has_comments}}}{Quick check for presence of comments}
#'   \item{\code{\link{comment_summary}}}{Generate summary statistics}
#'   \item{\code{\link{comments_by_reviewer}}}{Filter or split comments by author}
#'   \item{\code{\link{find_comment_threads}}}{Group comments into conversation threads}
#'   \item{\code{\link{generate_response_table}}}{Create reviewer response document}
#'   \item{\code{\link{export_comments}}}{Export to Excel, Word, CSV, Markdown, or HTML}
#' }
#'
#' @section Typical Workflow:
#' ```
#' # Check for comments
#' has_comments("manuscript.docx")
#'
#' # Extract and summarize
#' comments <- extract_comments("manuscript.docx")
#' comment_summary(comments)
#'
#' # Generate response table for journal revision
#' generate_response_table(comments, "response_to_reviewers.xlsx")
#' ```
#'
#' @docType package
#' @name wordcomments-package
#' @aliases wordcomments
#'
#' @importFrom dplyr %>% filter select mutate arrange group_by summarise ungroup
#'   n bind_rows all_of row_number desc
#' @importFrom rlang .data
"_PACKAGE"

## Quiets R CMD check notes about .data
utils::globalVariables(c(".data"))
