# Internal helper functions for wordcomments package
# These are not exported to users

#' Unpack a docx file and return temp directory path
#' @param docx_path Path to .docx file
#' @return Path to temporary directory containing unpacked files
#' @keywords internal
#' @noRd
.unpack_docx <- function(docx_path) {
  if (!file.exists(docx_path)) {
    stop("File not found: ", docx_path)
  }
  
 temp_dir <- tempfile()
  dir.create(temp_dir)
  utils::unzip(docx_path, exdir = temp_dir)
  
  return(temp_dir)
}

#' Get standard Word XML namespaces
#' @return Named character vector of XML namespaces
#' @keywords internal
#' @noRd
.get_namespaces <- function() {
  c(
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    w14 = "http://schemas.microsoft.com/office/word/2010/wordml",
    w15 = "http://schemas.microsoft.com/office/word/2012/wordml",
    w16cex = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
  )
}

#' Create empty comments data frame
#' @return Empty data frame with correct column structure
#' @keywords internal
#' @noRd
.empty_comments_df <- function() {
  data.frame(
    Text = character(0),
    Comment = character(0),
    Author = character(0),
    Date = as.POSIXct(character(0)),
    Line = integer(0),
    Page = integer(0),
    Resolved = logical(0),
    comment_id = character(0),
    para_id = character(0),
    parent_id = character(0),
    stringsAsFactors = FALSE
  )
}
