#' Launch the wordcomments Shiny Application
#'
#' Opens an interactive GUI for extracting and analyzing Word document comments.
#' The app provides a user-friendly interface for all wordcomments functions.
#'
#' @param launch.browser Logical. If TRUE (default), opens the app in your
#'   default web browser. If FALSE, returns the app URL for manual access.
#' @param port Integer. The port to run the app on. Default is a random available port.
#' @param host Character. The host to run on. Default "127.0.0.1" (localhost).
#'
#' @return Invisibly returns the Shiny app object.
#'
#' @details
#' The GUI provides:
#' \itemize{
#'   \item File browser for selecting one or multiple .docx files
#'   \item Options for including resolved comments and tracking source files
#'   \item Tabs for viewing comments, summary statistics, threads, and by-reviewer breakdowns
#'   \item Interactive response table generation
#'   \item Export to Excel, Word, CSV, Markdown, or HTML
#' }
#'
#' @examples
#' \dontrun{
#' # Launch the app
#' launch_app()
#'
#' # Launch without opening browser
#' launch_app(launch.browser = FALSE)
#' }
#'
#' @export
launch_app <- function(launch.browser = TRUE, port = NULL, host = "127.0.0.1") {
  
  # Check for required packages
 required_pkgs <- c("shiny", "bslib", "DT")
  missing_pkgs <- required_pkgs[!sapply(required_pkgs, requireNamespace, quietly = TRUE)]
  
  if (length(missing_pkgs) > 0) {
    stop(
      "The following packages are required for the GUI but not installed:\n",
      paste(" -", missing_pkgs, collapse = "\n"),
      "\n\nInstall them with:\n",
      sprintf("install.packages(c(%s))", 
              paste(sprintf('"%s"', missing_pkgs), collapse = ", ")),
      call. = FALSE
    )
  }
  
  # Find the app directory
  app_dir <- system.file("shiny-app", package = "wordcomments")
  
  if (app_dir == "") {
    stop("Could not find Shiny app. Try reinstalling wordcomments.", call. = FALSE)
  }
  
  # Build run arguments
  run_args <- list(
    appDir = app_dir,
    launch.browser = launch.browser,
    host = host
  )
  
  if (!is.null(port)) {
    run_args$port <- port
  }
  
  # Launch the app
  message("Launching wordcomments GUI...")
  message("Close the browser window or press Ctrl+C to stop the app.\n")
  
  do.call(shiny::runApp, run_args)
}


#' @rdname launch_app
#' @export
run_wordcomments <- launch_app
