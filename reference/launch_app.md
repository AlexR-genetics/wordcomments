# Launch the wordcomments Shiny Application

Opens an interactive GUI for extracting and analyzing Word document
comments. The app provides a user-friendly interface for all
wordcomments functions.

## Usage

``` r
launch_app(launch.browser = TRUE, port = NULL, host = "127.0.0.1")

run_wordcomments(launch.browser = TRUE, port = NULL, host = "127.0.0.1")
```

## Arguments

- launch.browser:

  Logical. If TRUE (default), opens the app in your default web browser.
  If FALSE, returns the app URL for manual access.

- port:

  Integer. The port to run the app on. Default is a random available
  port.

- host:

  Character. The host to run on. Default "127.0.0.1" (localhost).

## Value

Invisibly returns the Shiny app object.

## Details

The GUI provides:

- File browser for selecting one or multiple .docx files

- Options for including resolved comments and tracking source files

- Tabs for viewing comments, summary statistics, threads, and
  by-reviewer breakdowns

- Interactive response table generation

- Export to Excel, Word, CSV, Markdown, or HTML

## Examples

``` r
if (FALSE) { # \dontrun{
# Launch the app
launch_app()

# Launch without opening browser
launch_app(launch.browser = FALSE)
} # }
```
