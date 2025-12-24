# Export comments to various formats

Unified export function supporting Excel, Word, CSV, Markdown, and HTML
output formats.

## Usage

``` r
export_comments(
  comments,
  output_path,
  format = NULL,
  include_resolved = TRUE,
  columns = NULL
)
```

## Arguments

- comments:

  A data frame from
  [`extract_comments`](https://alexr-genetics.github.io/wordcomments/reference/extract_comments.md),
  or a path to a .docx file.

- output_path:

  Path for the output file. Format is auto-detected from the file
  extension.

- format:

  Output format. One of "excel", "word", "csv", "markdown", "html". If
  NULL (default), auto-detected from `output_path` extension.

- include_resolved:

  Logical. If FALSE, only exports unresolved comments. Default TRUE.

- columns:

  Character vector of column names to include. If NULL (default),
  includes all columns except internal IDs.

## Value

Invisibly returns the output path.

## Examples

``` r
if (FALSE) { # \dontrun{
comments <- extract_comments("manuscript.docx")

# Export to Excel
export_comments(comments, "comments.xlsx")

# Export only unresolved to Markdown
export_comments(comments, "open_comments.md", include_resolved = FALSE)

# Export specific columns
export_comments(comments, "simple.csv",
                columns = c("Author", "Comment", "Resolved"))
} # }
```
