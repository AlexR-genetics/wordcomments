# Extract comments from a Word document

Extracts all comments from a Microsoft Word (.docx) document, including
metadata such as author, date, position, resolution status, and thread
information.

## Usage

``` r
extract_comments(docx_path, include_resolved = TRUE)
```

## Arguments

- docx_path:

  Path to the .docx file

- include_resolved:

  Logical. If TRUE (default), includes resolved comments. If FALSE, only
  returns unresolved comments.

## Value

A data frame with the following columns:

- Text:

  The text that was commented on

- Comment:

  The comment content

- Author:

  Comment author name

- Date:

  Date and time the comment was made

- Line:

  Paragraph number in document (approximate line reference)

- Page:

  Estimated page number

- Resolved:

  Logical indicating if comment thread is resolved

- comment_id:

  Internal comment ID

- para_id:

  Paragraph ID for thread matching

- parent_id:

  Parent paragraph ID (for replies)

## Details

The function parses the internal XML structure of Word documents to
extract comment data. Line numbers refer to paragraph positions since
Word doesn't store true line numbers. Page numbers are estimated based
on paragraph count (~25 paragraphs per page).

Comments are returned sorted by position (Line) then by date.

## Examples

``` r
if (FALSE) { # \dontrun{
# Extract all comments
comments <- extract_comments("manuscript.docx")

# Extract only unresolved comments
open_comments <- extract_comments("manuscript.docx", include_resolved = FALSE)
} # }
```
