# Extract and merge comments from multiple document versions

Extracts comments from multiple versions of a document (or related
documents) and merges them into a single data frame. Comment IDs are
renumbered to avoid conflicts. The output format is identical to
[`extract_comments`](https://alexr-genetics.github.io/wordcomments/reference/extract_comments.md),
making it compatible with all other wordcomments functions.

## Usage

``` r
extract_comments_multiple(
  docx_paths,
  include_resolved = TRUE,
  add_source = FALSE
)
```

## Arguments

- docx_paths:

  Character vector of paths to .docx files

- include_resolved:

  Logical. If TRUE (default), includes resolved comments. If FALSE, only
  returns unresolved comments.

- add_source:

  Logical. If TRUE, adds a Source column indicating which document each
  comment came from. Default FALSE for compatibility with other
  functions.

## Value

A data frame with the same structure as
[`extract_comments`](https://alexr-genetics.github.io/wordcomments/reference/extract_comments.md):

- Text:

  The text that was commented on

- Comment:

  The comment content

- Author:

  Comment author name

- Date:

  Date and time the comment was made

- Line:

  Paragraph number in document

- Page:

  Estimated page number

- Resolved:

  Logical indicating if comment thread is resolved

- comment_id:

  Internal comment ID (renumbered to avoid conflicts)

- para_id:

  Paragraph ID for thread matching (renumbered)

- parent_id:

  Parent paragraph ID for replies (renumbered)

- Source:

  (Optional) Source document filename, if add_source = TRUE

## Details

This function is useful when you have multiple reviewers working on
separate copies of a document and you want to consolidate all their
comments.

Comments are sorted by Line (position) then by Date, regardless of which
document they came from. This gives a unified view of all feedback
organized by where it appears in the document.

Since comment IDs are renumbered, thread relationships (replies) are
preserved within each document but not across documents.

## Examples

``` r
if (FALSE) { # \dontrun{
# Extract comments from multiple reviewer versions
all_comments <- extract_comments_multiple(
  c("manuscript_reviewer1.docx",
    "manuscript_reviewer2.docx",
    "manuscript_editor.docx")
)

# Use with other wordcomments functions
comment_summary(all_comments)
comments_by_reviewer(all_comments, split = TRUE)
generate_response_table(all_comments, "consolidated_response.xlsx")

# Include source document info
all_comments <- extract_comments_multiple(
  c("reviewer1.docx", "reviewer2.docx"),
  add_source = TRUE
)
table(all_comments$Source)  # See comment counts per document
} # }
```
