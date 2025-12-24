# Generate a response table for reviewer comments

Creates a structured table for responding to reviewer comments, commonly
needed for journal manuscript revisions. The output includes columns for
documenting your response and actions taken.

## Usage

``` r
generate_response_table(
  comments,
  output_path = NULL,
  group_by_reviewer = TRUE,
  include_resolved = FALSE
)
```

## Arguments

- comments:

  A data frame from
  [`extract_comments`](https://alexr-genetics.github.io/wordcomments/reference/extract_comments.md),
  or a path to a .docx file.

- output_path:

  Optional path to save the output (.xlsx, .docx, or .csv). If NULL,
  only returns the data frame.

- group_by_reviewer:

  Logical. If TRUE (default), groups comments by reviewer with section
  headers in the output.

- include_resolved:

  Logical. If FALSE (default), excludes resolved comments from the
  response table.

## Value

A data frame with columns:

- \#:

  Comment number

- Reviewer:

  Author of the comment

- Location:

  Page and line reference

- Commented Text:

  The text that was commented on

- Reviewer Comment:

  The comment content

- Author Response:

  Empty column for your response

- Action Taken:

  Empty column to describe changes made

- Line in Revision:

  Empty column for new line references

## Examples

``` r
if (FALSE) { # \dontrun{
# Generate and export response table
response <- generate_response_table("manuscript.docx",
                                    "response_to_reviewers.xlsx")

# Without export, just get the data frame
response <- generate_response_table("manuscript.docx")
} # }
```
