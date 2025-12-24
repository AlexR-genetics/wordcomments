# Generate summary statistics for comments

Produces a comprehensive summary of comment data including counts by
author, resolution rates, date ranges, and comment length statistics.

## Usage

``` r
comment_summary(comments)
```

## Arguments

- comments:

  A data frame from
  [`extract_comments`](https://alexr-genetics.github.io/wordcomments/reference/extract_comments.md),
  or a path to a .docx file (which will be processed automatically).

## Value

A list of class "comment_summary" containing:

- total_comments:

  Total number of comments

- resolved:

  Number of resolved comments

- unresolved:

  Number of unresolved comments

- resolution_rate:

  Percentage of comments resolved

- unique_authors:

  Number of unique comment authors

- by_author:

  Data frame with counts per author

- by_page:

  Data frame with counts per page

- date_range:

  List with earliest, latest dates and span in days

- comment_length:

  Statistics on comment text length

- threads:

  Information about comment threading

## Examples

``` r
if (FALSE) { # \dontrun{
# From file
summary <- comment_summary("manuscript.docx")
print(summary)

# From existing comments data frame
comments <- extract_comments("manuscript.docx")
summary <- comment_summary(comments)
} # }
```
