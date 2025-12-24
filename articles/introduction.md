# Introduction to wordcomments

## Overview

The `wordcomments` package extracts and analyzes comments from Microsoft
Word (.docx) documents. This is useful when collaborating on documents
with co-authors who provide feedback via Word’s comment feature.

## Setup

``` r
library(wordcomments)
```

For this vignette, we’ll use example documents included with the
package:

``` r
# Get path to example document
doc_path <- system.file("extdata", "doc_v1.docx", package = "wordcomments")
```

## Checking for Comments

Before extracting, you can quickly check if a document has any comments:

``` r
has_comments(doc_path)
#> [1] TRUE
```

## Extracting Comments

The main function is
[`extract_comments()`](https://alexr-genetics.github.io/wordcomments/reference/extract_comments.md):

``` r
comments <- extract_comments(doc_path)
```

This returns a data frame with all comment metadata:

``` r
# View structure
str(comments)
#> 'data.frame':    3 obs. of  10 variables:
#>  $ Text      : chr  "Charlie Chaplin, the iconic silent film comedian known for his Tramp character, emerged during the early 20th c"| __truncated__ "This study employed a qualitative content analysis of historical Dada texts, manifestos, and artworks referenci"| __truncated__ "To quantify Chaplin's prominence in Dada discourse, mentions of \"Chaplin\" or \"Charlot\" were counted across "| __truncated__
#>  $ Comment   : chr  "The introduction effectively sets the historical context, but it could benefit from more nuance regarding the p"| __truncated__ "You mention \"Chaplinades\" as literary works thematizing Chaplin's persona—could you clarify specific examples"| __truncated__ "Minor typo/awkward phrasing: \"centrality scores indicating influence strength\" could be smoothed to \"central"| __truncated__
#>  $ Author    : chr  "Sarah Chen" "Sarah Chen" "Sarah Chen"
#>  $ Date      : POSIXct, format: "2025-12-24 13:43:00" "2025-12-24 13:43:00" ...
#>  $ Line      : int  3 5 7
#>  $ Page      : num  1 1 1
#>  $ Resolved  : logi  FALSE FALSE TRUE
#>  $ comment_id: chr  "0" "1" "2"
#>  $ para_id   : chr  "19217B9E" "6172C83E" "70800F1A"
#>  $ parent_id : logi  NA NA NA
```

``` r
# View the comments (excluding internal ID columns)
comments[, c("Author", "Comment", "Line", "Resolved")]
#>       Author
#> 1 Sarah Chen
#> 2 Sarah Chen
#> 3 Sarah Chen
#>                                                                                                                                                                                                                                                                                                                                                                                     Comment
#> 1 The introduction effectively sets the historical context, but it could benefit from more nuance regarding the promotional stunts involving Chaplin (e.g., Tzara's false announcements of his "adhesion" to Dada). Consider adding a brief mention of how these were strategic hoaxes to amplify Dada's absurdity rather than literal endorsements, to avoid overstating direct alignment.
#> 2                                                                                                      You mention "Chaplinades" as literary works thematizing Chaplin's persona—could you clarify specific examples (e.g., Yvan Goll's  Die Chaplinade  or Vítězslav Nezval's variations)? This would strengthen the qualitative analysis by grounding it in concrete Dada-adjacent texts.
#> 3                                                                                                                                         Minor typo/awkward phrasing: "centrality scores indicating influence strength" could be smoothed to "centrality scores to quantify the strength of Chaplin's perceived influence." Otherwise, the statistical framing is creatively applied here.
#>   Line Resolved
#> 1    3    FALSE
#> 2    5    FALSE
#> 3    7     TRUE
```

## Comment Summary

Get an overview of comment statistics:

``` r
summary <- comment_summary(comments)
print(summary)
#> === COMMENT SUMMARY ===
#> 
#> OVERVIEW
#>   Total comments:    3
#>   Resolved:          1 (33.3%)
#>   Unresolved:        2
#>   Unique authors:    1
#> 
#> BY AUTHOR
#>   Sarah Chen             3 comments (1 resolved, 2 open)
#> 
#> BY PAGE
#>   p1:3
#> 
#> DATE RANGE
#>   From: 2025-12-24 13:43
#>   To:   2025-12-24 13:44
#>   Span: 0.0 days
#> 
#> COMMENT LENGTH
#>   Mean: 298 chars | Median: 276 | Range: 241-377
```

You can access individual summary elements:

``` r
summary$total_comments
#> [1] 3
summary$by_author
#> # A tibble: 1 × 4
#>   Author     count resolved unresolved
#>   <chr>      <int>    <int>      <int>
#> 1 Sarah Chen     3        1          2
```

## Filtering by Resolved Status

To get only unresolved comments:

``` r
unresolved <- extract_comments(doc_path, include_resolved = FALSE)
nrow(unresolved)
#> [1] 2
```

## Exporting Comments

Export comments to various formats:

``` r
# Excel
export_comments(comments, "comments.xlsx")

# Word document
export_comments(comments, "comments.docx")

# CSV
export_comments(comments, "comments.csv")

# Markdown
export_comments(comments, "comments.md")
```

## Generating a Response Table

For responding to co-author feedback, generate a structured response
table:

``` r
generate_response_table(comments, "response_to_coauthors.xlsx")
```

This creates a table with columns for your responses that you can fill
in and share.

## Using the GUI

For an interactive experience, launch the Shiny GUI:

``` r
launch_app()
```

This opens a browser-based interface where you can:

- Browse and select files
- View comments in interactive tables
- Filter and sort by various columns
- Export to different formats

## Next Steps

See the “Working with Multiple Co-authors” vignette for handling
comments from multiple document versions.
