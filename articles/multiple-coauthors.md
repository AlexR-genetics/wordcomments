# Working with Multiple Co-authors

## Overview

When collaborating on a manuscript, different co-authors often work on
separate copies of the document. This vignette shows how to merge and
analyze comments from multiple document versions.

## Setup

``` r
library(wordcomments)
```

We’ll use three example documents, each with comments from a different
co-author:

``` r
# Get paths to example documents
doc1 <- system.file("extdata", "doc_v1.docx", package = "wordcomments")
doc2 <- system.file("extdata", "doc_v2.docx", package = "wordcomments")
doc3 <- system.file("extdata", "doc_v3.docx", package = "wordcomments")

all_docs <- c(doc1, doc2, doc3)
```

## Checking Multiple Documents

First, check which documents have comments:

``` r
has_comments_multiple(all_docs)
#>          file                                                             path
#> 1 doc_v1.docx /home/runner/work/_temp/Library/wordcomments/extdata/doc_v1.docx
#> 2 doc_v2.docx /home/runner/work/_temp/Library/wordcomments/extdata/doc_v2.docx
#> 3 doc_v3.docx /home/runner/work/_temp/Library/wordcomments/extdata/doc_v3.docx
#>   has_comments
#> 1         TRUE
#> 2         TRUE
#> 3         TRUE
```

## Extracting from Multiple Documents

Use
[`extract_comments_multiple()`](https://alexr-genetics.github.io/wordcomments/reference/extract_comments_multiple.md)
to merge comments from all documents:

``` r
all_comments <- extract_comments_multiple(all_docs)
#> Processing [1/3]: doc_v1.docx
#>   Found 3 comments
#> Processing [2/3]: doc_v2.docx
#>   Found 4 comments
#> Processing [3/3]: doc_v3.docx
#>   Found 2 comments
#> 
#> Total: 9 comments from 3 documents
```

This returns a single data frame with comments from all documents.
Comment IDs are automatically renumbered to avoid conflicts.

``` r
nrow(all_comments)
#> [1] 9
```

## Summary Across All Documents

``` r
comment_summary(all_comments)
#> === COMMENT SUMMARY ===
#> 
#> OVERVIEW
#>   Total comments:    9
#>   Resolved:          2 (22.2%)
#>   Unresolved:        7
#>   Unique authors:    3
#>   Comment threads:   8 root + 1 replies
#> 
#> BY AUTHOR
#>   James Morrison         4 comments (0 resolved, 4 open)
#>   Sarah Chen             3 comments (1 resolved, 2 open)
#>   Maria Garcia           2 comments (1 resolved, 1 open)
#> 
#> BY PAGE
#>   p1:9
#> 
#> DATE RANGE
#>   From: 2025-12-24 13:43
#>   To:   2025-12-24 13:48
#>   Span: 0.0 days
#> 
#> COMMENT LENGTH
#>   Mean: 293 chars | Median: 289 | Range: 241-377
```

## Splitting by Co-author

View comments organized by author:

``` r
by_author <- comments_by_reviewer(all_comments, split = TRUE)
names(by_author)
#> [1] "Sarah Chen"     "James Morrison" "Maria Garcia"
```

``` r
# View one author's comments
by_author[["Sarah Chen"]][, c("Comment", "Line", "Resolved")]
#>                                                                                                                                                                                                                                                                                                                                                                                     Comment
#> 1 The introduction effectively sets the historical context, but it could benefit from more nuance regarding the promotional stunts involving Chaplin (e.g., Tzara's false announcements of his "adhesion" to Dada). Consider adding a brief mention of how these were strategic hoaxes to amplify Dada's absurdity rather than literal endorsements, to avoid overstating direct alignment.
#> 2                                                                                                      You mention "Chaplinades" as literary works thematizing Chaplin's persona—could you clarify specific examples (e.g., Yvan Goll's  Die Chaplinade  or Vítězslav Nezval's variations)? This would strengthen the qualitative analysis by grounding it in concrete Dada-adjacent texts.
#> 3                                                                                                                                         Minor typo/awkward phrasing: "centrality scores indicating influence strength" could be smoothed to "centrality scores to quantify the strength of Chaplin's perceived influence." Otherwise, the statistical framing is creatively applied here.
#>   Line Resolved
#> 1    3    FALSE
#> 2    5    FALSE
#> 3    7     TRUE
```

Filter to a specific co-author:

``` r
sarah_comments <- comments_by_reviewer(all_comments, reviewer = "Sarah")
nrow(sarah_comments)
#> [1] 3
```

## Finding Comment Threads

Detect reply chains (where one comment replies to another):

``` r
threads <- find_comment_threads(all_comments)
```

``` r
# View thread information
threads[, c("thread_id", "thread_size", "reply_depth", "Author", "Comment")]
#> # A tibble: 9 × 5
#>   thread_id thread_size reply_depth Author         Comment                      
#>       <dbl>       <int>       <dbl> <chr>          <chr>                        
#> 1         1           1           0 Sarah Chen     "The introduction effectivel…
#> 2         2           1           0 James Morrison "The claim that Dadaists use…
#> 3         3           1           0 Maria Garcia   "Solid background, but expan…
#> 4         4           1           0 Sarah Chen     "You mention \"Chaplinades\"…
#> 5         5           2           0 James Morrison "Strong qualitative approach…
#> 6         5           2           1 James Morrison "Building on my previous sug…
#> 7         6           1           0 Sarah Chen     "Minor typo/awkward phrasing…
#> 8         7           1           0 Maria Garcia   "The use of chi-square tests…
#> 9         8           1           0 James Morrison "The tools are modern and ap…
```

Comments with `thread_size > 1` are part of a conversation.

## Tracking Source Documents

To track which document each comment came from:

``` r
all_comments_with_source <- extract_comments_multiple(all_docs, add_source = TRUE)
#> Processing [1/3]: doc_v1.docx
#>   Found 3 comments
#> Processing [2/3]: doc_v2.docx
#>   Found 4 comments
#> Processing [3/3]: doc_v3.docx
#>   Found 2 comments
#> 
#> Total: 9 comments from 3 documents
```

``` r
table(all_comments_with_source$Source)
#> 
#> doc_v1.docx doc_v2.docx doc_v3.docx 
#>           3           4           2
```

## Generating a Consolidated Response Table

Create a single response document for all co-authors:

``` r
generate_response_table(all_comments, "consolidated_response.xlsx")
```

Or group responses by co-author:

``` r
generate_response_table(all_comments, "response_by_coauthor.xlsx", 
                        group_by_reviewer = TRUE)
```

## Exporting Merged Comments

``` r
# Export all comments to Excel
export_comments(all_comments, "all_coauthor_comments.xlsx")

# Export only unresolved
export_comments(all_comments, "open_comments.xlsx", include_resolved = FALSE)
```

## Using the GUI

The Shiny GUI supports multiple file selection:

``` r
launch_app()
```

In the GUI:

1.  Click “Browse” and select multiple .docx files (Ctrl+click)
2.  Check “Add source column” to track document origins
3.  Click “Extract Comments”
4.  Use the tabs to explore and export

## Typical Workflow

``` r
# 1. Collect document versions from co-authors
docs <- c("draft_sarah.docx", "draft_james.docx", "draft_maria.docx")

# 2. Merge all comments
all_comments <- extract_comments_multiple(docs)

# 3. Review summary
comment_summary(all_comments)

# 4. Generate response table
generate_response_table(all_comments, "response_to_coauthors.xlsx")

# 5. Fill in responses in Excel and share with co-authors
```
