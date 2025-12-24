# Group comments into conversation threads

Identifies and groups comments that are part of the same conversation
thread, including replies to comments.

## Usage

``` r
find_comment_threads(comments, flatten = TRUE)
```

## Arguments

- comments:

  A data frame from
  [`extract_comments`](https://alexr-genetics.github.io/wordcomments/reference/extract_comments.md),
  or a path to a .docx file.

- flatten:

  Logical. If TRUE (default), returns a flat data frame with thread_id
  column added. If FALSE, returns a nested list structure.

## Value

If `flatten = TRUE`: a data frame with additional columns:

- thread_id:

  Integer identifying the conversation thread

- thread_size:

  Number of comments in this thread

- reply_depth:

  Nesting depth (0 for root comments)

If `flatten = FALSE`: a list of thread objects, each containing:

- thread_id:

  Thread identifier

- root_comment:

  Data frame with the original comment

- replies:

  Data frame with reply comments

- size:

  Total comments in thread

- authors:

  Character vector of participating authors

- resolved:

  Logical indicating if thread is resolved

## Examples

``` r
if (FALSE) { # \dontrun{
comments <- extract_comments("manuscript.docx")

# Flat format with thread IDs
threaded <- find_comment_threads(comments)

# Nested list format
threads <- find_comment_threads(comments, flatten = FALSE)
} # }
```
