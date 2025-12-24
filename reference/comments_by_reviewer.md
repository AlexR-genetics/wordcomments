# Filter or split comments by reviewer/author

Filters comments to a specific reviewer or splits all comments into
separate data frames by author.

## Usage

``` r
comments_by_reviewer(comments, reviewer = NULL, split = FALSE)
```

## Arguments

- comments:

  A data frame from
  [`extract_comments`](https://alexr-genetics.github.io/wordcomments/reference/extract_comments.md),
  or a path to a .docx file.

- reviewer:

  Optional character string. If provided, filters to comments from this
  reviewer (case-insensitive partial matching).

- split:

  Logical. If TRUE and no reviewer specified, returns a named list of
  data frames split by author. Default FALSE.

## Value

If `reviewer` is specified or `split = FALSE`: a data frame. If
`split = TRUE`: a named list of data frames, one per author.

## Examples

``` r
if (FALSE) { # \dontrun{
comments <- extract_comments("manuscript.docx")

# Filter to one reviewer
smith_comments <- comments_by_reviewer(comments, reviewer = "Smith")

# Split by all reviewers
by_reviewer <- comments_by_reviewer(comments, split = TRUE)
names(by_reviewer)  # Lists all reviewer names
} # }
```
