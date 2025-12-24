# Quick check for comments in multiple documents

Checks multiple documents for the presence of comments and returns a
summary.

## Usage

``` r
has_comments_multiple(docx_paths)
```

## Arguments

- docx_paths:

  Character vector of paths to .docx files

## Value

A data frame with columns:

- file:

  Document filename

- path:

  Full path to document

- has_comments:

  Logical indicating if document has comments

## Examples

``` r
if (FALSE) { # \dontrun{
# Check which documents have comments
status <- has_comments_multiple(
  c("doc1.docx", "doc2.docx", "doc3.docx")
)
print(status)

# Filter to only documents with comments
docs_with_comments <- status$path[status$has_comments]
} # }
```
