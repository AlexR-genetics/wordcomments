# Check if a Word document contains comments

Performs a fast check for the presence of comments without full parsing.
Useful for filtering files before batch processing.

## Usage

``` r
has_comments(docx_path)
```

## Arguments

- docx_path:

  Path to the .docx file

## Value

Logical. TRUE if the document contains at least one comment, FALSE
otherwise.

## Examples

``` r
if (FALSE) { # \dontrun{
# Quick check before processing
if (has_comments("document.docx")) {
  comments <- extract_comments("document.docx")
}

# Filter a list of files
files <- list.files(pattern = "\\.docx$")
files_with_comments <- files[sapply(files, has_comments)]
} # }
```
