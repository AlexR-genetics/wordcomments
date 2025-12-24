# wordcomments

**wordcomments** is an R package for extracting, analyzing, and managing
comments from Microsoft Word (.docx) documents. It’s particularly useful
for academic manuscript revision workflows and collaborative document
review.

## Features

- 📄 **Extract comments** with full metadata (author, date, position,
  resolution status)
- 📚 **Multi-document support** - merge comments from multiple reviewer
  versions
- 🧵 **Thread detection** for comment conversations and replies
- 👥 **Filter by reviewer** for multi-author documents
- 📊 **Summary statistics** on comment activity
- 📝 **Generate response tables** for journal revisions
- 📤 **Export** to Excel, Word, CSV, Markdown, or HTML

## Installation

Install the development version from GitHub:

``` r
# install.packages("devtools")
devtools::install_github("AlexR-genetics/wordcomments")
```

## Graphical User Interface

wordcomments includes an interactive Shiny GUI:

``` r
library(wordcomments)
launch_app()
```

This opens a browser-based interface with: - File browser for single or
multiple documents - Interactive tables with filtering and sorting -
Tabs for Comments, Summary, By Reviewer, Threads, and Response Table -
Export to multiple formats

## Quick Start

``` r
library(wordcomments)

# Check if a document has comments
has_comments("manuscript.docx")

# Extract all comments
comments <- extract_comments("manuscript.docx")

# View summary statistics
comment_summary(comments)
```

## Usage Examples

### Extract Comments

``` r
# Extract all comments
comments <- extract_comments("manuscript.docx")

# Extract only unresolved comments
open_comments <- extract_comments("manuscript.docx", include_resolved = FALSE)
```

The returned data frame includes: \| Column \| Description \|
\|——–\|————-\| \| Text \| The text that was commented on \| \| Comment
\| The comment content \| \| Author \| Comment author \| \| Date \|
Timestamp \| \| Line \| Paragraph number (position in document) \| \|
Page \| Estimated page number \| \| Resolved \| Whether the comment
thread is resolved \|

### Summary Statistics

``` r
summary <- comment_summary(comments)
print(summary)

# Output:
# === COMMENT SUMMARY ===
# 
# OVERVIEW
#   Total comments:    24
#   Resolved:          8 (33.3%)
#   Unresolved:        16
#   Unique authors:    3
#
# BY AUTHOR
#   Reviewer 1           12 comments (4 resolved, 8 open)
#   Reviewer 2            8 comments (3 resolved, 5 open)
#   Editor                4 comments (1 resolved, 3 open)
```

### Filter by Reviewer

``` r
# Get comments from a specific reviewer
reviewer1 <- comments_by_reviewer(comments, reviewer = "Smith")

# Split into separate data frames by reviewer
by_reviewer <- comments_by_reviewer(comments, split = TRUE)
names(by_reviewer)  # Lists all reviewers
```

### Find Comment Threads

``` r
# Add thread information to comments
threaded <- find_comment_threads(comments)
# Adds: thread_id, thread_size, reply_depth

# Get nested list structure
threads <- find_comment_threads(comments, flatten = FALSE)
```

### Generate Response Table

Perfect for journal manuscript revisions:

``` r
# Generate and export response table
generate_response_table("manuscript.docx", "response_to_reviewers.xlsx")
```

Creates a table with columns: - \# (comment number) - Reviewer -
Location (page/line) - Commented Text - Reviewer Comment - Author
Response (empty - for you to fill) - Action Taken (empty - for you to
fill) - Line in Revision (empty - for you to fill)

### Export Comments

``` r
# Export to various formats
export_comments(comments, "comments.xlsx")
export_comments(comments, "comments.docx")
export_comments(comments, "comments.csv")
export_comments(comments, "comments.md")
export_comments(comments, "comments.html")

# Export only unresolved comments
export_comments(comments, "open_comments.xlsx", include_resolved = FALSE)

# Export specific columns
export_comments(comments, "simple.csv", 
                columns = c("Author", "Comment", "Resolved"))
```

    ## Multi-Document Workflow

    When you have multiple reviewers working on separate copies of a document:

    ```r
    # Extract and merge comments from all reviewer versions
    all_comments <- extract_comments_multiple(
      c("manuscript_reviewer1.docx",
        "manuscript_reviewer2.docx",
        "manuscript_editor.docx")
    )

    # Works with all other functions
    comment_summary(all_comments)
    comments_by_reviewer(all_comments, split = TRUE)

    # Generate consolidated response table
    generate_response_table(all_comments, "consolidated_response.xlsx")

    # Optional: track which document each comment came from
    all_comments <- extract_comments_multiple(
      c("reviewer1.docx", "reviewer2.docx"),
      add_source = TRUE
    )
    table(all_comments$Source)

## Documentation

For detailed examples, see the vignettes:

``` r
# Introduction to wordcomments
vignette("introduction", package = "wordcomments")

# Working with multiple co-authors
vignette("multiple-coauthors", package = "wordcomments")
```

## Technical Notes

- **Line numbers**: Word doesn’t store true line numbers internally. The
  `Line` column represents paragraph position, which serves as a
  reliable reference point.
- **Page numbers**: These are estimated (~25 paragraphs per page) since
  Word calculates pages at render time based on fonts/margins.
- **Resolution status**: The package checks multiple XML files
  (`commentsExtended.xml`, `commentsExtensible.xml`) to detect resolved
  comments across different Word versions.

## Dependencies

- xml2
- officer  
- openxlsx
- dplyr
- stringr
- tidyr
- rlang

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT © Alexandros Rammos

## Citation

If you use this package in your research, please cite it:

    @software{wordcomments,
      author = {Rammos, Alexandros},
      title = {wordcomments: Extract and Analyze Comments from Word Documents},
      year = {2025},
      url = {https://github.com/AlexR-genetics/wordcomments}
    }
