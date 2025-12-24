# wordcomments 0.2.0

## New Features

* `extract_comments_multiple()` - Extract and merge comments from multiple document versions with automatic ID renumbering to avoid conflicts
* `has_comments_multiple()` - Quick check for comments across multiple documents
* `launch_app()` / `run_wordcomments()` - Interactive Shiny GUI for all wordcomments functions

## Shiny GUI Features

* File browser for single or multiple .docx files
* Interactive tables with filtering and sorting
* Tabs for Comments, Summary, By Reviewer, Threads, and Response Table
* Export to Excel, Word, CSV, Markdown, or HTML
* Real-time status and logging

## Documentation

* Added vignette: "Introduction to wordcomments"
* Added vignette: "Working with Multiple Co-authors"
* Included example documents in `inst/extdata/`


## Notes

* Output from `extract_comments_multiple()` is fully compatible with all existing functions
* GUI requires additional packages: shiny, bslib, DT (installed automatically if missing)

# wordcomments 0.1.0

## New Features

* Initial CRAN release
* `extract_comments()` - Extract comments from Word documents with full metadata
* `has_comments()` - Quick check for presence of comments
* `comment_summary()` - Generate summary statistics for comments
* `comments_by_reviewer()` - Filter or split comments by author
* `find_comment_threads()` - Group comments into conversation threads
* `generate_response_table()` - Create reviewer response documents for journal revisions
* `export_comments()` - Export to Excel, Word, CSV, Markdown, or HTML

## Supported Formats

* Input: Microsoft Word (.docx) files
* Output: Excel (.xlsx), Word (.docx), CSV, Markdown (.md), HTML

## Notes

* Line numbers represent paragraph positions (Word doesn't store true line numbers)
* Page numbers are estimated based on paragraph count
* Resolution status detection supports multiple Word versions
