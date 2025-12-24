# Changelog

## wordcomments 0.1.0

### New Features

- Initial CRAN release
- [`extract_comments()`](https://alexr-genetics.github.io/wordcomments/reference/extract_comments.md) -
  Extract comments from Word documents with full metadata
- [`has_comments()`](https://alexr-genetics.github.io/wordcomments/reference/has_comments.md) -
  Quick check for presence of comments
- [`comment_summary()`](https://alexr-genetics.github.io/wordcomments/reference/comment_summary.md) -
  Generate summary statistics for comments
- [`comments_by_reviewer()`](https://alexr-genetics.github.io/wordcomments/reference/comments_by_reviewer.md) -
  Filter or split comments by author
- [`find_comment_threads()`](https://alexr-genetics.github.io/wordcomments/reference/find_comment_threads.md) -
  Group comments into conversation threads
- [`generate_response_table()`](https://alexr-genetics.github.io/wordcomments/reference/generate_response_table.md) -
  Create reviewer response documents for journal revisions
- [`export_comments()`](https://alexr-genetics.github.io/wordcomments/reference/export_comments.md) -
  Export to Excel, Word, CSV, Markdown, or HTML

### Supported Formats

- Input: Microsoft Word (.docx) files
- Output: Excel (.xlsx), Word (.docx), CSV, Markdown (.md), HTML

### Notes

- Line numbers represent paragraph positions (Word doesn’t store true
  line numbers)
- Page numbers are estimated based on paragraph count
- Resolution status detection supports multiple Word versions
