# wordcomments: Extract and Analyze Comments from Word Documents

A toolkit for extracting, analyzing, and managing comments from
Microsoft Word (.docx) documents. Particularly useful for academic
manuscript revision workflows and collaborative document review.

## Main Functions

- [`extract_comments`](https://alexr-genetics.github.io/wordcomments/reference/extract_comments.md):

  Extract all comments with full metadata

- [`has_comments`](https://alexr-genetics.github.io/wordcomments/reference/has_comments.md):

  Quick check for presence of comments

- [`comment_summary`](https://alexr-genetics.github.io/wordcomments/reference/comment_summary.md):

  Generate summary statistics

- [`comments_by_reviewer`](https://alexr-genetics.github.io/wordcomments/reference/comments_by_reviewer.md):

  Filter or split comments by author

- [`find_comment_threads`](https://alexr-genetics.github.io/wordcomments/reference/find_comment_threads.md):

  Group comments into conversation threads

- [`generate_response_table`](https://alexr-genetics.github.io/wordcomments/reference/generate_response_table.md):

  Create reviewer response document

- [`export_comments`](https://alexr-genetics.github.io/wordcomments/reference/export_comments.md):

  Export to Excel, Word, CSV, Markdown, or HTML

## Typical Workflow

    # Check for comments
    has_comments("manuscript.docx")

    # Extract and summarize
    comments <- extract_comments("manuscript.docx")
    comment_summary(comments)

    # Generate response table for journal revision
    generate_response_table(comments, "response_to_reviewers.xlsx")

## See also

Useful links:

- <https://github.com/AlexR-genetics/wordcomments>

- Report bugs at <https://github.com/AlexR-genetics/wordcomments/issues>

## Author

**Maintainer**: Alexandros Rammos <vd18986@bristol.ac.uk>
