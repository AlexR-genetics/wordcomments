# wordcomments: Extract and Analyze Comments from Word Documents

A toolkit for extracting, analyzing, and managing comments from
Microsoft Word (.docx) documents. Particularly useful for academic
manuscript revision workflows and collaborative document review.

## Graphical User Interface

For an interactive experience, launch the Shiny GUI:

    launch_app()

## Main Functions

- [`launch_app`](https://alexr-genetics.github.io/wordcomments/reference/launch_app.md):

  Launch interactive Shiny GUI

- [`extract_comments`](https://alexr-genetics.github.io/wordcomments/reference/extract_comments.md):

  Extract all comments with full metadata

- [`extract_comments_multiple`](https://alexr-genetics.github.io/wordcomments/reference/extract_comments_multiple.md):

  Extract and merge comments from multiple documents

- [`has_comments`](https://alexr-genetics.github.io/wordcomments/reference/has_comments.md):

  Quick check for presence of comments

- [`has_comments_multiple`](https://alexr-genetics.github.io/wordcomments/reference/has_comments_multiple.md):

  Quick check across multiple documents

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

## Multi-Document Workflow

    # Extract comments from multiple reviewer versions
    all_comments <- extract_comments_multiple(
      c("reviewer1.docx", "reviewer2.docx", "editor.docx")
    )

    # Works with all other functions
    comment_summary(all_comments)
    generate_response_table(all_comments, "consolidated_response.xlsx")

## See also

Useful links:

- <https://github.com/AlexR-genetics/wordcomments>

- Report bugs at <https://github.com/AlexR-genetics/wordcomments/issues>

## Author

**Maintainer**: Alexandros Rammos <vd18986@bristol.ac.uk>
([ORCID](https://orcid.org/0000-0001-7491-9659))
