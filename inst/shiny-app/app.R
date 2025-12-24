# wordcomments Shiny App
# A GUI for extracting and analyzing Word document comments

library(shiny)
library(bslib)
library(DT)
library(wordcomments)

# ============================================================================
# UI
# ============================================================================

ui <- page_sidebar(
  title = "wordcomments",
  theme = bs_theme(
    version = 5,
    bootswatch = "flatly",
    primary = "#3498db",
    "navbar-bg" = "#2c3e50"
  ),
  
  # --- Sidebar: Inputs & Controls ---
  sidebar = sidebar(
    width = 350,
    
    # File Input Section
    card(
      card_header(
        class = "bg-primary text-white",
        "📂 Input Files"
      ),
      card_body(
        fileInput(
          "files",
          label = NULL,
          multiple = TRUE,
          accept = c(".docx"),
          placeholder = "Select .docx file(s)"
        ),
        verbatimTextOutput("file_list", placeholder = TRUE),
        actionButton("clear_files", "Clear Files", 
                     class = "btn-outline-secondary btn-sm mt-2")
      )
    ),
    
    # Options Section
    card(
      card_header(
        class = "bg-primary text-white",
        "⚙️ Options"
      ),
      card_body(
        checkboxInput("include_resolved", "Include resolved comments", value = TRUE),
        checkboxInput("add_source", "Add source column (multi-file)", value = FALSE),
        hr(),
        selectInput(
          "export_format",
          "Export Format:",
          choices = c("Excel (.xlsx)" = "excel",
                      "Word (.docx)" = "word", 
                      "CSV (.csv)" = "csv",
                      "Markdown (.md)" = "markdown",
                      "HTML (.html)" = "html"),
          selected = "excel"
        )
      )
    ),
    
    # Actions Section
    card(
      card_header(
        class = "bg-primary text-white",
        "▶️ Actions"
      ),
      card_body(
        actionButton("btn_extract", "Extract Comments", 
                     class = "btn-success w-100 mb-2",
                     icon = icon("file-import")),
        actionButton("btn_summary", "Show Summary", 
                     class = "btn-info w-100 mb-2",
                     icon = icon("chart-bar")),
        actionButton("btn_response", "Generate Response Table", 
                     class = "btn-warning w-100 mb-2",
                     icon = icon("table")),
        hr(),
        downloadButton("btn_export", "Export Results", 
                       class = "btn-primary w-100",
                       icon = icon("download"))
      )
    )
  ),
  
  # --- Main Panel: Results ---
  layout_columns(
    col_widths = c(12),
    
    # Status Card
    card(
      card_header("📊 Status"),
      card_body(
        uiOutput("status_message")
      ),
      height = "auto"
    )
  ),
  
  # Tabbed Results Area
  navset_card_tab(
    id = "results_tabs",
    
    # Comments Table Tab
    nav_panel(
      title = "Comments",
      icon = icon("comments"),
      card_body(
        DTOutput("comments_table")
      )
    ),
    
    # Summary Tab
    nav_panel(
      title = "Summary",
      icon = icon("chart-pie"),
      card_body(
        verbatimTextOutput("summary_output")
      )
    ),
    
    # By Reviewer Tab
    nav_panel(
      title = "By Reviewer",
      icon = icon("users"),
      card_body(
        selectInput("reviewer_select", "Select Reviewer:", choices = NULL),
        DTOutput("reviewer_table")
      )
    ),
    
    # Threads Tab
    nav_panel(
      title = "Threads",
      icon = icon("project-diagram"),
      card_body(
        DTOutput("threads_table")
      )
    ),
    
    # Response Table Tab
    nav_panel(
      title = "Response Table",
      icon = icon("reply-all"),
      card_body(
        checkboxInput("response_resolved", "Include resolved in response", value = FALSE),
        checkboxInput("response_group", "Group by reviewer", value = TRUE),
        DTOutput("response_table")
      )
    ),
    
    # Log Tab
    nav_panel(
      title = "Log",
      icon = icon("terminal"),
      card_body(
        verbatimTextOutput("log_output")
      )
    )
  )
)


# ============================================================================
# SERVER
# ============================================================================

server <- function(input, output, session) {
  
  # --- Reactive Values ---
  rv <- reactiveValues(
    comments = NULL,
    summary = NULL,
    threads = NULL,
    response = NULL,
    log = character(),
    files_loaded = character()
  )
  
  # --- Helper: Add to log ---
  add_log <- function(msg) {
    timestamp <- format(Sys.time(), "[%H:%M:%S]")
    rv$log <- c(rv$log, paste(timestamp, msg))
  }
  
  # --- File List Display ---
  output$file_list <- renderText({
    req(input$files)
    paste(input$files$name, collapse = "\n")
  })
  
  # --- Stop app when browser window closes ---
  session$onSessionEnded(function() {
    stopApp()
  })
  
  # --- Clear Files ---
  observeEvent(input$clear_files, {
    rv$comments <- NULL
    rv$summary <- NULL
    rv$threads <- NULL
    rv$response <- NULL
    rv$files_loaded <- character()
    add_log("Files cleared")
    
    # Reset file input
    session$sendCustomMessage("resetFileInput", "files")
  })
  
  # --- Status Message ---
  output$status_message <- renderUI({
    if (is.null(rv$comments)) {
      div(
        class = "alert alert-secondary",
        icon("info-circle"), 
        " Select one or more .docx files and click 'Extract Comments' to begin."
      )
    } else {
      n_files <- length(rv$files_loaded)
      n_comments <- nrow(rv$comments)
      n_resolved <- sum(rv$comments$Resolved, na.rm = TRUE)
      n_unresolved <- n_comments - n_resolved
      n_authors <- length(unique(rv$comments$Author))
      
      div(
        class = "alert alert-success",
        icon("check-circle"),
        sprintf(" Loaded %d comments from %d file(s) | %d resolved, %d open | %d reviewers",
                n_comments, n_files, n_resolved, n_unresolved, n_authors)
      )
    }
  })
  
  # --- Extract Comments ---
  observeEvent(input$btn_extract, {
    req(input$files)
    
    add_log(sprintf("Extracting comments from %d file(s)...", nrow(input$files)))
    
    tryCatch({
      file_paths <- input$files$datapath
      file_names <- input$files$name
      
      # Use single or multiple extraction based on file count
      if (length(file_paths) == 1) {
        rv$comments <- extract_comments(
          file_paths,
          include_resolved = input$include_resolved
        )
        add_log(sprintf("Extracted %d comments from %s", nrow(rv$comments), file_names))
      } else {
        rv$comments <- extract_comments_multiple(
          file_paths,
          include_resolved = input$include_resolved,
          add_source = input$add_source
        )
        add_log(sprintf("Extracted %d total comments from %d files", 
                        nrow(rv$comments), length(file_paths)))
      }
      
      rv$files_loaded <- file_names
      
      # Update reviewer dropdown
      reviewers <- unique(rv$comments$Author)
      updateSelectInput(session, "reviewer_select", 
                        choices = c("All" = "", reviewers))
      
      # Auto-generate summary
      if (nrow(rv$comments) > 0) {
        rv$summary <- comment_summary(rv$comments)
        rv$threads <- find_comment_threads(rv$comments)
        add_log("Summary and threads computed")
      }
      
      showNotification("Comments extracted successfully!", type = "message")
      
    }, error = function(e) {
      add_log(sprintf("ERROR: %s", e$message))
      showNotification(paste("Error:", e$message), type = "error")
    })
  })
  
  # --- Show Summary ---
  observeEvent(input$btn_summary, {
    req(rv$comments)
    bslib::nav_select("results_tabs", selected = "Summary")
  })
  
  # --- Generate Response Table ---
  observeEvent(input$btn_response, {
    req(rv$comments)
    
    tryCatch({
      rv$response <- generate_response_table(
        rv$comments,
        include_resolved = input$response_resolved,
        group_by_reviewer = input$response_group
      )
      add_log(sprintf("Response table generated with %d rows", nrow(rv$response)))
      bslib::nav_select("results_tabs", selected = "Response Table")
      showNotification("Response table generated!", type = "message")
    }, error = function(e) {
      add_log(sprintf("ERROR: %s", e$message))
      showNotification(paste("Error:", e$message), type = "error")
    })
  })
  
  # --- Comments Table Output ---
  output$comments_table <- renderDT({
    req(rv$comments)
    
    # Select display columns (exclude internal IDs)
    display_cols <- c("Text", "Comment", "Author", "Date", "Line", "Page", "Resolved")
    if ("Source" %in% names(rv$comments)) {
      display_cols <- c(display_cols, "Source")
    }
    
    df <- rv$comments[, display_cols, drop = FALSE]
    df$Date <- format(df$Date, "%Y-%m-%d %H:%M")
    
    datatable(
      df,
      options = list(
        pageLength = 10,
        scrollX = TRUE,
        dom = 'Bfrtip',
        buttons = c('copy', 'csv', 'excel')
      ),
      filter = "top",
      rownames = FALSE,
      class = "compact stripe hover"
    )
  })
  
  # --- Summary Output ---
  output$summary_output <- renderPrint({
    req(rv$summary)
    print(rv$summary)
  })
  
  # --- Reviewer Table ---
  output$reviewer_table <- renderDT({
    req(rv$comments)
    
    df <- if (input$reviewer_select == "" || is.null(input$reviewer_select)) {
      rv$comments
    } else {
      comments_by_reviewer(rv$comments, reviewer = input$reviewer_select)
    }
    
    display_cols <- c("Text", "Comment", "Author", "Date", "Line", "Page", "Resolved")
    df <- df[, display_cols, drop = FALSE]
    df$Date <- format(df$Date, "%Y-%m-%d %H:%M")
    
    datatable(
      df,
      options = list(pageLength = 10, scrollX = TRUE),
      filter = "top",
      rownames = FALSE,
      class = "compact stripe hover"
    )
  })
  
  # --- Threads Table ---
  output$threads_table <- renderDT({
    req(rv$threads)
    
    display_cols <- c("thread_id", "thread_size", "reply_depth", 
                      "Text", "Comment", "Author", "Date", "Resolved")
    display_cols <- intersect(display_cols, names(rv$threads))
    
    df <- rv$threads[, display_cols, drop = FALSE]
    if ("Date" %in% names(df)) {
      df$Date <- format(df$Date, "%Y-%m-%d %H:%M")
    }
    
    datatable(
      df,
      options = list(pageLength = 10, scrollX = TRUE),
      filter = "top",
      rownames = FALSE,
      class = "compact stripe hover"
    )
  })
  
  # --- Response Table ---
  output$response_table <- renderDT({
    req(rv$response)
    
    datatable(
      rv$response,
      options = list(pageLength = 10, scrollX = TRUE),
      filter = "top",
      rownames = FALSE,
      class = "compact stripe hover",
      editable = TRUE  # Allow editing responses directly!
    )
  })
  
  # --- Log Output ---
  output$log_output <- renderText({
    paste(rv$log, collapse = "\n")
  })
  
  # --- Export Download Handler ---
  output$btn_export <- downloadHandler(
    filename = function() {
      # Map format to file extension
      ext <- switch(input$export_format,
                    "excel" = "xlsx",
                    "word" = "docx",
                    "csv" = "csv",
                    "markdown" = "md",
                    "html" = "html",
                    "xlsx")  # default
      paste0("wordcomments_export_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".", ext)
    },
    content = function(file) {
      req(rv$comments)
      
      add_log(sprintf("Exporting to %s format...", input$export_format))
      
      # Check for required packages based on format
      fmt <- input$export_format
      if (fmt == "excel" && !requireNamespace("openxlsx", quietly = TRUE)) {
        showNotification("Excel export requires 'openxlsx' package. Install with: install.packages('openxlsx')", 
                         type = "error", duration = 10)
        add_log("ERROR: openxlsx package not installed")
        # Create a fallback CSV
        write.csv(rv$comments, file, row.names = FALSE)
        return()
      }
      
      if (fmt == "word" && !requireNamespace("officer", quietly = TRUE)) {
        showNotification("Word export requires 'officer' package. Install with: install.packages('officer')", 
                         type = "error", duration = 10)
        add_log("ERROR: officer package not installed")
        # Create a fallback CSV
        write.csv(rv$comments, file, row.names = FALSE)
        return()
      }
      
      tryCatch({
        export_comments(
          rv$comments,
          file,
          format = fmt,
          include_resolved = input$include_resolved
        )
        add_log(sprintf("Exported successfully to %s", basename(file)))
      }, error = function(e) {
        add_log(sprintf("Export ERROR: %s", e$message))
        showNotification(paste("Export error:", e$message), type = "error", duration = 10)
        # Fallback: write as CSV so download doesn't completely fail
        tryCatch({
          write.csv(rv$comments, file, row.names = FALSE)
          add_log("Fallback: exported as CSV")
        }, error = function(e2) {
          add_log(sprintf("Fallback also failed: %s", e2$message))
        })
      })
    }
  )
}


# ============================================================================
# RUN APP
# ============================================================================

shinyApp(ui = ui, server = server)
