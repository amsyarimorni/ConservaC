# Load the required libraries
library(shiny)
library(writexl)
library(readxl)

ui <- fluidPage(
  titlePanel("Conservac"),
  sidebarLayout(
    sidebarPanel(
      fileInput("species_query_file", "Upload Species Query Excel File"),
      fileInput("master_data_file", "Upload Master Data Excel File"),
      actionButton("run_button", "Run Excel Operations"),
      br(),
      HTML("<p>To ensure this package works properly, please do not make any modification in MasterData file</p>"),
      HTML("<p>No Match could indicate wrong species spelling or the queried species is not yet updated in our database</p>"),
      HTML("<p>Please contact me (mmamsyari@unimas.my) with list of No Match species for further clarification</p>"),
      HTML("<p>This package is in BETA mode, comments are highly appreciated</p>"),
      HTML("<p>-Morni M. A.</p>")
    ),
    mainPanel(
      downloadLink("download_result", "Download Result Excel File"),
      tableOutput("result_table"),
      textOutput("status_message")
    )
  )
)

server <- function(input, output) {
  observeEvent(input$run_button, {
    req(input$species_query_file, input$master_data_file)
    
    species_query <- read_xlsx(input$species_query_file$datapath)
    master_data <- read_xlsx(input$master_data_file$datapath)
    
    # Perform VLOOKUP using merge
    result <- merge(species_query, master_data, by = "Species", all.x = TRUE)
    
    # Replace NAs with "NOT AVAILABLE"
    result[is.na(result)] <- "NO MATCH"
    
    result_path <- "Result.xlsx"
    
    write_xlsx(result, result_path)
    
    output$download_result <- downloadHandler(
      filename = function() {
        "Result.xlsx"
      },
      content = function(file) {
        file.copy(result_path, file)
      }
    )
    
    output$result_table <- renderTable(result)
    output$status_message <- renderText({
      "Your query is completed, Database update: Sept 2023.- Morni M. A."
    })
  })
}

# Launch the Shiny app in a web browser by specifying the host and port
shinyApp(ui, server, options = list(host = "0.0.0.0", port = 8888))
