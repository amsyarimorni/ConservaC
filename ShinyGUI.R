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
      textInput("password", "Enter Excel Password", placeholder = "Password"),
      actionButton("run_button", "Run Excel Operations")
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
    req(input$species_query_file, input$master_data_file, input$password)
    
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

shinyApp(ui, server)
