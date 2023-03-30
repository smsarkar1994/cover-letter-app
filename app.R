options(shiny.sanitize.errors = FALSE)
library(dplyr)
library(stringr)
library(officer)
library(lubridate)
library(shiny)
library(shinythemes)

ui <- fluidPage(
  theme = shinytheme("superhero"),
  #Enlargen text
  tags$head(tags$style('
   body {
      font-size: 15px;
   }'
  )),
  tags$head(tags$style('
   label {
      font-size: 15px;
   }'
  )),
  #Jumbotrons are pretty, they make nice headers
  tags$div(class = "jumbotron text-center", style = "margin-bottom:12.5px;margin-top:0px",
           tags$h2(class = 'jumbotron-heading', style = 'margin-bottom:-75px;margin-top:-25px', 'Custom Cover Letter Tool')
  ),
  title = "Custom Cover Letter Tool",
  
  #Sidebar with inputs
  sidebarLayout(
    sidebarPanel(h4(""),
                 textInput('company', "Company Name", 
                           value = ""),
                 textInput('position', "Job Title", 
                           value = ""),
                 textInput('hiring_manager', "Hiring Manager", 
                           value = ""),
                 radioButtons('date_button', 'Date', 
                              c('Current Date' = 'date_current',
                                'Custom Date' = 'date_custom')),
                 conditionalPanel(condition = "input.date_button == 'date_custom'",
                                  dateInput("date", "Select Date:", value = Sys.Date())),
                 fileInput("input_file", HTML("Upload Cover Letter Template"),
                           multiple = FALSE,
                           accept = c(".docx")),
                 textInput('output_path', "Specify Output Folder", 
                           value = getwd()),
                 actionButton("do", "Generate Cover Letter", class = "btn-primary")
                 
    ),
    
    # Main panel for displaying outputs ----
    mainPanel(position = "right",
              h3(strong("Instructions")),
              p("Please upload your cover letter template and fill out all inputs."),
              p(strong("Important Notes:")),
              tags$ul(
                tags$li("Your cover letter template should specify placeholders for the date, hiring manager, company name, and job title wherever you want it changed as such:"),
                tags$ul(
                  tags$li("[date]"),
                  tags$li("[hiring manager]"),
                  tags$ul(
                    tags$li('If left blank, the app will use a generic', 
                            em('"Hiring Manager"'), 
                            'wherever the placeholder is specified.')
                  ),
                  tags$li("[company]"),
                  tags$li("[position]")
                ),
                tags$li(strong("The app only accepts .docx templates.")),
                tags$li("To download a sample cover letter template, click",
                        a(href="https://www.statswithsasa.com/wp-content/uploads/2023/03/cover_letter_template.docx", "here."))
              ),
              br(),
              h3(strong("Selected Inputs")),
              htmlOutput("error_message", style="color:red"),
              br(),
              htmlOutput("date_label"),
              br(), 
              htmlOutput("hiring_manager"),
              br(),
              htmlOutput("company_label"),
              br(),
              htmlOutput("position_label"),
              br(),
              p("If you like this tool and are interested
                           in checking out more cool stuff,", 
                strong("check out my personal site", 
                       a(href="https://www.statswithsasa.com/", "Stats with Sasa.")))
    )
  )
)

server <- function(input, output) {
  # font_factor = 17.5/15
  observeEvent(input$do, {
    
    #Only run if a file has been uploaded
    if(!is.null(input$input_file)) {
      
      #Error message if .docx file not selected
      if(!grepl(".docx", input$input_file$datapath)) {
        error_message = "Selected file is not .docx"
        output$error_message = renderText(error_message)
      } else {
        
        #Set date based on input
        if(input$date_button == "date_current") {
          date = Sys.Date()
          date = paste0(month.name[month(date)], " ", day(date), ", ", year(date))
        } else {
          date = input$date
          date = paste0(month.name[month(date)], " ", day(date), ", ", year(date))
        }
        
        company = input$company
        position = input$position
        hiring_manager = input$hiring_manager
        if(hiring_manager == "") {hiring_manager = "Hiring Manager"}
        
        #Print selected parameters
        if(date != "") {output$date_label = renderText(paste0("<b> Date: </b>", date))}
        if(hiring_manager != "") {output$hiring_manager = renderText(paste0("<b> Hiring Manager: </b>", hiring_manager))}
        if(company != "") {output$company_label = renderText(paste0("<b> Company: </b>", company))}
        if(position != "") {output$position_label = renderText(paste0("<b> Job Title: </b>", position))}
        
        
        doc <- read_docx(input$input_file$datapath)
        
        doc <- body_replace_all_text(doc, "\\[company\\]", company)
        doc <- body_replace_all_text(doc, "\\[date\\]", date)
        doc <- body_replace_all_text(doc, "\\[position\\]", position)
        doc <- body_replace_all_text(doc, "\\[hiring manager\\]", hiring_manager)
        
        
        company_file = tolower(company)
        company_file = gsub(" ", "_", company_file)
        print(doc, target = paste0(input$output_path, "/cover_letter_", company_file, ".docx"))
        
        
        
      }
    } else {
      error_message = "No File Selected"
      output$error_message = renderText(error_message)
    }
  })
  
}

shinyApp(ui = ui, server = server)     


