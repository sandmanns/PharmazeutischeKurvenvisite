library(shiny)
library(shinythemes)
library(DT)
library(stringi)
library(shinyWidgets)
library(jsonlite)

css <- "
#reverseSlider .irs-bar {
    border-top: 1px solid #ddd;
    border-bottom: 1px solid #ddd;
    background: linear-gradient(to bottom, #DDD -50%, #FFF 150%);
}
#reverseSlider .irs-bar-edge {
    border: 1px solid #ddd;
    background: linear-gradient(to bottom, #DDD -50%, #FFF 150%);
    border-right: 0;
}
#reverseSlider .irs-line {
    background: #428bca;
    border: 1px solid #428bca;
}
"

shinyUI(fluidPage(
  includeCSS("app.css"),
  theme = shinytheme("cerulean"),
  titlePanel(div("Dokumentationsdateien der pharmazeutischen Kurvenvisite",
                 img(height = 40, width = 60.6, src = "medlogo.jpeg",class="pull-right"),
                 img(height = 40, width = 20, src = "white.png",class="pull-right"),
                 img(height = 43, width = 71.3, src = "IMDS.png",class="pull-right")),
             windowTitle="Pharmazeutische Kurvenvisite"
  ),
  sidebarPanel(
    shinyjs::useShinyjs(),
    tabsetPanel(id="config",
                tabPanel("Input",
                         br(),
                         fileInput('inputFile',label = "Input",multiple = T,accept = "xlsx",
                                   buttonLabel = "Durchsuchen",
                                   placeholder = "Noch keine Datei ausgewählt"),
                         column(3,offset = 0,actionButton('do_in',"Daten einlesen",class = "btn-primary")),
                         column(3,offset = 2,actionButton('do_clear',"RESET",class = "btn-warning")),
                         br(),
                         br(),
                         br(),
                         hr(),
                         uiOutput("exportUI0"),
                         uiOutput("exportUI1"),
                         br(),
                         downloadButton('do_out',"Vollständigen Datensatz exportieren",class="btn-primary",disabled = T)
 
                ),
                tabPanel("Patienten pro Tag",
                         tags$style(type='text/css', css),
                         hr(),
                         h4(textOutput("text_analyse2a")),
                         h4(textOutput("text_analyse2a2")),
                         uiOutput("DAY_UI1"),
                         hr(),
                         uiOutput("DAY_UI2"),
                         uiOutput("DAY_UI1a"),
                         br(),
                         uiOutput("DAY_UI2t"),
                         uiOutput("DAY_UI2u"),
                         hr(),
                         uiOutput("DAY_UI3"),
                         uiOutput("DAY_UI3a"),
                         uiOutput("DAY_UI3b"),
                         uiOutput("DAY_UI3c"),
                         hr(),
                         uiOutput("DAY_UI4"),
                         uiOutput("DAY_UI4a"),
                         uiOutput("DAY_UI4b"),
                         br(),
                         actionButton('do_inDAY',"Patienten pro Tag analysieren",class = "btn-primary",disabled = T)
                ),
                tabPanel("Medikationsprüfung",
                         hr(),
                         h4(textOutput("text_analyse2a_med")),
                         h4(textOutput("text_analyse2a2_med")),
                         uiOutput("MED_UI1"),
                         hr(),
                         uiOutput("MED_UI2"),
                         uiOutput("MED_UI1a"),
                         br(),
                         uiOutput("MED_UI2t"),
                         uiOutput("MED_UI2u"),
                         hr(),
                         uiOutput("MED_UI3"),
                         br(),
                         actionButton('do_inMED',"Medikationsprüfung analysieren",class = "btn-primary",disabled = T)
                ),
                tabPanel("Export Analyse",
                         hr(),
                         h4(textOutput("text_export1")),
                         h4(textOutput("text_export2")),
                         h4(textOutput("text_export3")),
                         h5(textOutput("text_export4")),
                         hr(),
                         uiOutput("exportUI0b"),
                         uiOutput("exportUI1b"),
                         br(),
                         downloadButton('do_out2',"Datensatz + Analyse exportieren",class="btn-primary",disabled = T)
                         ),
                tabPanel("Info",
                         hr(),
                         h4(htmlOutput("text_info2h")),
                         hr(),
                         h4(textOutput("text_info1")),
                         h5(textOutput("text_info2b")),
                         h5(textOutput("text_info2a")),
                         hr(),
                         h4(textOutput("text_info1b")),
                         h5(textOutput("text_info2c")),
                         hr(),
                         br(),
                         br(),
                         h6(textOutput("text_info2g"))
                         )

    )
  ),
  mainPanel(
    shinyjs::useShinyjs(),
    tabsetPanel(id="main",
                tabPanel("Log",
                         h3("Log"),
                         div(id = "text")
                ),
                tabPanel("Input",
                         hr(),
                         h4(textOutput("text_analyse1")),
                         dataTableOutput("data_overview")
                ),
                tabPanel("Patienten pro Tag",
                         tabsetPanel(id="proTag",
                                     tabPanel("Tabulärer Output",
                                              h4(textOutput("text_analyse2b")),
                                              dataTableOutput("data_proTag"),
                                              hr(),
                                              dataTableOutput("egfr_proTag"),
                                              hr()
                                     ),
                                     tabPanel("Graphischer Output",
                                              h4(textOutput("text_analyse2c")),
                                              plotOutput("barplot_proTag",height = "700px"),
                                              br(),
                                              hr(),
                                              br(),
                                              plotOutput("barplot_proTag2",height = "400px"),
                                              br(),
                                              hr(),
                                              br(),
                                              plotOutput("barplot_proTag3",height = "600px"),
                                              br(),
                                              br()
                                              )
                                     )
                ),
                tabPanel("Medikationsprüfungen",
                         tabsetPanel(id="proMed",
                                     tabPanel("Tabulärer Output",
                                              h4(textOutput("text_analyse2b_med")),
                                              dataTableOutput("data_proMed"),
                                              hr(),
                                              dataTableOutput("data_fb"),
                                              hr(),
                                              dataTableOutput("data_proMed2"),
                                              hr(),
                                              dataTableOutput("data_proMed3"),
                                              hr()
                                     ),
                                     tabPanel("Graphischer Output",
                                              h4(textOutput("text_analyse2c_med")),
                                              plotOutput("barplot_proMed",height = "400px"),
                                              br(),
                                              hr(),
                                              br(),
                                              plotOutput("barplot_proMed1",height = "700px"),
                                              br(),
                                              hr(),
                                              br(),
                                              plotOutput("barplot_proMed2",height = "550px"),
                                              br(),
                                              hr(),
                                              br(),
                                              plotOutput("barplot_proMed3",height = "400px"),
                                              br(),
                                              br()
                                              )
                         )
                )
    )
  )
))
      
      
      
      
      
      
      
                         