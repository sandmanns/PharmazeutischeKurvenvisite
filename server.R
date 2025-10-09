options(shiny.maxRequestSize=30*1024^2)
library(shinyWidgets)
library(openxlsx)
library(memisc)
library(timetools)
library(stringr)
library(shinyFiles)
library(lubridate)
library(Hmisc)

own_pie<-function (x, labels = names(x), edges = 200, radius = 0.95, clockwise = FALSE, 
          init.angle = if (clockwise) 90 else 0, density = NULL, angle = 45, 
          col = NULL, border = NULL, lty = NULL, main = NULL) 
{
  if (!is.numeric(x) || any(is.na(x) | x < 0)) 
    stop("'x' values must be positive.")
  if (is.null(labels)) 
    labels <- as.character(seq_along(x))
  else labels <- as.graphicsAnnot(labels)
  x <- c(0, cumsum(x)/sum(x))
  dx <- diff(x)
  nx <- length(dx)
  plot.new()
  pin <- par("pin")
  xlim <- ylim <- c(-1,1)
  if (pin[1L] > pin[2L]) 
    xlim <- (pin[1L]/pin[2L]) * xlim
  else ylim <- (pin[2L]/pin[1L]) * ylim
  dev.hold()
  on.exit(dev.flush())
  plot.window(xlim, ylim, "", asp = 1)
  if (is.null(col)) 
    col <- if (is.null(density)) 
      c("white", "lightblue", "mistyrose", "lightcyan", 
        "lavender", "cornsilk")
  else par("fg")
  if (!is.null(col)) 
    col <- rep_len(col, nx)
  if (!is.null(border)) 
    border <- rep_len(border, nx)
  if (!is.null(lty)) 
    lty <- rep_len(lty, nx)
  angle <- rep(angle, nx)
  if (!is.null(density)) 
    density <- rep_len(density, nx)
  twopi <- if (clockwise) 
    -2 * pi
  else 2 * pi
  t2xy <- function(t) {
    t2p <- twopi * t + init.angle * pi/180
    list(x = radius * cos(t2p), y = radius * sin(t2p))
  }
  for (i in 1L:nx) {
    n <- max(2, floor(edges * dx[i]))
    P <- t2xy(seq.int(x[i], x[i + 1], length.out = n))
    polygon(c(P$x, 0), c(P$y, 0), density = density[i], angle = angle[i], 
            border = border[i], col = col[i], lty = lty[i])
    P <- t2xy(mean(x[i + 0:1]))
    lab <- as.character(labels[i])
    if (!is.na(lab) && nzchar(lab)) {
      lines(c(1, 1.05) * P$x, c(1, 1.05) * P$y)
      text(1.1 * P$x, 1.1 * P$y, labels[i], xpd = TRUE, 
           adj = ifelse(P$x < 0, 1, 0),cex=2)
    }
  }
  title(main = main,cex.main=3)
}

capitalize<-function(x){
  x2<-gsub("."," ",x,fixed=T)
  x3<-strsplit(x2,"")[[1]]
  x4<-c()
  for(i in 1:length(x3)){
    if(is.na(suppressWarnings(as.numeric(x3[i])))&&(i==1||(x3[i-1])==" ")){
      x4<-c(x4,toupper(x3[i]))
    }else{
      x4<-c(x4,x3[i])
    }
  }
  return(paste(x4,collapse=""))
}

file.opened <- function(path) {
  suppressWarnings(
    "try-error" %in% class(
      try(file(path, 
               open = "w"), 
          silent = TRUE
      )
    )
  )
}

shinyServer(function(input, output, session) {
  observeEvent(input$do_clear,{
    session$reload()
  })

  output$text_analyse1<-renderText({"Noch keine Daten verfügbar."})
  output$text_analyse2a<-renderText({"Upload noch nicht durchgeführt."})
  output$text_analyse2a2<-renderText({"Bitte lesen Sie zunächst einen Datensatz ein."})
  output$text_analyse2b<-renderText({"Noch keine Daten verfügbar."})
  output$text_analyse2c<-renderText({"Noch keine Daten verfügbar."})
  
  output$text_analyse2a_med<-renderText({"Upload noch nicht durchgeführt."})
  output$text_analyse2a2_med<-renderText({"Bitte lesen Sie zunächst einen Datensatz ein."})
  output$text_analyse2b_med<-renderText({"Noch keine Daten verfügbar."})
  output$text_analyse2c_med<-renderText({"Noch keine Daten verfügbar."})
  
  output$text_export1<-renderText({"Upload noch nicht durchgeführt."})
  output$text_export2<-renderText({"Bitte lesen Sie zunächst einen Datensatz ein."})
  output$text_export3<-renderText({NULL})
  output$text_export4<-renderText({NULL})
  
  output$data_overview<-renderDataTable({NULL})
  output$data_proTag<-renderDataTable({NULL})
  output$data_proMed<-renderDataTable({NULL})
  output$data_proMed2<-renderDataTable({NULL})
  output$data_proMed3<-renderDataTable({NULL})
  output$data_fb<-renderDataTable({NULL})
  output$barplot_proTag<-renderPlot({NULL})
  output$barplot_proTag2<-renderPlot({NULL})
  output$barplot_proTag3<-renderPlot({NULL})
  output$barplot_proMed<-renderPlot({NULL})
  output$barplot_proMed1<-renderPlot({NULL})
  output$barplot_proMed2<-renderPlot({NULL})
  output$barplot_proMed3<-renderPlot({NULL})
  output$DAY_UI1<-renderUI({NULL})
  output$DAY_UI1a<-renderUI({NULL})
  output$DAY_UI2<-renderUI({NULL})
  output$DAY_UI2t<-renderUI({NULL})
  output$DAY_UI2u<-renderUI({NULL})
  output$DAY_UI3<-renderUI({NULL})
  output$DAY_UI3a<-renderUI({NULL})
  output$DAY_UI3b<-renderUI({NULL})
  output$DAY_UI3c<-renderUI({NULL})
  output$DAY_UI4<-renderUI({NULL})
  output$DAY_UI4a<-renderUI({NULL})
  output$DAY_UI4b<-renderUI({NULL})
  output$MED_UI1<-renderUI({NULL})
  output$MED_UI1a<-renderUI({NULL})
  output$MED_UI2<-renderUI({NULL})
  output$MED_UI2t<-renderUI({NULL})
  output$MED_UI2u<-renderUI({NULL})
  output$MED_UI3<-renderUI({NULL})
  output$exportUI0<-renderUI({NULL})
  output$exportUI0b<-renderUI({NULL})
  output$exportUI1<-renderUI({NULL})
  output$exportUI1b<-renderUI({NULL})
  
  values<-reactiveValues()
  values$day_type<-T
  values$day_kliniken<-NA
  values$med_type<-T
  values$med_kliniken<-NA
  values$DAY_Start<-NA
  values$DAY_End<-NA
  values$MED_Start<-NA
  values$MED_End<-NA
  
  output$text_info2h<-renderText({"Die Zweckbestimmung dieser Anwendung ist eine Auswertungsübersicht über die pharmazeutische Kurvenvisite.
Die Anwendung ist kein Medizinprodukt nach Medizinproduktegesetz oder EU-Medical Device Regulation."})
  output$text_info1<-renderText({"Entwickelt von"})
  output$text_info2b<-renderText({"Institut für Medical Data Science"})
  output$text_info2a<-renderText({"PD Dr. Sarah Sandmann"})
  output$text_info1b<-renderText({"Kontakt:"})
  output$text_info2c<-renderText({"sarah.sandmann@med.ovgu.de"})
  output$text_info2g<-renderText({"Letzte Änderung: 06.09.2025"})
  
  shinyjs::disable("do_out")
  
  
  observeEvent(input$do_in,{
    updateTabsetPanel(session,"main",
                      selected="Log")
    output$text_analyse1<-renderText({"Noch keine Daten verfügbar."})
    output$text_analyse2a<-renderText({"Upload noch nicht durchgeführt."})
    output$text_analyse2a2<-renderText({"Bitte lesen Sie zunächst einen Datensatz ein."})
    output$text_analyse2b<-renderText({"Noch keine Daten verfügbar."})
    output$text_analyse2c<-renderText({"Noch keine Daten verfügbar."})
    
    output$text_analyse2a_med<-renderText({"Upload noch nicht durchgeführt."})
    output$text_analyse2a2_med<-renderText({"Bitte lesen Sie zunächst einen Datensatz ein."})
    output$text_analyse2b_med<-renderText({"Noch keine Daten verfügbar."})
    output$text_analyse2c_med<-renderText({"Noch keine Daten verfügbar."})
    
    output$text_export1<-renderText({"Upload noch nicht durchgeführt."})
    output$text_export2<-renderText({"Bitte lesen Sie zunächst einen Datensatz ein."})
    output$text_export3<-renderText({NULL})
    output$text_export4<-renderText({NULL})
    
    output$data_overview<-renderDataTable({NULL})
    output$data_proTag<-renderDataTable({NULL})
    output$egfr_proTag<-renderDataTable({NULL})
    output$data_proMed<-renderDataTable({NULL})
    output$data_proMed2<-renderDataTable({NULL})
    output$data_proMed3<-renderDataTable({NULL})
    output$data_fb<-renderDataTable({NULL})
    output$barplot_proTag<-renderPlot({NULL})
    output$barplot_proTag2<-renderPlot({NULL})
    output$barplot_proMed<-renderPlot({NULL})
    output$barplot_proMed1<-renderPlot({NULL})
    output$barplot_proMed2<-renderPlot({NULL})
    output$barplot_proMed3<-renderPlot({NULL})
    output$DAY_UI1<-renderUI({NULL})
    output$DAY_UI1a<-renderUI({NULL})
    output$DAY_UI2<-renderUI({NULL})
    output$DAY_UI3<-renderUI({NULL})
    output$DAY_UI3a<-renderUI({NULL})
    output$DAY_UI3b<-renderUI({NULL})
    output$DAY_UI3c<-renderUI({NULL})
    output$DAY_UI4<-renderUI({NULL})
    output$DAY_UI4a<-renderUI({NULL})
    output$DAY_UI4b<-renderUI({NULL})
    output$MED_UI1<-renderUI({NULL})
    output$MED_UI1a<-renderUI({NULL})
    output$MED_UI2<-renderUI({NULL})
    output$MED_UI3<-renderUI({NULL})
    output$exportUI0<-renderUI({NULL})
    output$exportUI0b<-renderUI({NULL})
    output$exportUI1<-renderUI({NULL})
    output$exportUI1b<-renderUI({NULL})
    shinyjs::disable("do_out")
    
    input_temp<-input$inputFile
    
    if(is.null(input_temp)){
      shinyjs::html("text", paste0("FEHLER: Keine Input Datei definiert.","<br>"), add = TRUE) 
      return()
    }
    
    input2<-list()
    for(i in 1:length(input$inputFile[,1])){
      
      if(length(input$inputFile[,1])==1){
        shinyjs::html("text", paste0("<br>Input Datei wird eingelesen.<br><br><br>"), add = FALSE)
      }else{
        shinyjs::html("text", paste0("<br>Input Dateien werden eingelesen.<br><br><br>"), add = FALSE)
      }

      if(length(grep(".xlsx",input_temp$name[i],fixed=T))==0){
        shinyjs::html("text", paste0("FEHLER: Datei ",input_temp$name[i]," ist keine .xlsx-Datei"), add = TRUE)
        return()
      }
      
      if(sum(getSheetNames(input_temp$datapath[i])=="Datensatz")==0){
        shinyjs::html("text", paste0("FEHLER: Datei ",input_temp$name[i]," enthält kein Sheet 'Datensatz'.","<br>"), add = TRUE) 
        return()
      }
      
      input2[[i]] <- read.xlsx(input_temp$datapath[i],sheet="Datensatz")
    }
    output$text_analyse1<-renderText({NULL})
    if(length(input$inputFile[,1])==1){
      shinyjs::html("text", paste0("Datei erfolgreich eingelesen.<br><br>"), add = TRUE)
      shinyjs::html("text", paste0("&nbsp;&nbsp;&nbsp;&nbsp;",input_temp$name[i],"<br>"), add = TRUE)
    }else{
      shinyjs::html("text", paste0("Dateien erfolgreich eingelesen.<br><br>"), add = TRUE)
      for(i in 1:length(input$inputFile[,1])){
        shinyjs::html("text", paste0("&nbsp;&nbsp;&nbsp;&nbsp;",input_temp$name[i],"<br>"), add = TRUE)
      }
    }
    shinyjs::html("text", paste0("<br><br>"), add = TRUE)
    
    for(i in 1:length(input2)){
      if(length(grep("^eGFR",colnames(input2[[i]])))==0){
        shinyjs::html("text", paste0("Information zur eGFR fehlt in ",input_temp$name[i],
                                     ".<br>Bitte korrigieren Sie Ihre Input-Datei.<br>"), add = TRUE)
        return()
      }
      colnames(input2[[i]])[grep("^eGFR",colnames(input2[[i]]))]<-"eGFR"
      colnames(input2[[i]])[substring(colnames(input2[[i]]),first = rep(1,ncol(input2[[i]])),last = rep(2,ncol(input2[[i]])))=="K."]<-"Kalium"
      colnames(input2[[i]])[substring(colnames(input2[[i]]),first = rep(1,ncol(input2[[i]])),last = rep(3,ncol(input2[[i]])))=="Na."]<-"Natrium"
      colnames(input2[[i]])[grep("&#10;",colnames(input2[[i]]),fixed=T)]<-gsub("&#10;","",colnames(input2[[i]])[grep("&#10;",colnames(input2[[i]]),fixed=T)])
      colnames(input2[[i]])[colnames(input2[[i]])=="Triple.Whammy"]<-"Triple.Whammy.=.potentielle.Nierenschädigung"
      colnames(input2[[i]])[colnames(input2[[i]])=="Präventiv-maßnahmen"]<-"Präventivmaßnahmen"
      colnames(input2[[i]])[colnames(input2[[i]])=="Child-Pugh.Score.A-C"]<-"Child-Pugh.Score.C"
      input2[[i]]<-input2[[i]][grep("^X",colnames(input2[[i]]),invert=T)]
      
      
      if(length(grep("Kalium",colnames(input2[[i]])))==0){
        input2[[i]]<-cbind(input2[[i]],Kalium=NA)
      }
      if(length(grep("Natrium",colnames(input2[[i]])))==0){
        input2[[i]]<-cbind(input2[[i]],Natrium=NA)
      }
    }
    
    spalten<-tolower(colnames(input2[[1]]))
    if(length(input2)>1){
      for(i in 2:length(input2)){
        spalten<-c(spalten,tolower(colnames(input2[[i]])))
      }
    }
    spalten<-unique(spalten)
    
    for(i in 1:length(input2)){
      temp<-input2[[i]]
      colnames(temp)<-tolower(colnames(temp))
      temp2<-data.frame(Test=NA)
      for(j in 1:length(spalten)){
        if(sum(spalten[j]==colnames(temp))==1){
          temp2<-cbind(temp2,temp[,spalten[j]])
        }else{
          temp2<-cbind(temp2,NA)
        }
      }
      temp2<-temp2[,-1]
      colnames(temp2)<-spalten
      if(i==1){
        input3<-temp2
      }else{
        input3<-rbind(input3,temp2)
      }
    }
    
    input3$dialyse<-NA
    input3$dialyse[grep("Dialyse",input3$egfr)]<-"Dialyse"
    input3$egfr<-gsub("Dialyse","",input3$egfr)
    input3$egfr<-gsub(" ","",input3$egfr)
    input3$egfr[input3$egfr==""]<-NA
    input3$egfr[grep(".",input3$egfr,fixed=T)]<-round(as.numeric(input3$egfr[grep(".",input3$egfr,fixed=T)]),1)
    input3$egfr<-gsub(",",".",input3$egfr)
    
    if(length(grep("aufnahmezeitpunkt",colnames(input3)))==0){
      shinyjs::html("text", paste0("Spalte 'Aufnahmezeitpunkt' kann nicht gefunden werden.",
                                   "<br>Bitte korrigieren Sie Ihren Input.<br>"), add = TRUE)
      return()
    }
    if(length(grep("prüfdatum",colnames(input3)))==0){
      shinyjs::html("text", paste0("Notwendige Spalte 'Prüfdatum' kann nicht gefunden werden.",
                                   "<br>Bitte korrigieren Sie Ihren Input.<br>"), add = TRUE)
      return()
    }
    if(length(grep("alter",colnames(input3)))==0){
      shinyjs::html("text", paste0("Notwendige Spalte 'Alter' kann nicht gefunden werden.",
                                   "<br>Bitte korrigieren Sie Ihren Input.<br>"), add = TRUE)
      return()
    }
    if(length(grep("medikationsprüfungen.pharmazeutische.visite",colnames(input3),fixed=T))==0){
      shinyjs::html("text", paste0("Notwendige Spalte 'Medikationsprüfungen pharmazeutische Visite' kann nicht gefunden werden.",
                                   "<br>Bitte korrigieren Sie Ihren Input.<br>"), add = TRUE)
      return()
    }
    if(length(grep("antibiotika.x.tage.(abs.visite)",colnames(input3),fixed=T))==0){
      shinyjs::html("text", paste0("Notwendige Spalte 'Antibiotika x Tage (ABS Visite)' kann nicht gefunden werden.",
                                   "<br>Bitte korrigieren Sie Ihren Input.<br>"), add = TRUE)
      return()
    }
    if(length(grep("abp.betreffend",colnames(input3),fixed=T))==0){
      shinyjs::html("text", paste0("Notwendige Spalte 'ABP betreffend' kann nicht gefunden werden.",
                                   "<br>Bitte korrigieren Sie Ihren Input.<br>"), add = TRUE)
      return()
    }
    if(length(grep("fehlerfolge.(ncc.merp)",colnames(input3),fixed=T))==0){
      shinyjs::html("text", paste0("Notwendige Spalte 'Fehlerfolge (NCC MERP)' kann nicht gefunden werden.",
                                   "<br>Bitte korrigieren Sie Ihren Input.<br>"), add = TRUE)
      return()
    }

    
    input4<-input3[,colnames(input3)%nin%c("zimmer","name","abrtyp","kurze.beschreibung","arzt","abr.art")]
    if(sum(!is.na(suppressWarnings(as.numeric(input4$aufnahmezeitpunkt))))>0){
      input4[,"aufnahmezeitpunkt"]<-suppressWarnings(convertToDateTime(input4[,"aufnahmezeitpunkt"]))
    }else{
      input4[,"aufnahmezeitpunkt"]<-suppressWarnings(dmy_hm(input4[,"aufnahmezeitpunkt"]))
    }
      
    if(sum(colnames(input3)=="entlasszeitpunkt")==1&&sum(is.na(input3$entlasszeitpunkt))!=nrow(input3)){
      if(sum(!is.na(suppressWarnings(as.numeric(input4$entlasszeitpunkt))))>0){
        input4[,"entlasszeitpunkt"]<-suppressWarnings(convertToDateTime(input4[,"entlasszeitpunkt"]))
      }else{
        input4[,"entlasszeitpunkt"]<-suppressWarnings(dmy_hm(input4[,"entlasszeitpunkt"]))
      }
    }
    
    if(sum(!is.na(suppressWarnings(as.numeric(input4$prüfdatum))))>0){
      input4[,"prüfdatum"]<-suppressWarnings(convertToDate(input4[,"prüfdatum"]))
    }else{
      input4[,"prüfdatum"]<-suppressWarnings(dmy(input4[,"prüfdatum"]))
    }
    
    #Spalten ausschließen, für die Informationen in allen Fällen fehlen
    input5<-input4[!is.na(input4$fallnummer),]
    input5<-input5[!is.na(suppressWarnings(as.numeric(input5$fallnummer))),colSums(is.na(input5))!=nrow(input5)|colnames(input5)=="natrium"|colnames(input5)=="kalium"|colnames(input5)=="antibiotika.x.tage.(abs.visite)"|colnames(input5)=="egfr"|colnames(input5)=="dialyse"]
    input5$alter_original<-input5$alter
    
    for(i in 1:nrow(input5)){
      if(length(grep("Jahre",input5$alter[i]))>0){
        input5$alter[i]<-gsub(" Jahre","",input5$alter[i])
      }else{
        if(length(grep("Monate",input5$alter[i]))>0){
          input5$alter[i]<-floor(as.numeric(gsub(" Monate","",input5$alter[i]))/12)
        }else{
          if(length(grep("Wochen",input5$alter[i]))>0){
            input5$alter[i]<-floor(as.numeric(gsub(" Wochen","",input5$alter[i]))/52)
          }else{
            if(length(grep("Tage",input5$alter[i]))>0){
              input5$alter[i]<-floor(as.numeric(gsub(" Tage","",input5$alter[i]))/365.25)
            }else{
              input5$alter[i]<-floor(as.numeric((input5$prüfdatum[i]-suppressWarnings(convertToDate(input5$alter[i])))/365.25))
            }
          }
        }
      }
    }
    input5b<-input5
    input5b$aufnahmezeitpunkt<-as.character(input5b$aufnahmezeitpunkt)
    if(sum(colnames(input5b)=="entlasszeitpunkt")==1){
      input5b[,"entlasszeitpunkt"]<-as.character(input5b[,"entlasszeitpunkt"])
    }
    input5b<-input5b[order(input5b$prüfdatum),]
    
    input5b_export<-input5b
    names(input5b_export)[names(input5b_export)=="fachrichtung"]<-"Fachrichtung"
    names(input5b_export)[names(input5b_export)=="fallnummer"]<-"Fallnummer"
    input5b_export$alter<-input5b_export$alter_original
    names(input5b_export)[names(input5b_export)=="alter"]<-"Alter"
    
    input5b_export$aufnahmezeitpunkt<-paste0(str_split_fixed(str_split_fixed(input5b_export$aufnahmezeitpunkt,"-",n=Inf)[,3]," ",n=Inf)[,1],".",
                                             str_split_fixed(input5b_export$aufnahmezeitpunkt,"-",n=Inf)[,2],".",
                                             str_split_fixed(input5b_export$aufnahmezeitpunkt,"-",n=Inf)[,1]," ",
                                             str_split_fixed(str_split_fixed(input5b_export$aufnahmezeitpunkt," ",n=Inf)[,2],":",n=Inf)[,1],":",
                                             str_split_fixed(str_split_fixed(input5b_export$aufnahmezeitpunkt," ",n=Inf)[,2],":",n=Inf)[,2])
    input5b_export$aufnahmezeitpunkt<-gsub(" :$","",input5b_export$aufnahmezeitpunkt)
    names(input5b_export)[names(input5b_export)=="aufnahmezeitpunkt"]<-"Aufnahmezeitpunkt"
    
    input5b_export$prüfdatum<-paste0(str_split_fixed(input5b_export$prüfdatum,"-",n=Inf)[,3],".",
                                             str_split_fixed(input5b_export$prüfdatum,"-",n=Inf)[,2],".",
                                             str_split_fixed(input5b_export$prüfdatum,"-",n=Inf)[,1])
    names(input5b_export)[names(input5b_export)=="prüfdatum"]<-"Prüfdatum"
    
    if(sum(names(input5b_export)=="entlasszeitpunkt")>0){
      input5b_export$entlasszeitpunkt<-paste0(str_split_fixed(str_split_fixed(input5b_export$entlasszeitpunkt,"-",n=Inf)[,3]," ",n=Inf)[,1],".",
                                               str_split_fixed(input5b_export$entlasszeitpunkt,"-",n=Inf)[,2],".",
                                               str_split_fixed(input5b_export$entlasszeitpunkt,"-",n=Inf)[,1]," ",
                                               str_split_fixed(str_split_fixed(input5b_export$entlasszeitpunkt," ",n=Inf)[,2],":",n=Inf)[,1],":",
                                               str_split_fixed(str_split_fixed(input5b_export$entlasszeitpunkt," ",n=Inf)[,2],":",n=Inf)[,2])
      input5b_export$entlasszeitpunkt<-gsub(" :$","",input5b_export$entlasszeitpunkt)
      input5b_export$entlasszeitpunkt<-gsub("..NA","",input5b_export$entlasszeitpunkt,fixed=T)
      input5b_export$entlasszeitpunkt[input5b_export$entlasszeitpunkt==""]<-NA
      names(input5b_export)[names(input5b_export)=="entlasszeitpunkt"]<-"Entlasszeitpunkt"
    }
    
    input5b_export$egfr<-paste(input5b_export$egfr,input5b_export$dialyse,sep=" ")
    input5b_export$egfr<-gsub("NA NA",NA,input5b_export$egfr)
    input5b_export$egfr<-gsub(" NA","",input5b_export$egfr)
    input5b_export$egfr<-gsub("NA ","",input5b_export$egfr)
    input5b_export$egfr<-gsub(".",",",input5b_export$egfr,fixed=T)
    names(input5b_export)[names(input5b_export)=="egfr"]<-"eGFR außerhalb Normbereich"
    
    input5b_export$kalium<-str_split_fixed(input5b_export$kalium," ",Inf)[,1]
    input5b_export$kalium[grep(".",input5b_export$kalium,fixed=T)]<-round(as.numeric(input5b_export$kalium[grep(".",input5b_export$kalium,fixed=T)]),1)
    input5b_export$kalium<-gsub(".",",",input5b_export$kalium,fixed=T)
    names(input5b_export)[names(input5b_export)=="kalium"]<-"K außerhalb Normbereich"

    input5b_export$natrium<-str_split_fixed(input5b_export$natrium," ",Inf)[,1]
    input5b_export$natrium[grep(".",input5b_export$natrium,fixed=T)]<-round(as.numeric(input5b_export$natrium[grep(".",input5b_export$natrium,fixed=T)]),1)
    input5b_export$natrium<-gsub(".",",",input5b_export$natrium,fixed=T)
    names(input5b_export)[names(input5b_export)=="natrium"]<-"Na außerhalb Normbereich"
    names(input5b_export)[names(input5b_export)=="antibiotika.x.tage.(abs.visite)"]<-"Antibiotika x Tage\n(ABS Visite)"
    names(input5b_export)[names(input5b_export)=="medikationsprüfungen.pharmazeutische.visite"]<-"Medikationsprüfungen\npharmazeutische Visite"
    names(input5b_export)[names(input5b_export)=="abp.betreffend"]<-"ABP betreffend"
    names(input5b_export)[names(input5b_export)=="fehlerfolge.(ncc.merp)"]<-"Fehlerfolge (NCC MERP)"
    names(input5b_export)[names(input5b_export)=="korrigierende.maßnahmen"]<-"Korrigierende Maßnahmen"
    input5b_export<-input5b_export[,which(names(input5b_export)!="dialyse"&names(input5b_export)!="alter_original")]
    names(input5b_export)[names(input5b_export)=="outcome.problem"]<-"Outcome Problem"
    names(input5b_export)[names(input5b_export)=="präventivmaßnahmen"]<-"Präventivmaßnahmen"
    names(input5b_export)[names(input5b_export)=="triple.whammy.=.potentielle.nierenschädigung"]<-"Triple Whammy =\npotentielle Nierenschädigung"
 
    names(input5b_export)[names(input5b_export)=="antibiotika.bei.egfr<30ml/min"]<-"Antibiotika bei\neGFR<30ml/min"
    names(input5b_export)[names(input5b_export)=="fehlerart"]<-"Fehlerart"
    names(input5b_export)[names(input5b_export)=="status.des.abp"]<-"Status des ABP"
    names(input5b_export)[names(input5b_export)=="child-pugh.score.c"]<-"Child-Pugh Score C"
    names(input5b_export)[names(input5b_export)=="fehlerart.medikation.(angelehnt.an.pcne)"]<-"Fehlerart Medikation\n(angelehnt an PCNE)"
    
    for(name in 1:ncol(input5b_export)){
      if(names(input5b_export)[name]%nin%c("Fachrichtung","Fallnummer","Alter","Aufnahmezeitpunkt","Entlasszeitpunkt","Prüfdatum",
                                           "eGFR außerhalb Normbereich","K außerhalb Normbereich","Na außerhalb Normbereich",
                                           "Antibiotika x Tage\n(ABS Visite)","Medikationsprüfungen\npharmazeutische Visite",
                                           "ABP betreffend","Fehlerfolge (NCC MERP)","Korrigierende Maßnahmen",
                                           "Outcome Problem","Präventivmaßnahmen","Triple Whammy =\npotentielle Nierenschädigung",
                                           "Antibiotika bei\neGFR<30ml/min","Fehlerart","Status des ABP","Child-Pugh Score C",
                                           "Fehlerart Medikation\n(angelehnt an PCNE)")){
        names(input5b_export)[name]<-capitalize(names(input5b_export)[name])
    
      }
    }
    
    
    
    output$data_overview<-renderDataTable(datatable(input5b_export,rownames=F,
                                                   options = list(
                                                     dom = 'Blfrtip',pageLength=25,
                                                     language=list(search='Suchen:',lengthMenu='Zeige _MENU_ Einträge',
                                                                   info='Zeige _START_ bis _END_ von _TOTAL_ Einträgen',
                                                                   paginate=list(previous='Vorherige','next'='Nächste')),
                                                     columnDefs = list(list(className = 'dt-right', targets = c(1:(ncol(input5b_export)-1)))))),
                                         server=T
    )
    
    

    shinyjs::enable("do_out")
    shinyjs::enable("do_inDAY")
    shinyjs::enable("do_inMED")
    
   # shinyFiles::shinyDirChoose(input=input,session = session,id = 'export_dir',roots=c(home="~"))
    shinyFiles::shinyDirChoose(input=input,session = session,id = 'export_dir',roots=c(home="C:/"))
    output$exportUI1<-renderUI({shinyDirButton('export_dir',title = "Export-Ordner",
                                               label = "Durchsuchen")})
    observe(
      if(!is.null(input$export_dir)){
        if(paste(unlist(input$export_dir[[1]]),collapse="/")==0||paste(unlist(input$export_dir[[1]]),collapse="/")==1){
          #output$exportUI0<-renderUI({h5(paste0("Ordner für den Export: ~/"))})
          output$exportUI0<-renderUI({h5(paste0("Ordner für den Export: C:/"))})
        }else{
         # output$exportUI0<-renderUI({h5(paste0("Ordner für den Export: ~",
        #                                        paste(unlist(input$export_dir[[1]]),collapse="/")))})
          output$exportUI0<-renderUI({h5(paste0("Ordner für den Export: C:/",
                                                paste(unlist(input$export_dir[[1]]),collapse="/")))})
        }
      }else{
        #output$exportUI0<-renderUI({h5(paste0("Ordner für den Export: ~/"))})
        output$exportUI0<-renderUI({h5(paste0("Ordner für den Export: C:/"))})
      }      
    )

    
    output$text_export1<-renderText({"Analysen noch nicht durchgeführt."})
    output$text_export2<-renderText({"Bitte führen Sie die Analysen 'Patienten pro Tag' und 'Medikationsprüfung' durch."})
    
    
    
    ########Analyse pro Tag
    output$text_analyse2a<-renderText({NULL})
    output$text_analyse2a2<-renderText({NULL})
    
    start<-as.Date(min(input5$prüfdatum),format="yyyy-mm-dd")
    helper<-as.numeric(strsplit(as.character(start),"-",fixed=T)[[1]][3])-1
    start<-start-helper
    
    monat<-as.numeric(strsplit(as.character(start),"-",fixed=T)[[1]][2])
    if(monat==1||monat==3||monat==5||monat==7||monat==8||monat==10||monat==12){
      monat_tage<-31
    }else{
      if(monat==4||monat==6||monat==9||monat==11){
        monat_tage<-30
      }else{
        if((as.numeric(strsplit(as.character(start),"-",fixed=T)[[1]][1])%%4)==0&&(as.numeric(strsplit(as.character(start),"-",fixed=T)[[1]][1])%%1000)!=0){
          monat_tage<-29
        }else{
          monat_tage<-28
        }
      }
    }
    ende_default<-start+monat_tage-1
    
    ende<-as.Date(max(input5$prüfdatum),format="yyyy-mm-dd")
    monat2<-as.numeric(strsplit(as.character(ende),"-",fixed=T)[[1]][2])
    if(monat2==1||monat2==3||monat2==5||monat2==7||monat2==8||monat2==10||monat2==12){
      monat_tage2<-31
    }else{
      if(monat2==4||monat2==6||monat2==9||monat2==11){
        monat_tage2<-30
      }else{
        if((as.numeric(strsplit(as.character(ende),"-",fixed=T)[[1]][1])%%4)==0&&(as.numeric(strsplit(as.character(ende),"-",fixed=T)[[1]][1])%%1000)!=0){
          monat_tage2<-29
        }else{
          monat_tage2<-28
        }
      }
    }
    helper<-as.numeric(strsplit(as.character(ende),"-",fixed=T)[[1]][3])
    ende_max<-ende-helper+monat_tage2
    
    start_anders<-paste(strsplit(as.character(start),"-",fixed = T)[[1]][c(3,2,1)],collapse = ".")
    ende_max_anders<-paste(strsplit(as.character(ende_max),"-",fixed = T)[[1]][c(3,2,1)],collapse = ".")
    
    output$DAY_UI1<-renderUI({dateInput("DAY_Start","Beginn des Analysezeitraums",
                                        value=start,
                                        min = start,
                                        max = ende_max,
                                        weekstart = 1,language = "de",format="yyyy-mm-dd")})
    
    #output$DAY_UI1<-renderUI({airDatepickerInput('DAY_Start',"Analysezeitraum",range=T,value = c(start,ende_max),
    #                                             minDate = start,maxDate = ende_max,language = "de",autoClose = T)})
    
    output$DAY_UI1a<-renderUI({h6(paste0("Aufgrund der eingelesenen Daten können nur Tage zwischen dem ",
                                           start_anders,
                                           " und ",ende_max_anders," ausgewählt werden.\n",
                                           "Standard ist ein Zeitraum von 1 Monat. Dies kann allerdings nach Bedarf geändert werden (Maximum ",
                                           as.numeric(as.Date(ende_max)-as.Date(start)),
                                           " Tage). Bitte beachten Sie, dass die Analyse im Fall von sehr langen Zeiträumen unübersichtlich werden kann."))})

    observe({if(!is.null(input$DAY_Start)&&is.null(input$DAY_End)){
      output$DAY_UI1a<-renderUI({h6(paste0("Aufgrund der eingelesenen Daten können nur Tage zwischen dem ",
                                           start_anders,
                                           " und ",ende_max_anders," ausgewählt werden.\n",
                                           "Standard ist ein Zeitraum von 1 Monat. Dies kann allerdings nach Bedarf geändert werden (Maximum ",
                                           as.numeric(as.Date(ende_max)-as.Date(start)),
                                           " Tage). Bitte beachten Sie, dass die Analyse im Fall von sehr langen Zeiträumen unübersichtlich werden kann."))})
      
        if(ende_default<=ende_max){
          #output$DAY_UI1<-renderUI({airDatepickerInput('DAY_Start',"Analysezeitraum",range=T,value = c(start,ende_default),
          #                                             minDate = start,maxDate = ende_max,language = "de",autoClose = T)})
          output$DAY_UI2<-renderUI({dateInput("DAY_End","Ende des Analysezeitraums",
                                              value=ende_default,
                                              min = start,
                                              max = ende_max,
                                              weekstart = 1,language = "de",format="yyyy-mm-dd")})                            
        }else{
          #output$DAY_UI1<-renderUI({airDatepickerInput('DAY_Start',"Analysezeitraum",range=T,value = c(start,ende_max),
          #                                             minDate = start,maxDate = ende_max,language = "de",autoClose = T)})
          output$DAY_UI2<-renderUI({dateInput("DAY_End","Ende des Analysezeitraums",
                                              value=ende_max,
                                              min = start,
                                              max = ende_max,
                                              weekstart = 1,language = "de",format="yyyy-mm-dd")})  
        }
        
      }
    })
    
    observe({if(!is.null(input$DAY_End)&&input$DAY_Start>input$DAY_End){
      start2<-as.Date(input$DAY_Start,format="yyyy-mm-dd")
      helper<-as.numeric(strsplit(as.character(start2),"-",fixed=T)[[1]][3])-1
      start2<-start2-helper
      
      monat<-as.numeric(strsplit(as.character(start2),"-",fixed=T)[[1]][2])
      if(monat==1||monat==3||monat==5||monat==7||monat==8||monat==10||monat==12){
        monat_tage<-31
      }else{
        if(monat==4||monat==6||monat==9||monat==11){
          monat_tage<-30
        }else{
          if((as.numeric(strsplit(as.character(start2),"-",fixed=T)[[1]][1])%%4)==0&&(as.numeric(strsplit(as.character(start2),"-",fixed=T)[[1]][1])%%1000)!=0){
            monat_tage<-29
          }else{
            monat_tage<-28
          }
        }
      }
      start2<-input$DAY_Start
      ende_default<-start2+monat_tage-1
      
      ende<-as.Date(max(input5$prüfdatum),format="yyyy-mm-dd")
      monat2<-as.numeric(strsplit(as.character(ende),"-",fixed=T)[[1]][2])
      if(monat2==1||monat2==3||monat2==5||monat2==7||monat2==8||monat2==10||monat2==12){
        monat_tage2<-31
      }else{
        if(monat2==4||monat2==6||monat2==9||monat2==11){
          monat_tage2<-30
        }else{
          if((as.numeric(strsplit(as.character(ende),"-",fixed=T)[[1]][1])%%4)==0&&(as.numeric(strsplit(as.character(ende),"-",fixed=T)[[1]][1])%%1000)!=0){
            monat_tage2<-29
          }else{
            monat_tage2<-28
          }
        }
      }
      helper<-as.numeric(strsplit(as.character(ende),"-",fixed=T)[[1]][3])
      ende_max<-ende-helper+monat_tage2
      
      ende_max_anders<-paste(strsplit(as.character(ende_max),"-",fixed = T)[[1]][c(3,2,1)],collapse = ".")
      
      if(ende_default<=ende_max){
        #output$DAY_UI1<-renderUI({airDatepickerInput('DAY_Start',"Analysezeitraum",range=T,value = c(start,ende_default),
        #                                             minDate = start,maxDate = ende_max,language = "de",autoClose = T)})
        output$DAY_UI2<-renderUI({dateInput("DAY_End","Ende des Analysezeitraums",
                                            value=ende_default,
                                            min = start,
                                            max = ende_max,
                                            weekstart = 1,language = "de",format="yyyy-mm-dd")})                            
      }else{
        #output$DAY_UI1<-renderUI({airDatepickerInput('DAY_Start',"Analysezeitraum",range=T,value = c(start,ende_max),
        #                                             minDate = start,maxDate = ende_max,language = "de",autoClose = T)})
        output$DAY_UI2<-renderUI({dateInput("DAY_End","Ende des Analysezeitraums",
                                            value=ende_max,
                                            min = start,
                                            max = ende_max,
                                            weekstart = 1,language = "de",format="yyyy-mm-dd")})  
      }
      
    }
    })
    
    output$DAY_UI2t<-renderUI({materialSwitch(inputId = "day_type", 
                                              label = "Analyse über alle Kliniken?", status = "primary",value = T)})
   
    observe({
      if(!is.null(input$day_type)&&input$day_type==F&&!is.null(input$DAY_End)&&!is.null(input$DAY_Start)){
      input5_helper<-input5[input5$prüfdatum>=as.Date(input$DAY_Start)&input5$prüfdatum<=as.Date(input$DAY_End),]

      lexikon_fach<-data.frame(Abk=c("KAIT","KHAU","KEND","KGHI","CHG","KMKG","KGYN","KHAE","KCHH","KKAR","CHK","KKJP","KNEP",
                                     "KCHN","KNEU","KNRAD","KNUK","KAUG","KORT","KHNO","KPAE","KCHP","KPNE","KPSY","KPSM","KDR",
                                     "KSTR","KCHU","KURO","KCHI","KPHO","CHT",
                                     "KCHK","PNE","KNRA"),
                               Full=c("Anästhesiologie","Dermatologie","Endokrinologie","Gastro/Hepato/Infektiologie","Gefäßchirurgie",
                                      "Gesichtschirurgie","Gynäkologie","Hämato/Onkologie","Herz/Thoraxchirurgie","Kardiologie",
                                      "Kinderchirurgie","Kinderpsychiatrie","Nephrologie","Neurochirurgie","Neurologie","Neuroradiologie",
                                      "Nuklearmedizin","Ophtalmologie","Orthopädie","Oto/Rhino/Laryngologie","Pädiatrie","Plastische Chirurgie",
                                      "Pneumologie","Psychiatrie","Psychosomatik","Radiologie","Strahlentherapie","Unfallchirurgie",
                                      "Urologie","Viszeralchirurgie","Kinder Hämato/Onkologie","Thoraxchirurgie",
                                      "Kinderchirurgie","Pneumologie","Neuroradiologie"))
      
      day_fb<-as.data.frame(table(input5_helper$fachrichtung))
      if(nrow(day_fb)>0){
        day_fb$Var1<-paste0(day_fb$Var1," ",lexikon_fach$Full[match(day_fb$Var1,lexikon_fach$Abk)])
        
        day_fb1<-as.data.frame(table(input5$fachrichtung))
        day_fb1$Var1<-paste0(day_fb1$Var1," ",lexikon_fach$Full[match(day_fb1$Var1,lexikon_fach$Abk)])
        names(day_fb1)[2]<-"Freq_all"
        
        day_fb<-merge(day_fb,day_fb1,all=T)
        day_fb$Freq[is.na(day_fb$Freq)]<-0
        
        day_fb_choices<-paste0(day_fb$Var1," (n=",day_fb$Freq," von ",day_fb$Freq_all,")")
        
        output$DAY_UI2u<-renderUI({pickerInput('day_kliniken',label="Klinik (n 'im Analysezeitraum' von 'gesamt')",
                                               choices=day_fb_choices,
                                               options=list(`actions-box` = FALSE,
                                                            `live-search` = TRUE),
                                               multiple = F,inline = T)})
      }else{
        day_fb1<-as.data.frame(table(input5$fachrichtung))
        day_fb1$Var1<-paste0(day_fb1$Var1," ",lexikon_fach$Full[match(day_fb1$Var1,lexikon_fach$Abk)])
        names(day_fb1)[2]<-"Freq_all"
        day_fb1$Freq<-0
        
        day_fb_choices<-paste0(day_fb1$Var1," (n=",day_fb1$Freq," von ",day_fb1$Freq_all,")")
        
        output$DAY_UI2u<-renderUI({pickerInput('day_kliniken',label="Klinik (n 'im Analysezeitraum' von 'gesamt')",
                                               choices=day_fb_choices,
                                               options=list(`actions-box` = FALSE,
                                                            `live-search` = TRUE),
                                               multiple = F,inline = T)})
        
        shinyjs::disable("do_inDAY")
      }
      }
    })
    observe({if(!is.null(input$day_type)&&input$day_type==T){
      output$DAY_UI2u<-renderUI({NULL})
      shinyjs::enable("do_inDAY")
    }
    })
    
    

    
    
    output$DAY_UI3<-renderUI({h4("Normalbereiche Laborparameter:")})
    output$DAY_UI3a<-renderUI({div(id = "reverseSlider", 
                                   sliderInput("DAY_egfr","eGFR ml/min/1,73m²",min = 0,max = 150,value = 30))})
    output$DAY_UI3b<-renderUI({sliderInput("DAY_k","K mmol/l",min = 0,max = 7,value = c(2.8,5.2),step = 0.1)})
    output$DAY_UI3c<-renderUI({sliderInput("DAY_na","Na mmol/l",min = 100,max = 170,value = c(130,145))})
    
    
    
    output$DAY_UI4<-renderUI({h4("Darstellung Balkendiagramme:")})
    output$DAY_UI4a<-renderUI({pickerInput('barplot1',label = "Anzahl pro Tag",
                                           choices = c("Patienten/Tag","Patienten >=80 Jahre",
                                                       "Patienten eGFR unter Grenzwert ohne Dialyse",
                                                       "Patienten mit K außerhalb Grenzbereich",
                                                       "Patienten mit Na außerhalb Grenzbereich",
                                                       "Patienten für Medikationsanalyse",
                                                       "Patienten für ABS Visite"),
                                           selected = c("Patienten/Tag","Patienten >=80 Jahre",
                                                        "Patienten eGFR unter Grenzwert ohne Dialyse"),
                                           options = list(`actions-box` = FALSE,
                                                          #`deselect-all-text` = "Zurücksetzen",
                                                          #`select-all-text` = "Alles auswählen",
                                                          #`none-selected-text` = "Nichts ausgewählt",
                                                          `max-options`=3),multiple=TRUE)})
    
    output$DAY_UI4b<-renderUI({pickerInput('barplot2',label = "Gesamt und Mittelwert",
                                           choices = c("Patienten/Tag","Patienten >=80 Jahre",
                                                       "Patienten eGFR unter Grenzwert ohne Dialyse",
                                                       "Patienten mit K außerhalb Grenzbereich",
                                                       "Patienten mit Na außerhalb Grenzbereich",
                                                       "Patienten für Medikationsanalyse",
                                                       "Patienten für ABS Visite"),
                                           selected = c("Patienten/Tag","Patienten für Medikationsanalyse",
                                                        "Patienten für ABS Visite"),
                                           options = list(`actions-box` = TRUE,
                                                          `deselect-all-text` = "Zurücksetzen",
                                                          `select-all-text` = "Alles auswählen",
                                                          `none-selected-text` = "Nichts ausgewählt"),multiple=TRUE)})
    
    
    
    ####################################
    ##Medikationsprüfung
    output$text_analyse2a_med<-renderText({NULL})
    output$text_analyse2a2_med<-renderText({NULL})
    
    start_med<-as.Date(min(input5$prüfdatum),format="yyyy-mm-dd")
    helper<-as.numeric(strsplit(as.character(start_med),"-",fixed=T)[[1]][3])-1
    start_med<-start_med-helper
    
    monat_med<-as.numeric(strsplit(as.character(start_med),"-",fixed=T)[[1]][2])
    if(monat_med==1||monat_med==3||monat_med==5||monat_med==7||monat_med==8||monat_med==10||monat_med==12){
      monat_tage_med<-31
    }else{
      if(monat_med==4||monat_med==6||monat_med==9||monat_med==11){
        monat_tage_med<-30
      }else{
        if((as.numeric(strsplit(as.character(start_med),"-",fixed=T)[[1]][1])%%4)==0&&(as.numeric(strsplit(as.character(start_med),"-",fixed=T)[[1]][1])%%1000)!=0){
          monat_tage_med<-29
        }else{
          monat_tage_med<-28
        }
      }
    }
    ende_default_med<-start_med+monat_tage_med-1
    
    ende_med<-as.Date(max(input5$prüfdatum),format="yyyy-mm-dd")
    monat2_med<-as.numeric(strsplit(as.character(ende_med),"-",fixed=T)[[1]][2])
    if(monat2_med==1||monat2_med==3||monat2_med==5||monat2_med==7||monat2_med==8||monat2_med==10||monat2_med==12){
      monat_tage2_med<-31
    }else{
      if(monat2_med==4||monat2_med==6||monat2_med==9||monat2_med==11){
        monat_tage2_med<-30
      }else{
        if((as.numeric(strsplit(as.character(ende_med),"-",fixed=T)[[1]][1])%%4)==0&&(as.numeric(strsplit(as.character(ende_med),"-",fixed=T)[[1]][1])%%1000)!=0){
          monat_tage2_med<-29
        }else{
          monat_tage2_med<-28
        }
      }
    }
    helper<-as.numeric(strsplit(as.character(ende_med),"-",fixed=T)[[1]][3])
    ende_max_med<-ende_med-helper+monat_tage2_med
    
    start_anders_med<-paste(strsplit(as.character(start_med),"-",fixed = T)[[1]][c(3,2,1)],collapse = ".")
    ende_max_anders_med<-paste(strsplit(as.character(ende_max_med),"-",fixed = T)[[1]][c(3,2,1)],collapse = ".")
    
    output$MED_UI1<-renderUI({dateInput("MED_Start","Beginn des Analysezeitraums",
                                        value=start_med,
                                        min = start_med,
                                        max = ende_max_med,
                                        weekstart = 1,language = "de",format="yyyy-mm-dd")})
    
    output$MED_UI1a<-renderUI({h6(paste0("Aufgrund der eingelesenen Daten können nur Tage zwischen dem ",
                                         start_anders_med,
                                         " und ",ende_max_anders_med," ausgewählt werden.\n",
                                         "Standard ist ein Zeitraum von 1 Monat. Dies kann allerdings nach Bedarf geändert werden (Maximum ",
                                         as.numeric(as.Date(ende_max_med)-as.Date(start_med)),
                                         " Tage). Bitte beachten Sie, dass die Analyse im Fall von sehr langen Zeiträumen unübersichtlich werden kann."))})
    
    observe({if(!is.null(input$MED_Start)&&is.null(input$MED_End)){
      output$MED_UI1a<-renderUI({h6(paste0("Aufgrund der eingelesenen Daten können nur Tage zwischen dem ",
                                           start_anders_med,
                                           " und ",ende_max_anders_med," ausgewählt werden.\n",
                                           "Standard ist ein Zeitraum von 1 Monat. Dies kann allerdings nach Bedarf geändert werden (Maximum ",
                                           as.numeric(as.Date(ende_max_med)-as.Date(start_med)),
                                           " Tage). Bitte beachten Sie, dass die Analyse im Fall von sehr langen Zeiträumen unübersichtlich werden kann."))})
      
      if(ende_default_med<=ende_max_med){
        output$MED_UI2<-renderUI({dateInput("MED_End","Ende des Analysezeitraums",
                                            value=ende_default_med,
                                            min = start_med,
                                            max = ende_max_med,
                                            weekstart = 1,language = "de",format="yyyy-mm-dd")})                            
      }else{
        output$MED_UI2<-renderUI({dateInput("MED_End","Ende des Analysezeitraums",
                                            value=ende_max_med,
                                            min = start_med,
                                            max = ende_max_med,
                                            weekstart = 1,language = "de",format="yyyy-mm-dd")})  
      }
      
    }
      
      if(!is.null(input$MED_Start)&&!is.null(input$MED_End)){
        test<-input5[input5$prüfdatum>=as.Date(input$MED_Start)&input5$prüfdatum<=as.Date(input$MED_End),]
        
        kliniken<-unique(test$fachrichtung)
        output$MED_UI3<-renderUI({pickerInput('fb_in',"Analyse der Fachrichtungen mit den meisten Medikationsprüfungen",
                                              choices = c(paste0("Top-",c(1:(length(kliniken)-1))),"Alle"),selected = "Alle",multiple = F)})
        
        if(!is.null(input$med_type)&&input$med_type==F){
          output$MED_UI3<-renderUI({NULL})
        }
      }
    })

    
    observe({if(!is.null(input$MED_End)&&input$MED_Start>input$MED_End){
      start_med2<-as.Date(input$MED_Start,format="yyyy-mm-dd")
      helper<-as.numeric(strsplit(as.character(start_med2),"-",fixed=T)[[1]][3])-1
      start_med2<-start_med2-helper
      monat_med<-as.numeric(strsplit(as.character(start_med2),"-",fixed=T)[[1]][2])
      if(monat_med==1||monat_med==3||monat_med==5||monat_med==7||monat_med==8||monat_med==10||monat_med==12){
        monat_tage_med<-31
      }else{
        if(monat_med==4||monat_med==6||monat_med==9||monat_med==11){
          monat_tage_med<-30
        }else{
          if((as.numeric(strsplit(as.character(start_med2),"-",fixed=T)[[1]][1])%%4)==0&&(as.numeric(strsplit(as.character(start_med2),"-",fixed=T)[[1]][1])%%1000)!=0){
            monat_tage_med<-29
          }else{
            monat_tage_med<-28
          }
        }
      }
      start_med2<-input$MED_Start
      ende_default_med<-start_med2+monat_tage_med-1
      
      ende_med<-as.Date(max(input5$prüfdatum),format="yyyy-mm-dd")
      monat2_med<-as.numeric(strsplit(as.character(ende_med),"-",fixed=T)[[1]][2])
      if(monat2_med==1||monat2_med==3||monat2_med==5||monat2_med==7||monat2_med==8||monat2_med==10||monat2_med==12){
        monat_tage2_med<-31
      }else{
        if(monat2_med==4||monat2_med==6||monat2_med==9||monat2_med==11){
          monat_tage2_med<-30
        }else{
          if((as.numeric(strsplit(as.character(ende_med),"-",fixed=T)[[1]][1])%%4)==0&&(as.numeric(strsplit(as.character(ende_med),"-",fixed=T)[[1]][1])%%1000)!=0){
            monat_tage2_med<-29
          }else{
            monat_tage2_med<-28
          }
        }
      }
      helper<-as.numeric(strsplit(as.character(ende_med),"-",fixed=T)[[1]][3])
      ende_max_med<-ende_med-helper+monat_tage2_med
      
      ende_max_anders_med<-paste(strsplit(as.character(ende_max_med),"-",fixed = T)[[1]][c(3,2,1)],collapse = ".")
      
      if(ende_default_med<=ende_max_med){
        output$MED_UI2<-renderUI({dateInput("MED_End","Ende des Analysezeitraums",
                                            value=ende_default_med,
                                            min = start_med,
                                            max = ende_max_med,
                                            weekstart = 1,language = "de",format="yyyy-mm-dd")})                            
      }else{
        output$MED_UI2<-renderUI({dateInput("MED_End","Ende des Analysezeitraums",
                                            value=ende_max_med,
                                            min = start_med,
                                            max = ende_max_med,
                                            weekstart = 1,language = "de",format="yyyy-mm-dd")})  
      }
      
    }
    })
    
    
    output$MED_UI2t<-renderUI({materialSwitch(inputId = "med_type", 
                                              label = "Analyse über alle Kliniken?", status = "primary",value = T)})
    
    observe({if(!is.null(input$med_type)&&input$med_type==F&&!is.null(input$MED_End)&&!is.null(input$MED_Start)){
      input5_helper_med<-input5[input5$prüfdatum>=as.Date(input$MED_Start)&input5$prüfdatum<=as.Date(input$MED_End),]
      
      lexikon_fach<-data.frame(Abk=c("KAIT","KHAU","KEND","KGHI","CHG","KMKG","KGYN","KHAE","KCHH","KKAR","CHK","KKJP","KNEP",
                                     "KCHN","KNEU","KNRAD","KNUK","KAUG","KORT","KHNO","KPAE","KCHP","KPNE","KPSY","KPSM","KDR",
                                     "KSTR","KCHU","KURO","KCHI","KPHO","CHT",
                                     "KCHK","PNE","KNRA"),
                               Full=c("Anästhesiologie","Dermatologie","Endokrinologie","Gastro/Hepato/Infektiologie","Gefäßchirurgie",
                                      "Gesichtschirurgie","Gynäkologie","Hämato/Onkologie","Herz/Thoraxchirurgie","Kardiologie",
                                      "Kinderchirurgie","Kinderpsychiatrie","Nephrologie","Neurochirurgie","Neurologie","Neuroradiologie",
                                      "Nuklearmedizin","Ophtalmologie","Orthopädie","Oto/Rhino/Laryngologie","Pädiatrie","Plastische Chirurgie",
                                      "Pneumologie","Psychiatrie","Psychosomatik","Radiologie","Strahlentherapie","Unfallchirurgie",
                                      "Urologie","Viszeralchirurgie","Kinder Hämato/Onkologie","Thoraxchirurgie",
                                      "Kinderchirurgie","Pneumologie","Neuroradiologie"))
      
      med_fb<-as.data.frame(table(input5_helper_med$fachrichtung))
      if(nrow(med_fb)>0){
        med_fb$Var1<-paste0(med_fb$Var1," ",lexikon_fach$Full[match(med_fb$Var1,lexikon_fach$Abk)])
        
        med_fb1<-as.data.frame(table(input5$fachrichtung))
        med_fb1$Var1<-paste0(med_fb1$Var1," ",lexikon_fach$Full[match(med_fb1$Var1,lexikon_fach$Abk)])
        names(med_fb1)[2]<-"Freq_all"
        
        med_fb<-merge(med_fb,med_fb1,all=T)
        med_fb$Freq[is.na(med_fb$Freq)]<-0
        
        med_fb_choices<-paste0(med_fb$Var1," (n=",med_fb$Freq," von ",med_fb$Freq_all,")")
        
        output$MED_UI2u<-renderUI({pickerInput('med_kliniken',label="Klinik (n 'im Analysezeitraum' von 'gesamt')",
                                               choices=med_fb_choices,
                                               options=list(`actions-box` = FALSE,
                                                            `live-search` = TRUE),
                                               multiple = F,inline = T)})        
      }else{
        med_fb1<-as.data.frame(table(input5$fachrichtung))
        med_fb1$Var1<-paste0(med_fb1$Var1," ",lexikon_fach$Full[match(med_fb1$Var1,lexikon_fach$Abk)])
        names(med_fb1)[2]<-"Freq_all"
        med_fb1$Freq<-0
        
        med_fb_choices<-paste0(med_fb1$Var1," (n=",med_fb1$Freq," von ",med_fb1$Freq_all,")")
        
        output$MED_UI2u<-renderUI({pickerInput('med_kliniken',label="Klinik (n 'im Analysezeitraum' von 'gesamt')",
                                               choices=med_fb_choices,
                                               options=list(`actions-box` = FALSE,
                                                            `live-search` = TRUE),
                                               multiple = F,inline = T)})
        
        shinyjs::disable("do_inMED")
      }

    }
    })
    observe({if(!is.null(input$med_type)&&input$med_type==T){
      output$MED_UI2u<-renderUI({NULL})
      test<-input5[input5$prüfdatum>=as.Date(input$MED_Start)&input5$prüfdatum<=as.Date(input$MED_End),]
      
      kliniken<-unique(test$fachrichtung)
      output$MED_UI3<-renderUI({pickerInput('fb_in',"Analyse der Fachrichtungen mit den meisten Medikationsprüfungen",
                                            choices = c(paste0("Top-",c(1:(length(kliniken)-1))),"Alle"),selected = "Alle",multiple = F)})
      
      shinyjs::enable("do_inMED")
    }
    })
    
    observe({if(!is.null(input$med_type)&&input$med_type==F){
      output$MED_UI3<-renderUI({NULL})
    }
    })
    
    
    observe({
      if(!is.null(input$day_type)&&input$day_type==F&&!is.null(input$day_kliniken)&&
         as.numeric(strsplit(strsplit(input$day_kliniken," (n=",fixed = T)[[1]][2]," ",fixed=T)[[1]][1])==0){
        shinyjs::disable("do_inDAY")
      }
      if(!is.null(input$day_type)&&input$day_type==F&&!is.null(input$day_kliniken)&&
         as.numeric(strsplit(strsplit(input$day_kliniken," (n=",fixed = T)[[1]][2]," ",fixed=T)[[1]][1])>0){
        shinyjs::enable("do_inDAY")
      }
    })
    observe({
      if(!is.null(input$med_type)&&input$med_type==F&&!is.null(input$med_kliniken)&&
         as.numeric(strsplit(strsplit(input$med_kliniken," (n=",fixed = T)[[1]][2]," ",fixed=T)[[1]][1])==0){
        shinyjs::disable("do_inMED")
      }
      if(!is.null(input$med_type)&&input$med_type==F&&!is.null(input$med_kliniken)&&
         as.numeric(strsplit(strsplit(input$med_kliniken," (n=",fixed = T)[[1]][2]," ",fixed=T)[[1]][1])>0){
        shinyjs::enable("do_inMED")
      }
    })

    
    shinyjs::disable("do_in")
    
    observeEvent(input$do_out,{
      shinyjs::html("text", paste0("<br><br>Exportiere kompletten Datensatz.<br><br>"), add = FALSE)
      
      updateTabsetPanel(session,"main",
                        selected="Log")
      #message("hallo")
      wb<-createWorkbook()
      addWorksheet(wb,sheetName="Datensatz")
      writeData(wb,"Datensatz",input5b_export,rowNames = F)
      bold<-createStyle(textDecoration = "bold",wrapText = T)
      addStyle(wb,"Datensatz",bold,rows = 1,cols=c(1:ncol(input5b_export)))      
      
      width_vec1<-apply(input5b_export,2,function(x){max(nchar(as.character(x)) + 2,1, na.rm = TRUE)})
      width_vec2<-nchar(names(input5b_export)) + 2
      width_vec<-apply(cbind(width_vec1,width_vec2),1,max)
      width_vec[width_vec>50]<-50
      setColWidths(wb,"Datensatz",cols = c(1:ncol(input5b_export)),widths = width_vec)
      name_export<-paste0("Export Pharmazeutische Kurvenvisite ",start_anders_med,"-",ende_max_anders_med,".xlsx")
      #pfad<-paste0("C:",paste(unlist(input$export_dir[[1]]),collapse="/"))
      
        if(paste(unlist(input$export_dir[[1]]),collapse="/")==0||paste(unlist(input$export_dir[[1]]),collapse="/")==1){
          #pfad<-"~"
          pfad<-"C:/"
        }else{
          #pfad<-paste0("~",paste(unlist(input$export_dir[[1]]),collapse="/"))
          pfad<-paste0("C:",paste(unlist(input$export_dir[[1]]),collapse="/"))
        }
      
      tt <- tryCatch(saveWorkbook(wb,file=paste0(pfad,"/",name_export),overwrite = T),
                     warning=function(w) w)
      if(is.list(tt)==F){
        saveWorkbook(wb,file=paste0(pfad,"/",name_export),overwrite = T)
      }else{
        if(length(grep("Permission denied",tt$message))>0){
          shinyjs::html("text", paste0("<br><br>Zugriff verweitert.<br><br>
                                         Falls Sie die exportierte Datei aus einer vorherigen Analyse noch geöffnet haben: bitte schließen Sie diese oder ändern Sie den Dateinamen, damit ein Update exportiert werden kann. 
                                         <br>Bitte stellen Sie weiterhin sicher, dass Sie für den ausgewählten Ordner Schreibrechte besitzen (für C:/ häufig nicht der Fall)."), add = TRUE)
        }else{
          shinyjs::html("text", paste0("<br><br>",tt$message), add = TRUE)
        }
        return()
      }
      
      saveWorkbook(wb,file=paste0(pfad,"/",name_export),overwrite = T)
      
      
      
      shinyjs::html("text", paste0("<br><br>Kompletter Datensatz erfolgreich exportiert.<br><br>"), add = TRUE)
      shinyjs::html("text", paste0("&nbsp;&nbsp;&nbsp;&nbsp;",paste0(pfad,"/",name_export),"<br>"), add = TRUE)
      
    })
    
    observeEvent(input$do_inDAY,{
      updateTabsetPanel(session,"main",
                        selected="Patienten pro Tag")
      
      input6<-input5[input5$prüfdatum>=as.Date(input$DAY_Start)&input5$prüfdatum<=as.Date(input$DAY_End),]
      values$DAY_Start<-input$DAY_Start
      values$DAY_End<-input$DAY_End
      
      if(!is.null(input$day_type)&&input$day_type==F){
        values$day_type<-F
        values$day_kliniken<-input$day_kliniken
        help_klinik1<-strsplit(input$day_kliniken," (n=",fixed = T)[[1]][1]
        help_klinik1b<-strsplit(help_klinik1," ",fixed=T)[[1]][1]
        
        help_klinik2<-strsplit(input$day_kliniken," (n=",fixed = T)[[1]][2]
        help_klinik2b<-strsplit(help_klinik2," ",fixed=T)[[1]][1]
        help_klinik2b<-as.numeric(help_klinik2b)
        
        if(help_klinik2b==0){
          output$text_analyse2b<-renderText({paste0("Für Klinik ",help_klinik1," gibt es im gewählten Zeitraum keine Beobachtungen.")})
          output$text_analyse2c<-renderText({paste0("Für Klinik ",help_klinik1," gibt es im gewählten Zeitraum keine Beobachtungen.")})
          output$data_proTag<-renderDataTable({NULL})
          output$egfr_proTag<-renderDataTable({NULL})
          output$barplot_proTag<-renderPlot({NULL})
          output$barplot_proTag2<-renderPlot({NULL})
          output$barplot_proTag3<-renderPlot({NULL})
          return()
        }
        
        input6<-input6[input6$fachrichtung==help_klinik1b,]
      }else{
        values$day_type<-T
        values$day_kliniken<-NA
      }
      
      shinyjs::html("text", paste0("<br>Analysiere Patienten pro Tag.<br><br><br>"), add = FALSE)
      if(sum(is.na(suppressWarnings(as.numeric(input6$egfr))))-sum(is.na(input6$egfr))>0)
        shinyjs::html("text", paste0("<br>",sum(is.na(suppressWarnings(as.numeric(input6$egfr))))-sum(is.na(input6$egfr)),
                                     " eGFR-Werte können nicht ausgewertet werden (falsches Format).<br>"), add = TRUE)
      input6$egfr<-suppressWarnings(as.numeric(input6$egfr))
      
      if(sum(is.na(suppressWarnings(as.numeric(input6$kalium))))-sum(is.na(input6$kalium))>0)
      shinyjs::html("text", paste0("<br>",sum(is.na(suppressWarnings(as.numeric(input6$kalium))))-sum(is.na(input6$kalium)),
                                   " Kalium-Werte können nicht ausgewertet werden (falsches Format).<br>"), add = TRUE)
      input6$kalium<-suppressWarnings(as.numeric(input6$kalium))
      
      if(sum(is.na(suppressWarnings(as.numeric(input6$natrium))))-sum(is.na(input6$natrium))>0)
      shinyjs::html("text", paste0("<br>",sum(is.na(suppressWarnings(as.numeric(input6$natrium))))-sum(is.na(input6$natrium)),
                                   " Natrium-Werte können nicht ausgewertet werden (falsches Format).<br>"), add = TRUE)
      input6$natrium<-suppressWarnings(as.numeric(input6$natrium))
      
      if(sum(is.na(suppressWarnings(as.numeric(input6$`antibiotika.x.tage.(abs.visite)`))))-sum(is.na(input6$`antibiotika.x.tage.(abs.visite)`))>0)
        shinyjs::html("text", paste0("<br>",sum(is.na(suppressWarnings(as.numeric(input6$`antibiotika.x.tage.(abs.visite)`))))-sum(is.na(input6$`antibiotika.x.tage.(abs.visite)`)),
                                     " Angaben zur ABS Visite können nicht ausgewertet werden (falsches Format).<br>"), add = TRUE)
      input6$`antibiotika.x.tage.(abs.visite)`<-suppressWarnings(as.numeric(input6$`antibiotika.x.tage.(abs.visite)`))
      
      
      output_proTag<-data.frame(Datum=as.Date(as.Date(input$DAY_Start):as.Date(input$DAY_End)),
                                Patienten.Tag=0,
                                Patienten.groesser.gleich.80.Jahre=0,
                                egfr.ohne.Dialyse=0,
                                k=0,
                                na=0,
                                Patienten.fuer.Medikationsanalyse=0,
                                Patienten.fuer.ABS.Visite=0,
                                Datum_intern=as.Date(as.Date(input$DAY_Start):as.Date(input$DAY_End)))
      
      wochentage<-c("Sonntag","Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag")
      
      helper<-str_split_fixed(output_proTag$Datum,pattern = "-",Inf)
      helper2<-paste0(helper[,3],".",helper[,2],".",helper[,1])
      output_proTag$Datum<-paste0(wochentage[1+as.POSIXlt(output_proTag$Datum)$wday],", ",helper2)
      ##tagesdarstellung optimieren
      
      for(i in 1:length(output_proTag$Datum_intern)){
        output_proTag$Patienten.Tag[i]<-sum(input6$prüfdatum==output_proTag$Datum_intern[i])
        output_proTag$Patienten.groesser.gleich.80.Jahre[i]<-sum(input6$prüfdatum==output_proTag$Datum_intern[i]&input6$alter>=80)
        output_proTag$Patienten.fuer.Medikationsanalyse[i]<-sum(input6$prüfdatum==output_proTag$Datum_intern[i]&input6$medikationsprüfungen.pharmazeutische.visite!="nicht geprüft")
        output_proTag$egfr.ohne.Dialyse[i]<-sum(input6$prüfdatum==output_proTag$Datum_intern[i]&is.na(input6$dialyse)&!is.na(input6$egfr)&input6$egfr<input$DAY_egfr)
        output_proTag$k[i]<-sum(input6$prüfdatum==output_proTag$Datum_intern[i]&is.na(input6$dialyse)&!is.na(input6$k)&(input6$k<input$DAY_k[1]|input6$k>input$DAY_k[2]))
        output_proTag$na[i]<-sum(input6$prüfdatum==output_proTag$Datum_intern[i]&is.na(input6$dialyse)&!is.na(input6$na)&(input6$k<input$DAY_na[1]|input6$k>input$DAY_na[2]))
        output_proTag$Patienten.fuer.ABS.Visite[i]<-sum(input6$prüfdatum==output_proTag$Datum_intern[i]&!is.na(input6$`antibiotika.x.tage.(abs.visite)`)&input6$`antibiotika.x.tage.(abs.visite)`>0)
        
      }
      output_proTag<-rbind(output_proTag,rep(NA,9),rep(NA,9))
      output_proTag[(nrow(output_proTag)-1),1]<-"Gesamtsumme"
      output_proTag[(nrow(output_proTag)-1),2:(ncol(output_proTag)-1)]<-colSums(output_proTag[1:(nrow(output_proTag)-2),2:(ncol(output_proTag)-1)])
      
      output_proTag[nrow(output_proTag),1]<-"Mittelwert pro Arbeitstg (ohne 0)"
      output_proTag[nrow(output_proTag),2:(ncol(output_proTag)-1)]<-apply(output_proTag[1:(nrow(output_proTag)-2),2:(ncol(output_proTag)-1)],MARGIN = 2,FUN = function(x){
        return(round(mean(x[x!=0])))
      })
      
      output_proTag2<-output_proTag[,1:8]
      names(output_proTag2)[2]<-"Patienten/Tag"
      names(output_proTag2)[3]<-"Patienten >=80 Jahre"
      names(output_proTag2)[4]<-paste0("Patienten eGFR <",input$DAY_egfr," ml/min ohne Dialyse")
      names(output_proTag2)[5]<-paste0("Patienten mit K <",input$DAY_k[1]," oder >",input$DAY_k[2]," mmol/l")
      names(output_proTag2)[6]<-paste0("Patienten mit Na <",input$DAY_na[1]," oder >",input$DAY_na[2]," mmol/l")
      names(output_proTag2)[7]<-paste0("Patienten für Medikationsanalyse")
      names(output_proTag2)[8]<-paste0("Patienten für ABS Visite")
      
      
      
      output$data_proTag<-renderDataTable(datatable(output_proTag2,rownames = F,
                                                    options=list(language=list(search='Suchen:',lengthMenu='Zeige _MENU_ Einträge',
                                                                               info='Zeige _START_ bis _END_ von _TOTAL_ Einträgen',
                                                                               paginate=list(previous='Vorherige','next'='Nächste')))))
      output$text_analyse2b<-renderText({NULL})
      
      
      ##eGFR Sonderanalyse
      egfr<-data.frame(Kategorie=c("eGFR >= 90ml/min","60ml/min <= eGFR < 90ml/min",
                                   "30ml/min <= eGFR < 60ml/min","15ml/min <= eGFR < 30ml/min",
                                   "eGFR < 15ml/min","Dialyse","keine Angabe"),
                       Abs=c(sum(!is.na(input6$egfr)&input6$egfr>=90&is.na(input6$dialyse)),
                             sum(!is.na(input6$egfr)&input6$egfr<90&input6$egfr>=60&is.na(input6$dialyse)),
                             sum(!is.na(input6$egfr)&input6$egfr<60&input6$egfr>=30&is.na(input6$dialyse)),
                             sum(!is.na(input6$egfr)&input6$egfr<30&input6$egfr>=15&is.na(input6$dialyse)),
                             sum(!is.na(input6$egfr)&input6$egfr<15&is.na(input6$dialyse)),
                             sum(!is.na(input6$dialyse)),
                             sum(is.na(input6$egfr&is.na(input6$dialyse))))
        )
      egfr$Rel1<-paste0(round(100*egfr$Abs/sum(egfr$Abs[1:6])),"%")
      egfr$Rel1[egfr$Rel1=="0%"&egfr$Fälle!=0]<-"<1%"
      egfr$Rel1[7]<-NA
      egfr$Rel2<-round(100*egfr$Abs/sum(egfr$Abs[1:6]))
      names(egfr)<-c("Kategorie","Fälle","Relativ","Rel")
      egfr$Relativ[egfr$Relativ=="NaN%"]<-"0%"
      output$egfr_proTag<-renderDataTable(datatable(egfr[,1:3],rownames = F,
                                                    options=list(
                                                      columnDefs=list(list(className = 'dt-right',targets=2)),
                                                      language=list(search='Suchen:',lengthMenu='Zeige _MENU_ Einträge',
                                                                               info='Zeige _START_ bis _END_ von _TOTAL_ Einträgen',
                                                                               paginate=list(previous='Vorherige','next'='Nächste')))))
      
      
      
      ######################
      ##barplot 1
      auswahl<-input$barplot1
      auswahl<-gsub("Patienten eGFR unter Grenzwert ohne Dialyse",
                    paste0("Patienten eGFR <",input$DAY_egfr," ml/min ohne Dialyse"),auswahl)
      auswahl<-gsub("Patienten mit K außerhalb Grenzbereich",
                    paste0("Patienten mit K <",input$DAY_k[1]," oder >",input$DAY_k[2]," mmol/l"),auswahl)
      auswahl<-gsub("Patienten mit Na außerhalb Grenzbereich",
                    paste0("Patienten mit Na <",input$DAY_na[1]," oder >",input$DAY_na[2]," mmol/l"),auswahl)
      
      auswahl2<-input$barplot1
      auswahl2<-gsub("Patienten/Tag","Patienten.Tag",auswahl2)
      auswahl2<-gsub("Patienten >=80 Jahre","Patienten.groesser.gleich.80.Jahre",auswahl2)
      auswahl2<-gsub("Patienten eGFR unter Grenzwert ohne Dialyse","egfr.ohne.Dialyse",auswahl2)
      auswahl2<-gsub("Patienten mit K außerhalb Grenzbereich","k",auswahl2)
      auswahl2<-gsub("Patienten mit Na außerhalb Grenzbereich","na",auswahl2)
      auswahl2<-gsub("Patienten für Medikationsanalyse","Patienten.fuer.Medikationsanalyse",auswahl2)
      auswahl2<-gsub("Patienten für ABS Visite","Patienten.fuer.ABS.Visite",auswahl2)
      
      auswahl3<-auswahl
      auswahl3<-gsub("Patienten","",auswahl3)
      auswahl3<-gsub("/Tag","pro Tag",auswahl3,fixed=T)
      auswahl3<-gsub("^ ","",auswahl3)

      output$barplot_proTag<-renderPlot({
        par(mar=c(9,3,5,1))
        if(sum(t(as.matrix(output_proTag[1:(nrow(output_proTag)-2),auswahl2])),na.rm=T)>0){
          if(length(auswahl)==3){
            barplot(t(as.matrix(output_proTag[1:(nrow(output_proTag)-2),auswahl2])),
                    beside=T,las=2,col=c("grey","skyblue","orange"),xaxt="n",
                    ylim=c(0,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1),
                    main=paste0("Anzahl Patienten ",paste(auswahl3,collapse = ", ")),cex.main=1.9)
            text(x=seq(2.5,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                 y=rep(-1*max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)/55,(nrow(output_proTag)-2)),
                 labels=output_proTag$Datum[1:(nrow(output_proTag)-2)],srt=60,adj=1,xpd=T)
          }else{
            if(length(auswahl)==2){
              barplot(t(as.matrix(output_proTag[1:(nrow(output_proTag)-2),auswahl2])),
                      beside=T,las=2,col=c("grey","skyblue"),xaxt="n",
                      ylim=c(0,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1),
                      main=paste0("Anzahl Patienten ",paste(auswahl3,collapse = ", ")),cex.main=1.9)
              text(x=seq(2,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                   y=rep(-1*(max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1)/55,(nrow(output_proTag)-2)),
                   labels=output_proTag$Datum[1:(nrow(output_proTag)-2)],srt=60,adj=1,xpd=T)
            }else{
              if(length(auswahl)==1){
                barplot(t(as.matrix(output_proTag[1:(nrow(output_proTag)-2),auswahl2])),
                        beside=T,las=2,col=c("grey"),xaxt="n",
                        ylim=c(0,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1),
                        main=paste0("Anzahl Patienten ",paste(auswahl3,collapse = ", ")),cex.main=1.9)
                text(x=seq(1.5,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                     y=rep(-1*(max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1)/55,(nrow(output_proTag)-2)),
                     labels=output_proTag$Datum[1:(nrow(output_proTag)-2)],srt=60,adj=1,xpd=T)
              }
            }
          }
          
          for(i in 1:length(auswahl2)){
            if(nrow(output_proTag)<=6){
              text(x=seq(0.5+i,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                   y=output_proTag[1:(nrow(output_proTag)-2),auswahl2[i]]+max(max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1*0.04,0.4),
                   output_proTag[1:(nrow(output_proTag)-2),auswahl2[i]],cex=2)
            }else{
              text(x=seq(0.5+i,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                   y=output_proTag[1:(nrow(output_proTag)-2),auswahl2[i]]+max(max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1*0.04,0.4),
                   output_proTag[1:(nrow(output_proTag)-2),auswahl2[i]],cex=1/(nrow(output_proTag)-2)*15)
            }
          }
          
          if((max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1)>50){
            for(i in seq(50,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1,50)){
              abline(h=i,col="gray80")
            }
          }
          
          legend("topright",legend=auswahl[1],bty="n",fill="grey",cex=1.3)
          if(length(auswahl)>1)
            legend("top",legend=auswahl[2],bty="n",fill="skyblue",cex=1.3)
          if(length(auswahl)>2)
            legend("topleft",legend=auswahl[3],bty="n",fill="orange",cex=1.3)
          
        }else{
          if(length(auswahl)>0){
            if(length(auswahl)==3){
              barplot(t(as.matrix(output_proTag[1:(nrow(output_proTag)-2),auswahl2])),
                      beside=T,las=2,col=c("grey","skyblue","orange"),xaxt="n",
                      ylim=c(0,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1),
                      main=paste0("Anzahl Patienten ",paste(auswahl3,collapse = ", ")),cex.main=1.9)
              text(x=seq(2.5,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                   y=rep(-1*(max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1)/55,(nrow(output_proTag)-2)),
                   labels=output_proTag$Datum[1:(nrow(output_proTag)-2)],srt=60,adj=1,xpd=T)
            }else{
              if(length(auswahl)==2){
                barplot(t(as.matrix(output_proTag[1:(nrow(output_proTag)-2),auswahl2])),
                        beside=T,las=2,col=c("grey","skyblue"),xaxt="n",
                        ylim=c(0,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1),
                        main=paste0("Anzahl Patienten ",paste(auswahl3,collapse = ", ")),cex.main=1.9)
                text(x=seq(2,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                     y=rep(-1*(max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1)/55,(nrow(output_proTag)-2)),
                     labels=output_proTag$Datum[1:(nrow(output_proTag)-2)],srt=60,adj=1,xpd=T)
              }else{
                if(length(auswahl)==1){
                  barplot(t(as.matrix(output_proTag[1:(nrow(output_proTag)-2),auswahl2])),
                          beside=T,las=2,col=c("grey"),xaxt="n",
                          ylim=c(0,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1),
                          main=paste0("Anzahl Patienten ",paste(auswahl3,collapse = ", ")),cex.main=1.9)
                  text(x=seq(1.5,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                       y=rep(-1*(max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1)/55,(nrow(output_proTag)-2)),
                       labels=output_proTag$Datum[1:(nrow(output_proTag)-2)],srt=60,adj=1,xpd=T)
                }
              }
            }
            
            for(i in 1:length(auswahl2)){
              if(nrow(output_proTag)<=6){
                text(x=seq(0.5+i,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                     y=output_proTag[1:(nrow(output_proTag)-2),auswahl2[i]]+max(0.4,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1*0.04),
                     output_proTag[1:(nrow(output_proTag)-2),auswahl2[i]],cex=2)
              }else{
                text(x=seq(0.5+i,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                     y=output_proTag[1:(nrow(output_proTag)-2),auswahl2[i]]+max(0.4,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1*0.04),
                     output_proTag[1:(nrow(output_proTag)-2),auswahl2[i]],cex=1/(nrow(output_proTag)-2)*15)
              }
            }
            
            if((max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1)>50){
              for(i in seq(50,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1,50)){
                abline(h=i,col="gray80")
              }
            }
            
            legend("topright",legend=auswahl[1],bty="n",fill="grey",cex=1.3)
            if(length(auswahl)>1)
              legend("top",legend=auswahl[2],bty="n",fill="skyblue",cex=1.3)
            if(length(auswahl)>2)
              legend("topleft",legend=auswahl[3],bty="n",fill="orange",cex=1.3)
            
            
            
          }else{
            plot(NULL,xlim=c(0,1),ylim=c(0,1),xaxt="n",yaxt="n",xlab="",ylab="",bty="n",
                 main=paste0("Anzahl Patienten ",paste(auswahl3,collapse = ", ")),cex.main=1.9)
            text(0.5,0.5,"Keine Daten vorhanden",cex=2) 
          }

        }
        
      })
      output$text_analyse2c<-renderText({NULL})
      
      
      ######################
      ##barplot 2
      auswahl_b<-input$barplot2
      auswahl_b<-gsub("Patienten eGFR unter Grenzwert ohne Dialyse",
                    paste0("Patienten eGFR <",input$DAY_egfr,"\nml/min ohne Dialyse"),auswahl_b)
      auswahl_b<-gsub("Patienten mit K außerhalb Grenzbereich",
                    paste0("Patienten mit K <",input$DAY_k[1],"\noder >",input$DAY_k[2]," mmol/l"),auswahl_b)
      auswahl_b<-gsub("Patienten mit Na außerhalb Grenzbereich",
                    paste0("Patienten mit Na <",input$DAY_na[1],"\noder >",input$DAY_na[2]," mmol/l"),auswahl_b)
      auswahl_b<-gsub("Patienten für Medikationsanalyse",
                      paste0("Patienten für\nMedikationsanalyse"),auswahl_b)
      auswahl_b<-gsub("Patienten für ABS Visite",
                      paste0("Patienten für\nABS Visite"),auswahl_b)
      
      auswahl2b<-input$barplot2
      auswahl2b<-gsub("Patienten/Tag","Patienten.Tag",auswahl2b)
      auswahl2b<-gsub("Patienten >=80 Jahre","Patienten.groesser.gleich.80.Jahre",auswahl2b)
      auswahl2b<-gsub("Patienten eGFR unter Grenzwert ohne Dialyse","egfr.ohne.Dialyse",auswahl2b)
      auswahl2b<-gsub("Patienten mit K außerhalb Grenzbereich","k",auswahl2b)
      auswahl2b<-gsub("Patienten mit Na außerhalb Grenzbereich","na",auswahl2b)
      auswahl2b<-gsub("Patienten für Medikationsanalyse","Patienten.fuer.Medikationsanalyse",auswahl2b)
      auswahl2b<-gsub("Patienten für ABS Visite","Patienten.fuer.ABS.Visite",auswahl2b)
      
      auswahl3b<-auswahl_b
      auswahl3b<-gsub("Patienten","",auswahl3b)
      auswahl3b<-gsub("/Tag","pro Tag",auswahl3b,fixed=T)
      auswahl3b<-gsub("^ ","",auswahl3b)
      
      output$barplot_proTag2<-renderPlot({
        par(mar=c(4,5,6,1))
        if(sum(t(as.matrix(output_proTag[nrow(output_proTag)-1,auswahl2b])))>0){
          layout(matrix(data = c(1,1,2,2,3),ncol = 5))
          barplot(t(as.matrix(output_proTag[nrow(output_proTag)-1,auswahl2b])),
                  beside=T,las=1,col=c("#3E7C90","#E69EBB","#EC952E","#F0D7E2","#9EBC9F","#E7EFAF","#6A6BE5","#EEED39"[1:length(auswahl_b)]),
                  ylim=c(0,max(6,round(max(output_proTag[nrow(output_proTag)-1,auswahl2b],na.rm=T),-1)*1.2)),
                  #names.arg=auswahl_b,
                  main=paste0("Gesamtsumme \n",
                              paste(strsplit(as.character(input$DAY_Start),"-")[[1]][c(3,2,1)],collapse = "."),
                              " bis ",
                              paste(strsplit(as.character(input$DAY_End),"-")[[1]][c(3,2,1)],collapse = ".")),
                  cex.main=2,cex.axis=1.5,names.arg="")
          abline(h=0)
          
          sd<-apply(as.data.frame(output_proTag[1:(nrow(output_proTag)-2),auswahl2b]),2,sd)
          sd[is.na(sd)]<-0
          
          barplot(t(as.matrix(output_proTag[nrow(output_proTag),auswahl2b])),
                  beside=T,las=1,col=c("#3E7C90","#E69EBB","#EC952E","#F0D7E2","#9EBC9F","#E7EFAF","#6A6BE5","#EEED39"[1:length(auswahl_b)]),
                  ylim=c(0,max(6,round(max(output_proTag[nrow(output_proTag),auswahl2b]+sd,na.rm = T),-1)*1.2)),xaxt="n",
                  #names.arg=auswahl_b,
                  main=paste0("Mittelwert pro Arbeitstag (ohne 0) \n",
                              paste(strsplit(as.character(input$DAY_Start),"-")[[1]][c(3,2,1)],collapse = "."),
                              " bis ",
                              paste(strsplit(as.character(input$DAY_End),"-")[[1]][c(3,2,1)],collapse = ".")),
                  cex.main=2,cex.axis=1.5)
          abline(h=0)
          
          for(i in 1:length(sd)){
            points(x=c(i+0.5,i+0.5),y=c(output_proTag[nrow(output_proTag),auswahl2b[i]]-sd[i],output_proTag[nrow(output_proTag),auswahl2b[i]]+sd[i]),
                   type="l",lwd=4)
          }
          par(mar=c(0.5,0.5,0.5,0.5))
          plot(NULL,xlim=c(0,1),ylim=c(0,2),xaxt="n",yaxt="n",xlab="",ylab="",bty="n")
          legend("center",legend=auswahl_b,cex=1.7,y.intersp = 2,
                 fill=c("#3E7C90","#E69EBB","#EC952E","#F0D7E2","#9EBC9F","#E7EFAF","#6A6BE5","#EEED39"[1:length(auswahl_b)]))
          
        }else{
          plot(NULL,xlim=c(0,1),ylim=c(0,1),xaxt="n",yaxt="n",xlab="",ylab="",bty="n",
               main=paste0("Gesamtsumme \n",
                           paste(strsplit(as.character(input$DAY_Start),"-")[[1]][c(3,2,1)],collapse = "."),
                           " bis ",
                           paste(strsplit(as.character(input$DAY_End),"-")[[1]][c(3,2,1)],collapse = ".")),cex.main=1.9)
          text(0.5,0.5,"Keine Daten vorhanden",cex=2)
        }
        })
      
      
      ##barplot 3
      output$barplot_proTag3<-renderPlot({
        par(mar=c(12,5,5,5.5))
        if(sum(egfr$Fälle[1:6])>0){
          barplot(egfr$Fälle[1:6],las=2,col=c(colorRampPalette(c("skyblue","orange"))(5),"grey"),xaxt="n",
                  ylim=c(0,round(max(egfr$Fälle[1:6],na.rm=T))*1.1),
                  main=paste0("Verteilung eGFR"),cex.main=1.9,ylab="Absolut",cex.lab=2,cex.axis=1.3)
          add_helper<-seq(0,round(max(egfr$Fälle[1:6],na.rm=T))*1.1,sum(egfr$Fälle[1:6])/10)
          axis(4,at=add_helper,
               labels = paste0(seq(0,10*(length(add_helper)-1),10),"%"),las=2,cex.axis=1.4)
          mtext("Relativ",side = 4,line = 4.5,cex=2)
          text(x=seq(0.7,7,1.2),
               y=rep(-1*(round(max(egfr$Fälle[1:6],na.rm=T))*1.1)/17,6),cex=1.07,
               labels=egfr$Kategorie[1:6],srt=60,adj=1,xpd=T)
        }else{
          plot(NULL,xlim=c(0,1),ylim=c(0,1),xaxt="n",yaxt="n",xlab="",ylab="",bty="n",
               main=paste0("Verteilung eGFR"),cex.main=1.9)
          text(0.5,0.5,"Keine Daten vorhanden",cex=2)
        }
      })
      
      shinyjs::html("text", paste0("<br>Analysiere Patienten pro Tag erfolgreich durchgeführt.<br><br><br>"), add = TRUE)
    })
    
    
    
    observeEvent(input$do_inMED,{
      updateTabsetPanel(session,"main",
                        selected="Medikationsprüfungen")
      
      input6<-input5[input5$prüfdatum>=as.Date(input$MED_Start)&input5$prüfdatum<=as.Date(input$MED_End),]
      values$MED_Start<-input$MED_Start
      values$MED_End<-input$MED_End

      if(!is.null(input$med_type)&&input$med_type==F){
        values$med_type<-F
        values$med_kliniken<-input$med_kliniken
        help_klinik1<-strsplit(input$med_kliniken," (n=",fixed = T)[[1]][1]
        help_klinik1b<-strsplit(help_klinik1," ",fixed=T)[[1]][1]
        
        help_klinik2<-strsplit(input$med_kliniken," (n=",fixed = T)[[1]][2]
        help_klinik2b<-strsplit(help_klinik2," ",fixed=T)[[1]][1]
        help_klinik2b<-as.numeric(help_klinik2b)
        
        if(help_klinik2b==0){
          output$text_analyse2b_med<-renderText({paste0("Für Klinik ",help_klinik1," gibt es im gewählten Zeitraum keine Beobachtungen.")})
          output$text_analyse2c_med<-renderText({paste0("Für Klinik ",help_klinik1," gibt es im gewählten Zeitraum keine Beobachtungen.")})
          output$data_proMed<-renderDataTable({NULL})
          output$data_fb<-renderDataTable({NULL})
          output$data_proMed2<-renderDataTable({NULL})
          output$data_proMed3<-renderDataTable({NULL})
          output$barplot_proMed<-renderPlot({NULL})
          output$barplot_proMed1<-renderPlot({NULL})
          output$barplot_proMed2<-renderPlot({NULL})
          output$barplot_proMed3<-renderPlot({NULL})
          return()
        }
        
        input6<-input6[input6$fachrichtung==help_klinik1b,]
      }else{
        values$med_type<-T
        values$med_kliniken<-NA
      }
      
      
      shinyjs::html("text", paste0("<br>Führe Medikationsprüfungen durch.<br><br><br>"), add = FALSE)
      
      med_pruefung<-as.data.frame(table(input6$medikationsprüfungen.pharmazeutische.visite))
      med_pruefung$Rel<-round(100*med_pruefung$Freq/sum(med_pruefung$Freq))
      
      med_pruefung2<-med_pruefung[med_pruefung$Var1!="nicht geprüft",]
      med_pruefung2$Rel<-round(100*med_pruefung2$Freq/sum(med_pruefung2$Freq))
      label_med_pruefung2<-med_pruefung2$Var1
      label_med_pruefung2<-gsub("Anamnese und Umstellung prüfen","Anamnese und\n Umstellung prüfen",label_med_pruefung2)
      label_med_pruefung2<-gsub("Verlegung von IMESO nach MEDICO prüfen","Verlegung von IMESO\n nach MEDICO prüfen",label_med_pruefung2)
      label_med_pruefung2<-gsub("Klinikmedikation prüfen","Klinikmedikation\n prüfen",label_med_pruefung2)
      label_med_pruefung2<-gsub("Klinikmedikation ABS Visite prüfen","Klinikmedikation ABS\n Visite prüfen",label_med_pruefung2)
      label_med_pruefung2<-gsub("Entlassmedikation prüfen","Entlassmedikation\n prüfen",label_med_pruefung2)
      #ausprägungen:
      #nicht geprüft; Anamnese und Umstellung prüfen; Verlegung von IMESO nach MEDICO prüfen; Klinikmedikation prüfen
      #Klinikmedikation ABS Visite prüfen; Entlassmedikation prüfen
      
      med_abp<-as.data.frame(table(input6$abp.betreffend))
      med_abp$Rel<-round(100*med_abp$Freq/sum(med_abp$Freq))
      
      med_abp2<-med_abp[med_abp$Var1!="nicht geprüft"&med_abp$Var1!="keine ABP"&med_abp$Var1!="keine Medikation erfasst",]
      med_abp2$Rel<-round(100*med_abp2$Freq/sum(med_abp2$Freq))
      label_med_abp<-med_abp$Var1[c(which(med_abp$Var1=="nicht geprüft"),
                                   which(med_abp$Var1=="keine ABP"),
                                   which(med_abp$Var1=="keine Medikation erfasst"),
                                   which(med_abp$Var1!="nicht geprüft"&med_abp$Var1!="keine ABP"&med_abp$Var1!="keine Medikation erfasst"))]
      label_med_abp<-gsub("keine Medikation erfasst","keine Medikation\n erfasst",label_med_abp)
      label_med_abp<-gsub("Verlegung von IMESO nach MEDICO","Verlegung von\n IMESO nach MEDICO",label_med_abp)
      label_med_abp<-gsub("Umstellung auf Hausliste","Umstellung auf\n Hausliste",label_med_abp)
      label_med_abp<-gsub("Arzneimittelanamnese","Arzneimittelanamnese",label_med_abp)
      label_med_abp<-gsub("Ambulante Medikation","Ambulante Medikation",label_med_abp)
      #ausprägungen:
      #keine Medikation erfasst; keine ABP; Arzneimittelanamnese; Ambulante Medikation; Umstellung auf Hausliste
      #Verlegung von IMESO nach MEDICO; Klinikmedikation; Entlassmedikation; nicht geprüft
      
      
      med_fehler<-as.data.frame(table(input6$`fehlerfolge.(ncc.merp)`))
      med_fehler$Rel<-round(100*med_fehler$Freq/sum(med_fehler$Freq))
      med_fehler$Var1<-gsub("^A-.*","A-Fehler",med_fehler$Var1)
      med_fehler$Var1<-gsub("^B-.*","B-Fehler",med_fehler$Var1)
      med_fehler$Var1<-gsub("^C-.*","C-Fehler",med_fehler$Var1)
      med_fehler$Var1<-gsub("^D-.*","D-Fehler",med_fehler$Var1)
      med_fehler$Var1<-gsub("^E-.*","E-Fehler",med_fehler$Var1)
      med_fehler$Var1<-gsub("^F-.*","F-Fehler",med_fehler$Var1)
      med_fehler$Var1<-gsub("^G-.*","G-Fehler",med_fehler$Var1)
      med_fehler$Var1<-gsub("^H-.*","H-Fehler",med_fehler$Var1)
      med_fehler$Var1<-gsub("^I-.*","I-Fehler",med_fehler$Var1)
      med_fehler$Var1<-gsub("^W-.*","W-Fehler",med_fehler$Var1)
      
      med_fehler2<-med_fehler[med_fehler$Var1!="nicht geprüft",]
      med_fehler2$Rel<-round(100*med_fehler2$Freq/sum(med_fehler2$Freq))
      label_med_fehler<-med_fehler2$Var1
      
      
      output_proMed<-med_pruefung[c(which(med_pruefung$Var1=="nicht geprüft"),which(med_pruefung$Var1!="nicht geprüft")),]
      if(nrow(output_proMed)>0){
        if(nrow(med_pruefung2)==0){
          output_proMed<-cbind(output_proMed,NA)
        }else{
          if(length(which(output_proMed$Var1=="nicht geprüft"))>0){
            output_proMed<-cbind(output_proMed,rbind(NA,
                                                     med_pruefung2[c(which(med_pruefung2$Var1=="nicht geprüft"),which(med_pruefung2$Var1!="nicht geprüft")),])[,3])
          }else{
            output_proMed<-cbind(output_proMed,rbind(med_pruefung2[c(which(med_pruefung2$Var1=="nicht geprüft"),which(med_pruefung2$Var1!="nicht geprüft")),])[,3])
          } 
        }
      }else{
        output_proMed[1,1]<-NA
        output_proMed<-cbind(output_proMed,NA,NA)
      }
      output_proMed[,3]<-paste0(output_proMed[,3],"%")
      output_proMed[,4]<-paste0(output_proMed[,4],"%")
      output_proMed[output_proMed$Var1=="nicht geprüft",4]<-NA
      names(output_proMed)<-c("Bereiche der Medikationsprüfung","Medikationsprüfungen pro Bereich","Relativ (alle)","Relativ (nur 'geprüft')")
      output_proMed$`Relativ (alle)`[output_proMed$`Relativ (alle)`=="0%"]<-"<1%"
      output_proMed$`Relativ (nur 'geprüft')`[output_proMed$`Relativ (nur 'geprüft')`=="0%"]<-"<1%"
      output_proMed[output_proMed=="NA%"]<-NA
      
      output$data_proMed<-renderDataTable(datatable(output_proMed,rownames = F,
                                                    options = list(
                                                      columnDefs=list(list(className = 'dt-right',targets=2:3)),
                                                      language=list(search='Suchen:',lengthMenu='Zeige _MENU_ Einträge',
                                                                    info='Zeige _START_ bis _END_ von _TOTAL_ Einträgen',
                                                                    paginate=list(previous='Vorherige','next'='Nächste')
                                                    ))))
      

      
      output$barplot_proMed<-renderPlot({
        if(nrow(med_pruefung)>0){
          layout(matrix(c(1,1,2,2,3),ncol=5))
          par(mar=c(4,5,5,1))
          if(sum(med_pruefung$Var1=="nicht geprüft")>0){barplot(xlim=c(0,2.5),ylim=c(0,100),med_pruefung$Rel[med_pruefung$Var1=="nicht geprüft"],col="grey",
                                                                names.arg = "nicht geprüft",las=1,main="Medikationsprüfungen",yaxt="n",cex.names=2,cex.main=3)
          }
        if(sum(med_pruefung$Var1!="nicht geprüft")>0){
            if(sum(med_pruefung$Var1=="nicht geprüft")>0){
              barplot(xlim=c(0,2.5),ylim=c(0,100),as.matrix(med_pruefung$Rel[med_pruefung$Var1!="nicht geprüft"]),
                      add = T,space=1.5,col=c("skyblue1","orange","gold","green3","dodgerblue3","forestgreen")[1:nrow(med_pruefung2)],
                      names.arg="geprüft",yaxt="n",cex.names=2,main=3)              
            }else{
              barplot(xlim=c(0,2.5),ylim=c(0,100),as.matrix(med_pruefung$Rel[med_pruefung$Var1!="nicht geprüft"]),
                      add = F,space=1.5,col=c("skyblue1","orange","gold","green3","dodgerblue3","forestgreen")[1:nrow(med_pruefung2)],
                      names.arg="geprüft",yaxt="n",cex.names=2,main=3)
            }
          }
          
          axis(2,seq(0,100,20),labels=paste0(seq(0,100,20),"%"),las=1,cex.axis=2)
          
          if(nrow(med_pruefung2)>0){
            par(mar=c(0,1.5,5,2))
            own_pie(med_pruefung2$Freq,col = c("skyblue1","orange","gold","green3","dodgerblue3","forestgreen")[1:nrow(med_pruefung2)],
                    clockwise = T,labels = paste0(med_pruefung2$Rel,"%"),main="Pro Bereich")
            
            par(mar=c(0.5,0.5,0.5,0.5))
            plot(NULL,xlim=c(0,1),ylim=c(0,2),xaxt="n",yaxt="n",xlab="",ylab="",bty="n")
            legend("center",legend=label_med_pruefung2,cex=1.7,y.intersp = 2,
                   fill=c("skyblue1","orange","gold","green3","dodgerblue3","forestgreen")[1:nrow(med_pruefung2)])            
          }else{
            plot(NULL,xlim=c(0,1),ylim=c(0,1),xaxt="n",yaxt="n",xlab="",ylab="",bty="n",
                 main="Pro Bereich",cex.main=3)
            text(0.5,0.5,"Keine Daten vorhanden",cex=2)
            plot(NULL,xlim=c(0,1),ylim=c(0,2),xaxt="n",yaxt="n",xlab="",ylab="",bty="n")
          }
        }else{
          plot(NULL,xlim=c(0,1),ylim=c(0,1),xaxt="n",yaxt="n",xlab="",ylab="",bty="n",
               main="Medikationsprüfungen",cex.main=3)
          text(0.5,0.5,"Keine Daten vorhanden",cex=2)
        }
      })
      
      ###pro Fachrichtung

      lexikon_fach<-data.frame(Abk=c("KAIT","KHAU","KEND","KGHI","CHG","KMKG","KGYN","KHAE","KCHH","KKAR","CHK","KKJP","KNEP",
                                     "KCHN","KNEU","KNRAD","KNUK","KAUG","KORT","KHNO","KPAE","KCHP","KPNE","KPSY","KPSM","KDR",
                                     "KSTR","KCHU","KURO","KCHI","KPHO","CHT",
                                     "KCHK","PNE","KNRA"),
                               Full=c("Anästhesiologie","Dermatologie","Endokrinologie","Gastro/Hepato/Infektiologie","Gefäßchirurgie",
                                      "Gesichtschirurgie","Gynäkologie","Hämato/Onkologie","Herz/Thoraxchirurgie","Kardiologie",
                                      "Kinderchirurgie","Kinderpsychiatrie","Nephrologie","Neurochirurgie","Neurologie","Neuroradiologie",
                                      "Nuklearmedizin","Ophtalmologie","Orthopädie","Oto/Rhino/Laryngologie","Pädiatrie","Plastische Chirurgie",
                                      "Pneumologie","Psychiatrie","Psychosomatik","Radiologie","Strahlentherapie","Unfallchirurgie",
                                      "Urologie","Viszeralchirurgie","Kinder Hämato/Onkologie","Thoraxchirurgie",
                                      "Kinderchirurgie","Pneumologie","Neuroradiologie"))
      
      med_fb<-as.data.frame(table(input6$fachrichtung))
      if(nrow(med_fb)>0){
        med_fb<-med_fb[order(med_fb$Freq,decreasing = T),]
        if(input$fb_in!="Alle"){
          auswahl<-as.numeric(substr(input$fb_in,start = 5,stop = nchar(input$fb_in)))
          med_fb<-med_fb[c(1:auswahl),]
        }
        
        med_fb$Rel<-round(100*med_fb$Freq/sum(med_fb$Freq))
        
        output_fb<-med_fb
      }else{
        output_fb<-data.frame(Var1=NA,Freq=NA,Rel=NA)
      }

      output_fb$Var1<-paste0(output_fb$Var1," ",lexikon_fach$Full[match(med_fb$Var1,lexikon_fach$Abk)])
      output_fb[,3]<-paste0(output_fb[,3],"%")
      names(output_fb)<-c("Fachrichtungen","Medikationsprüfungen pro Frachrichtung","Relativ")
      output_fb$Relativ[output_fb$Relativ=="0%"]<-"<1%"
      output_fb[output_fb=="NA%"]<-NA
      
      output$data_fb<-renderDataTable(datatable(output_fb,rownames = F,
                                                     options = list(
                                                       columnDefs=list(list(className = 'dt-right',targets=2)),
                                                       language=list(search='Suchen:',lengthMenu='Zeige _MENU_ Einträge',
                                                                     info='Zeige _START_ bis _END_ von _TOTAL_ Einträgen',
                                                                     paginate=list(previous='Vorherige','next'='Nächste'))
                                                     )))
      
      
      
      output$barplot_proMed1<-renderPlot({
        par(mar=c(10,5,5,0.5))
        if(nrow(med_fb)>0){
          barplot(med_fb$Freq,names.arg =rep("",nrow(med_fb)),las=1,cex.main=2,cex.axis=1.5,
                  main="Medikationsprüfungen pro Fachrichtung",col=c("skyblue1","orange","gold","green3","dodgerblue3","forestgreen"))
          
        if(nrow(med_fb)>1){
          text(x=seq(0.75,(nrow(med_fb)*1.2),1.2),
               y=rep(-3,nrow(med_fb)),
               labels= paste0("",med_fb$Var1,"    \n",lexikon_fach$Full[match(med_fb$Var1,lexikon_fach$Abk)]),srt=60,adj=1,xpd=T)
        }else{
          text(x=0.75,
               y=-1*max(med_fb$Freq)/30,
               labels= paste0("",med_fb$Var1,"    \n",lexikon_fach$Full[match(med_fb$Var1,lexikon_fach$Abk)]),srt=0,xpd=T)
        }
        }else{
          plot(NULL,xlim=c(0,1),ylim=c(0,1),xaxt="n",yaxt="n",xlab="",ylab="",bty="n",
               main="Medikationsprüfungen",cex.main=3)
          text(0.5,0.5,"Keine Daten vorhanden",cex=2)       }
       })
      
      
      ###mit ABP - Arzneimittelbezogenes Problem
      output_proMed2<-med_abp[c(which(med_abp$Var1=="nicht geprüft"),
                                which(med_abp$Var1=="keine ABP"),
                                which(med_abp$Var1=="keine Medikation erfasst"),
                                which(med_abp$Var1!="nicht geprüft"&med_abp$Var1!="keine ABP"&med_abp$Var1!="keine Medikation erfasst")),]
      if(length(c(which(med_abp$Var1=="nicht geprüft"),which(med_abp$Var1=="keine ABP"),which(med_abp$Var1=="keine Medikation erfasst")))==3&&nrow(med_abp2)>0){
        output_proMed2<-cbind(output_proMed2,rbind(NA,NA,NA,med_abp2)[,3])
      }else{
        if(length(c(which(med_abp$Var1=="nicht geprüft"),which(med_abp$Var1=="keine ABP"),which(med_abp$Var1=="keine Medikation erfasst")))==2&&nrow(med_abp2)>0){
          output_proMed2<-cbind(output_proMed2,rbind(NA,NA,med_abp2)[,3])
        }else{
          if(length(c(which(med_abp$Var1=="nicht geprüft"),which(med_abp$Var1=="keine ABP"),which(med_abp$Var1=="keine Medikation erfasst")))==1&&nrow(med_abp2)>0){
            output_proMed2<-cbind(output_proMed2,rbind(NA,med_abp2)[,3])
          }else{
            if(nrow(output_proMed2)>0){
              output_proMed2<-cbind(output_proMed2,NA)
            }else{
              output_proMed2<-data.frame(Var1=NA,Freq=NA,Rel=NA,Rel2=NA)
            }
          }
        }
      }
      output_proMed2[,3]<-paste0(output_proMed2[,3],"%")
      output_proMed2[,4]<-paste0(output_proMed2[,4],"%")
      output_proMed2[c(which(output_proMed2$Var1=="nicht geprüft"),
                       which(output_proMed2$Var1=="keine ABP"),
                       which(output_proMed2$Var1=="keine Medikation erfasst")),4]<-NA
      names(output_proMed2)<-c("Bereiche mit ABP","ABP pro Bereich","Relativ (alle)","Relativ (nur 'mit ABP')")
      output_proMed2$`Relativ (alle)`[output_proMed2$`Relativ (alle)`=="0%"]<-"<1%"
      output_proMed2$`Relativ (nur 'mit ABP')`[output_proMed2$`Relativ (nur 'mit ABP')`=="0%"]<-"<1%"
      output_proMed2[output_proMed2=="NA%"]<-NA
      
      output$data_proMed2<-renderDataTable(datatable(output_proMed2,rownames = F,
                                                     options = list(
                                                       columnDefs=list(list(className = 'dt-right',targets=2:3)),
                                                       language=list(search='Suchen:',lengthMenu='Zeige _MENU_ Einträge',
                                                                     info='Zeige _START_ bis _END_ von _TOTAL_ Einträgen',
                                                                     paginate=list(previous='Vorherige','next'='Nächste'))
                                                     )))
      
      
      
      output$barplot_proMed2<-renderPlot({
        layout(matrix(c(1,1,2,2,3),ncol=5))
        par(mar=c(4,5,5,1))
        if(nrow(as.matrix(med_abp$Rel[c(which(med_abp$Var1=="nicht geprüft"),
                                        which(med_abp$Var1=="keine ABP"),
                                        which(med_abp$Var1=="keine Medikation erfasst"))]))>0||
           nrow(as.matrix(med_abp$Rel[med_abp$Var1!="nicht geprüft"&med_abp$Var1!="keine ABP"&med_abp$Var1!="keine Medikation erfasst"]))>0){
          if(nrow(as.matrix(med_abp$Rel[c(which(med_abp$Var1=="nicht geprüft"),
                                          which(med_abp$Var1=="keine ABP"),
                                          which(med_abp$Var1=="keine Medikation erfasst"))]))>0){
            barplot(xlim=c(0,2.5),ylim=c(0,100),as.matrix(med_abp$Rel[c(which(med_abp$Var1=="nicht geprüft"),
                                                                        which(med_abp$Var1=="keine ABP"),
                                                                        which(med_abp$Var1=="keine Medikation erfasst"))]),
                    col=c("grey90","grey60","grey30"),beside = F,
                    names.arg = "ohne ABP",las=1,main="Medikationsprüfungen",yaxt="n",cex.names=2,cex.main=3)
            if(nrow(as.matrix(med_abp$Rel[med_abp$Var1!="nicht geprüft"&med_abp$Var1!="keine ABP"&med_abp$Var1!="keine Medikation erfasst"]))>0){
              barplot(xlim=c(0,2.5),ylim=c(0,100),as.matrix(med_abp$Rel[med_abp$Var1!="nicht geprüft"&med_abp$Var1!="keine ABP"&med_abp$Var1!="keine Medikation erfasst"]),
                      add = T,space=1.5,col=c("skyblue1","orange","gold","green3","dodgerblue3","forestgreen","darkorange3","goldenrod3","darkseagreen3")[1:(nrow(med_abp)-1)],
                      names.arg="mit ABP",yaxt="n",cex.names=2,main=3)
              axis(2,seq(0,100,20),labels=paste0(seq(0,100,20),"%"),las=1,cex.axis=2)
            }else{
              barplot(xlim=c(0,2.5),ylim=c(0,100),0,
                      add = T,space=1.5,col=c("skyblue1","orange","gold","green3","dodgerblue3","forestgreen","darkorange3","goldenrod3","darkseagreen3")[1:(nrow(med_abp)-1)],
                      names.arg="mit ABP",yaxt="n",cex.names=2,main=3)
              axis(2,seq(0,100,20),labels=paste0(seq(0,100,20),"%"),las=1,cex.axis=2)
            }
          }else{
            barplot(xlim=c(0,2.5),ylim=c(0,100),as.matrix(med_abp$Rel[med_abp$Var1!="nicht geprüft"&med_abp$Var1!="keine ABP"&med_abp$Var1!="keine Medikation erfasst"]),
                    space=1.5,col=c("skyblue1","orange","gold","green3","dodgerblue3","forestgreen","darkorange3","goldenrod3","darkseagreen3")[1:(nrow(med_abp)-1)],
                    names.arg="mit ABP",main="Medikationsprüfungen",yaxt="n",cex.names=2,cex.main=3)
            axis(2,seq(0,100,20),labels=paste0(seq(0,100,20),"%"),las=1,cex.axis=2)
          }

        }else{
          plot(NULL,xlim=c(0,1),ylim=c(0,1),xaxt="n",yaxt="n",xlab="",ylab="",bty="n",
               main="Medikationsprüfungen",cex.main=3)
          text(0.5,0.5,"Keine Daten vorhanden",cex=2)
        }
        
        
        par(mar=c(0,1.5,5,2))
        if(nrow(med_abp2)>0){
          own_pie(med_abp2$Freq,col = c("skyblue1","orange","gold","green3","dodgerblue3","forestgreen","darkorange3","goldenrod3","darkseagreen3")[1:nrow(med_abp2)],
                  clockwise = T,labels = paste0(med_abp2$Rel,"%"),main="Mit ABP")
        }
        else{
          plot(NULL,xlim=c(0,1),ylim=c(0,1),xaxt="n",yaxt="n",xlab="",ylab="",bty="n",
               main="Mit ABP",cex.main=3)
          text(0.5,0.5,"Keine Daten vorhanden",cex=2)
        }
        
        par(mar=c(0.5,0.5,0.5,0.5))
        plot(NULL,xlim=c(0,1),ylim=c(0,2),xaxt="n",yaxt="n",xlab="",ylab="",bty="n")
        if(length(c(which(med_abp$Var1=="nicht geprüft"),which(med_abp$Var1=="keine ABP"),which(med_abp$Var1=="keine Medikation erfasst")))==3&&nrow(med_abp)>0){
          legend("center",legend=label_med_abp,cex=1.6,y.intersp = 1.5,
                 fill=c("grey90","grey60","grey30","skyblue1","orange","gold","green3","dodgerblue3","forestgreen","darkorange3","goldenrod3","darkseagreen3")[1:nrow(med_abp)])
        }else{
          if(length(c(which(med_abp$Var1=="nicht geprüft"),which(med_abp$Var1=="keine ABP"),which(med_abp$Var1=="keine Medikation erfasst")))==2&&nrow(med_abp)>0){
            legend("center",legend=label_med_abp,cex=1.6,y.intersp = 1.5,
                   fill=c("grey90","grey60","skyblue1","orange","gold","green3","dodgerblue3","forestgreen","darkorange3","goldenrod3","darkseagreen3")[1:nrow(med_abp)])
          }else{
            if(length(c(which(med_abp$Var1=="nicht geprüft"),which(med_abp$Var1=="keine ABP"),which(med_abp$Var1=="keine Medikation erfasst")))==1&&nrow(med_abp)>0){
              legend("center",legend=label_med_abp,cex=1.6,y.intersp = 1.5,
                     fill=c("grey90","skyblue1","orange","gold","green3","dodgerblue3","forestgreen","darkorange3","goldenrod3","darkseagreen3")[1:nrow(med_abp)])
            }
          }
        }
      })

      
      
      ###Fehlerfolgenzuordnung
      lexikon_fehler<-data.frame(Fehlerfolge=c("A-Fehler","B-Fehler","C-Fehler","D-Fehler","E-Fehler","F-Fehler",
                                               "G-Fehler","H-Fehler","I-Fehler","W-Fehler","keine Fehlerfolge","nicht geprüft"),
                                 Beschreibung=c("Umstände oder Ereignisse, die zu Fehlern\n führen können",
                                                "Fehler ist aufgetreten, hat jedoch den\n Patienten nicht erreicht",
                                                "Fehler ist aufgetreten, der den Patienten\n zwar erreicht, diesem jedoch keinen\n Schaden zugefügt hat",
                                                "Fehler, der intensiviertes Monitoring und/\n der Intervention erforderte, um\n schädliche Folgen für Patienten zu vermeiden",
                                                "Fehler, der zu vorübergehenden schädlichen\n Folgen geführt oder beigetragen hat",
                                                "Fehler, der zu vorübergehenden schädlichen\n Folgen geführt oder beigetragen hat\n und Hospitalisation oder verlängerte\n Hospitalisation erforderte",
                                                "Fehler, der zu permanentem Schaden führte\n oder dazu beigetragen hat",
                                                "Fehler, der lebensrettende Interventionen erforderte",
                                                "Fehler, der zum Tod des Patienten führte",
                                                "fehlende Wirtschaftlichkeit: unnötige Kosten",
                                                NA,NA),
                                 Schweregrad=c("ohne schädliche Folgen","ohne schädliche Folgen","ohne schädliche Folgen","potentiell schädliche Folgen","schädliche Folgen",
                                               "schädliche Folgen","schädliche Folgen","schädliche Folgen","schädliche Folgen",NA,NA,NA))
      
      output_fehler<-med_fehler[c(which(med_fehler$Var1=="nicht geprüft"),
                                  which(med_fehler$Var1!="nicht geprüft"&med_fehler$Var1!="keine Fehlerfolge"),
                                  which(med_fehler$Var1=="keine Fehlerfolge")),]
      if(sum(output_fehler$Var1=="nicht geprüft")>0){
        if(sum(med_fehler2$Var1=="nicht geprüft")>0){
          output_fehler<-cbind(output_fehler,rbind(NA,
                                                   med_fehler2[c(which(med_fehler2$Var1=="nicht geprüft"),
                                                                 which(med_fehler2$Var1!="nicht geprüft"&med_fehler2$Var1!="keine Fehlerfolge"),
                                                                 which(med_fehler2$Var1=="keine Fehlerfolge")),])[,3])
        }else{
          output_fehler<-cbind(output_fehler,NA)
        }
       }else{
        output_fehler<-cbind(output_fehler,med_fehler2[c(which(med_fehler2$Var1=="nicht geprüft"),
                                                               which(med_fehler2$Var1!="nicht geprüft"&med_fehler2$Var1!="keine Fehlerfolge"),
                                                               which(med_fehler2$Var1=="keine Fehlerfolge")),][,3])
      }
      
      if(nrow(output_fehler)==0){
        output_fehler[1,1]<-NA
      }
         output_fehler[,3]<-paste0(output_fehler[,3],"%")
      output_fehler[,4]<-paste0(output_fehler[,4],"%")
      output_fehler[output_fehler$Var1=="nicht geprüft",4]<-NA
      output_fehler$Beschreibung<-lexikon_fehler$Beschreibung[match(output_fehler$Var1,lexikon_fehler$Fehlerfolge)]
      output_fehler$Schweregrad<-lexikon_fehler$Schweregrad[match(output_fehler$Var1,lexikon_fehler$Fehlerfolge)]
      output_fehler<-output_fehler[,c(1,5,6,2,3,4)]
      names(output_fehler)<-c("Bereich/ABP","Beschreibung","Schweregrad","Prüfungen mit ABP und Fehlerfolgenzuordnung","Relativ (alle)","Relativ (nur 'geprüft')")
      output_fehler$`Relativ (alle)`[output_fehler$`Relativ (alle)`=="0%"]<-"<1%"
      output_fehler$`Relativ (nur 'geprüft')`[output_fehler$`Relativ (nur 'geprüft')`=="0%"]<-"<1%"
      output_fehler[output_fehler=="NA%"]<-NA
      
      output$data_proMed3<-renderDataTable(datatable(output_fehler,rownames = F,
                                                    options = list(
                                                      columnDefs=list(list(className = 'dt-right',targets=4:5)),
                                                      language=list(search='Suchen:',lengthMenu='Zeige _MENU_ Einträge',
                                                                    info='Zeige _START_ bis _END_ von _TOTAL_ Einträgen',
                                                                    paginate=list(previous='Vorherige','next'='Nächste'))
                                                    )))
      
      med_fehler2<-med_fehler2[c(which(med_fehler2$Var1=="keine Fehlerfolge"),which(med_fehler2$Var1!="keine Fehlerfolge")),]
      output$barplot_proMed3<-renderPlot({
        if(length(med_fehler2$Freq[med_fehler2$Rel!=0])>0){
          layout(matrix(c(1,1,1,2),ncol=4))
          par(mar=c(0,1.5,5,1.5))
          if(sum(label_med_fehler=="keine Fehlerfolge")==1){
            own_pie(med_fehler2$Freq[med_fehler2$Rel!=0],col = c("grey","skyblue1","orange","gold","green3","dodgerblue3",
                                                                 "forestgreen","darkorange3","goldenrod3","darkseagreen3","royalblue4")[1:nrow(med_fehler2[med_fehler2$Rel!=0,])],
                    clockwise = T,labels = paste0(med_fehler2$Rel,"%"),main="Prüfungen mit ABP und Anteil Fehlerfolgenzuordnung")          
          }else{
            own_pie(med_fehler2$Freq[med_fehler2$Rel!=0],col = c("skyblue1","orange","gold","green3","dodgerblue3",
                                                                 "forestgreen","darkorange3","goldenrod3","darkseagreen3","royalblue4")[1:nrow(med_fehler2[med_fehler2$Rel!=0,])],
                    clockwise = T,labels = paste0(med_fehler2$Rel,"%"),main="Prüfungen mit ABP und Anteil Fehlerfolgenzuordnung")
          }
          
          par(mar=c(0.5,0.5,0.5,0.5))
          plot(NULL,xlim=c(0,1),ylim=c(0,2),xaxt="n",yaxt="n",xlab="",ylab="",bty="n")
          if(sum(label_med_fehler=="keine Fehlerfolge")==1){
            legend("center",legend=label_med_fehler[c(which(label_med_fehler=="keine Fehlerfolge"),which(label_med_fehler!="keine Fehlerfolge"))],
                   cex=2,y.intersp = 1.5,
                   fill=c("grey","skyblue1","orange","gold","green3","dodgerblue3",
                          "forestgreen","darkorange3","goldenrod3","darkseagreen3","royalblue4")[1:length(label_med_fehler)])
          }else{
            legend("center",legend=label_med_fehler[c(which(label_med_fehler=="keine Fehlerfolge"),which(label_med_fehler!="keine Fehlerfolge"))],
                   cex=2,y.intersp = 1.5,
                   fill=c("skyblue1","orange","gold","green3","dodgerblue3",
                          "forestgreen","darkorange3","goldenrod3","darkseagreen3","royalblue4")[1:length(label_med_fehler)])
          }
          
        }else{
          plot(NULL,xlim=c(0,1),ylim=c(0,1),xaxt="n",yaxt="n",xlab="",ylab="",bty="n",
               main="Prüfungen mit ABP und Anteil Fehlerfolgenzuordnung",cex.main=3)
          text(0.5,0.5,"Keine Daten vorhanden",cex=2)
        }
        
       })
      
      output$text_analyse2b_med<-renderText({NULL})
      output$text_analyse2c_med<-renderText({NULL})
      
      shinyjs::html("text", paste0("<br>Medikationsprüfungen erfolgreich durchgeführt.<br><br><br>"), add = TRUE)
    })
    
    
    ##########################
    ##Export
    
    observe({if((!is.null(input$do_inDAY)&&input$do_inDAY!=0)&&(is.null(input$do_inMED)||input$do_inMED==0)){
      output$text_export1<-renderText({NULL})
      output$text_export2<-renderText({"Bitte führen Sie die Analyse 'Medikationsprüfung' durch."})
    }})
    observe({if((!is.null(input$do_inMED)&&input$do_inMED!=0)&&(is.null(input$do_inDAY)||input$do_inDAY==0)){
      output$text_export1<-renderText({NULL})
      output$text_export2<-renderText({"Bitte führen Sie die Analyse 'Patienten pro Tag' durch."})
    }})
    
    observe({if((!is.null(input$do_inDAY)&&input$do_inDAY!=0)&&(!is.null(input$do_inMED)&&input$do_inMED!=0)){
      shinyjs::enable("do_out2")
      #shinyFiles::shinyDirChoose(input=input,session = session,id = 'export_dir2',roots=c(home="~"))
      shinyFiles::shinyDirChoose(input=input,session = session,id = 'export_dir2',roots=c(home="C:/"))
      output$exportUI1b<-renderUI({shinyDirButton('export_dir2',title = "Export-Ordner",
                                                  label = "Durchsuchen")})
      observe(
        if(!is.null(input$export_dir2)){
          if(paste(unlist(input$export_dir2[[1]]),collapse="/")==0||paste(unlist(input$export_dir2[[1]]),collapse="/")==1){
            #output$exportUI0b<-renderUI({h5(paste0("Ordner für den Export: ~/"))})
            output$exportUI0b<-renderUI({h5(paste0("Ordner für den Export: C:/"))})
          }else{
            #output$exportUI0b<-renderUI({h5(paste0("Ordner für den Export: ~",
            #                                       paste(unlist(input$export_dir2[[1]]),collapse="/")))})
            output$exportUI0b<-renderUI({h5(paste0("Ordner für den Export: C:/",
                                                   paste(unlist(input$export_dir2[[1]]),collapse="/")))})
          }
        }else{
          #output$exportUI0b<-renderUI({h5(paste0("Ordner für den Export: ~/"))})
          output$exportUI0b<-renderUI({h5(paste0("Ordner für den Export: C:/"))})
        }      
      )
      
      
      output$text_export1<-renderText({NULL})
      output$text_export2<-renderText({NULL})
      output$text_export3<-renderText({"Hinweis"})
      output$text_export4<-renderText({"Die Analyseergebnisse werden mit der Konfiguration exportiert, in der sie zuletzt durchgeführt worden sind. 
        Bitte achten Sie darauf, dass für die 'Patienten pro Tag' und die 'Medikationsprüfung' derselbe Zeitraum gewählt wurde. 
        Achten Sie weiterhin darauf, dass beide Analysen für alle oder jeweils für dieselbe Klinik durchgeführt wurden."})
    }
    })
    
    observeEvent(input$do_out2,{
      
      updateTabsetPanel(session,"main",
                        selected="Log")
      
      if(values$DAY_Start!=values$MED_Start||values$DAY_End!=values$MED_End){
        shinyjs::html("text", paste0("<br>Unterschiedliche Zeiträume für Analysen gewählt. 
                                     Bitte führen Sie vor dem Export die Analysen für denselben Zeitraum durch."), add = FALSE)
        return()
      }

#        if((!is.null(input$day_kliniken)&&!is.null(input$med_kliniken)&&(input$day_type!=input$med_type||input$day_kliniken!=input$med_kliniken))||
#           (input$day_type!=input$med_type)){
            if(values$day_type!=values$med_type||((values$day_type==F||values$med_type==F)&&values$day_kliniken!=values$med_kliniken)){
        text_dazu1<-ifelse(values$day_type==T,"alle Kliniken",strsplit(values$day_kliniken," (n=",fixed = T)[[1]][1])
        text_dazu2<-ifelse(values$med_type==T,"alle Kliniken",strsplit(values$med_kliniken," (n=",fixed = T)[[1]][1])
        shinyjs::html("text", paste0("<br>Unterschiedliche Konfigurationen für Analysen gewählt (Patienten pro Tag: ",
        text_dazu1,"; Medikationsprüfung: ",text_dazu2,"). 
                                     Bitte führen Sie vor dem Export die Analysen für alle oder für dieselbe Klinik durch."), add = FALSE)
        
        return()
        
      } 
      
      
      shinyjs::html("text", paste0("<br>Starte Export von Daten und Analyse-Ergebnissen.<br><br>"), add = FALSE)
      
      input6<-input5[input5$prüfdatum>=as.Date(input$DAY_Start)&input5$prüfdatum<=as.Date(input$DAY_End),]
      
      if(nrow(input6)==0){
        shinyjs::html("text", paste0("<br>Im gewählten Zeitraum gibt es keine Beobachtungen und folglich keine Analyseergebnisse. 
                                       Ein Export ist daher nicht möglich."), add = FALSE)
        return()
      }
      
      if(input$day_type==F){
        help_klinik1<-strsplit(input$day_kliniken," (n=",fixed = T)[[1]][1]
        help_klinik1b<-strsplit(help_klinik1," ",fixed=T)[[1]][1]
        
        help_klinik2<-strsplit(input$day_kliniken," (n=",fixed = T)[[1]][2]
        help_klinik2b<-strsplit(help_klinik2," ",fixed=T)[[1]][1]
        help_klinik2b<-as.numeric(help_klinik2b)
        
        if(help_klinik2b==0){
          shinyjs::html("text", paste0("<br>Für Klinik ",help_klinik1," gibt es im gewählten Zeitraum keine Beobachtungen und folglich keine Analyseergebnisse. 
                                       Ein Export ist daher nicht möglich."), add = FALSE)
          return()
        }
        
        input6<-input6[input6$fachrichtung==help_klinik1b,]
      }
      
      
      
      shinyjs::html("text", paste0("<br>&nbsp;&nbsp;&nbsp;&nbsp;1. Datensatz<br>"), add = TRUE)
      if(sum(is.na(suppressWarnings(as.numeric(input6$egfr))))-sum(is.na(input6$egfr))>0)
        shinyjs::html("text", paste0("<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;",sum(is.na(suppressWarnings(as.numeric(input6$egfr))))-sum(is.na(input6$egfr)),
                                     " eGFR-Werte konnten nicht ausgewertet werden (falsches Format).<br>"), add = TRUE)
      input6$egfr<-suppressWarnings(as.numeric(input6$egfr))
      
      if(sum(is.na(suppressWarnings(as.numeric(input6$kalium))))-sum(is.na(input6$kalium))>0)
        shinyjs::html("text", paste0("<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;",sum(is.na(suppressWarnings(as.numeric(input6$kalium))))-sum(is.na(input6$kalium)),
                                     " Kalium-Werte konnten nicht ausgewertet werden (falsches Format).<br>"), add = TRUE)
      input6$kalium<-suppressWarnings(as.numeric(input6$kalium))
      
      if(sum(is.na(suppressWarnings(as.numeric(input6$natrium))))-sum(is.na(input6$natrium))>0)
        shinyjs::html("text", paste0("<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;",sum(is.na(suppressWarnings(as.numeric(input6$natrium))))-sum(is.na(input6$natrium)),
                                     " Natrium-Werte konnten nicht ausgewertet werden (falsches Format).<br>"), add = TRUE)
      input6$natrium<-suppressWarnings(as.numeric(input6$natrium))
      
      if(sum(is.na(suppressWarnings(as.numeric(input6$`antibiotika.x.tage.(abs.visite)`))))-sum(is.na(input6$`antibiotika.x.tage.(abs.visite)`))>0)
        shinyjs::html("text", paste0("<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;",sum(is.na(suppressWarnings(as.numeric(input6$`antibiotika.x.tage.(abs.visite)`))))-sum(is.na(input6$`antibiotika.x.tage.(abs.visite)`)),
                                     " Angaben zur ABS Visite konnten nicht ausgewertet werden (falsches Format).<br>"), add = TRUE)
      input6$`antibiotika.x.tage.(abs.visite)`<-suppressWarnings(as.numeric(input6$`antibiotika.x.tage.(abs.visite)`))
      
      input6_export<-input6
      names(input6_export)[names(input6_export)=="fachrichtung"]<-"Fachrichtung"
      names(input6_export)[names(input6_export)=="fallnummer"]<-"Fallnummer"
      input6_export$alter<-input6_export$alter_original
      names(input6_export)[names(input6_export)=="alter"]<-"Alter"
      
      input6_export$aufnahmezeitpunkt<-paste0(str_split_fixed(str_split_fixed(input6_export$aufnahmezeitpunkt,"-",n=Inf)[,3]," ",n=Inf)[,1],".",
                                               str_split_fixed(input6_export$aufnahmezeitpunkt,"-",n=Inf)[,2],".",
                                               str_split_fixed(input6_export$aufnahmezeitpunkt,"-",n=Inf)[,1]," ",
                                               str_split_fixed(str_split_fixed(input6_export$aufnahmezeitpunkt," ",n=Inf)[,2],":",n=Inf)[,1],":",
                                               str_split_fixed(str_split_fixed(input6_export$aufnahmezeitpunkt," ",n=Inf)[,2],":",n=Inf)[,2])
      input6_export$aufnahmezeitpunkt<-gsub(" :$","",input6_export$aufnahmezeitpunkt)
      names(input6_export)[names(input6_export)=="aufnahmezeitpunkt"]<-"Aufnahmezeitpunkt"
      
      input6_export$prüfdatum<-paste0(str_split_fixed(input6_export$prüfdatum,"-",n=Inf)[,3],".",
                                       str_split_fixed(input6_export$prüfdatum,"-",n=Inf)[,2],".",
                                       str_split_fixed(input6_export$prüfdatum,"-",n=Inf)[,1])
      names(input6_export)[names(input6_export)=="prüfdatum"]<-"Prüfdatum"
      
      if(sum(names(input6_export)=="entlasszeitpunkt")>0&&sum(is.na(input6_export$entlasszeitpunkt))!=nrow(input6_export)){
        input6_export$entlasszeitpunkt<-paste0(str_split_fixed(str_split_fixed(input6_export$entlasszeitpunkt,"-",n=Inf)[,3]," ",n=Inf)[,1],".",
                                                str_split_fixed(input6_export$entlasszeitpunkt,"-",n=Inf)[,2],".",
                                                str_split_fixed(input6_export$entlasszeitpunkt,"-",n=Inf)[,1]," ",
                                                str_split_fixed(str_split_fixed(input6_export$entlasszeitpunkt," ",n=Inf)[,2],":",n=Inf)[,1],":",
                                                str_split_fixed(str_split_fixed(input6_export$entlasszeitpunkt," ",n=Inf)[,2],":",n=Inf)[,2])
        input6_export$entlasszeitpunkt<-gsub(" :$","",input6_export$entlasszeitpunkt)
        input6_export$entlasszeitpunkt<-gsub("..NA","",input6_export$entlasszeitpunkt,fixed=T)
        input6_export$entlasszeitpunkt[input6_export$entlasszeitpunkt==""]<-NA
        names(input6_export)[names(input6_export)=="entlasszeitpunkt"]<-"Entlasszeitpunkt"
      }

      input6_export$egfr<-paste(input6_export$egfr,input6_export$dialyse,sep=" ")
      input6_export$egfr<-gsub("NA NA",NA,input6_export$egfr)
      input6_export$egfr<-gsub(" NA","",input6_export$egfr)
      input6_export$egfr<-gsub(".",",",input6_export$egfr,fixed=T)
      names(input6_export)[names(input6_export)=="egfr"]<-paste0("eGFR <",input$DAY_egfr,"\nml/min/1,73m²")
      
      input6_export$kalium<-gsub(".",",",input6_export$kalium,fixed=T)      
      names(input6_export)[names(input6_export)=="kalium"]<-paste0("K <",input$DAY_k[1],"oder\n>",input$DAY_k[2]," mmol/l")
      input6_export$natrium<-gsub(".",",",input6_export$natrium,fixed=T)
      names(input6_export)[names(input6_export)=="natrium"]<-paste0("Na <",input$DAY_na[1],"oder\n>",input$DAY_na[2]," mmol/l")

      names(input6_export)[names(input6_export)=="antibiotika.x.tage.(abs.visite)"]<-"Antibiotika x Tage\n(ABS Visite)"
      names(input6_export)[names(input6_export)=="medikationsprüfungen.pharmazeutische.visite"]<-"Medikationsprüfungen\npharmazeutische Visite"
      names(input6_export)[names(input6_export)=="abp.betreffend"]<-"ABP betreffend"
      names(input6_export)[names(input6_export)=="fehlerfolge.(ncc.merp)"]<-"Fehlerfolge (NCC MERP)"
      names(input6_export)[names(input6_export)=="korrigierende.maßnahmen"]<-"Korrigierende Maßnahmen"
      input6_export<-input6_export[,which(names(input6_export)!="dialyse"&names(input6_export)!="alter_original")]
      names(input6_export)[names(input6_export)=="outcome.problem"]<-"Outcome Problem"
      names(input6_export)[names(input6_export)=="präventivmaßnahmen"]<-"Präventivmaßnahmen"
      names(input6_export)[names(input6_export)=="triple.whammy.=.potentielle.nierenschädigung"]<-"Triple Whammy =\npotentielle Nierenschädigung"
      
      names(input6_export)[names(input6_export)=="antibiotika.bei.egfr<30ml/min"]<-"Antibiotika bei\neGFR<30ml/min"
      names(input6_export)[names(input6_export)=="fehlerart"]<-"Fehlerart"
      names(input6_export)[names(input6_export)=="status.des.abp"]<-"Status des ABP"
      names(input6_export)[names(input6_export)=="child-pugh.score.c"]<-"Child-Pugh Score C"
      names(input6_export)[names(input6_export)=="fehlerart.medikation.(angelehnt.an.pcne)"]<-"Fehlerart Medikation\n(angelehnt an PCNE)"

      for(name in 1:ncol(input6_export)){
        if(names(input6_export)[name]%nin%c("Fachrichtung","Fallnummer","Alter","Aufnahmezeitpunkt","Entlasszeitpunkt","Prüfdatum",
                                             "eGFR außerhalb Normbereich","K außerhalb Normbereich","Na außerhalb Normbereich",
                                             "Antibiotika x Tage\n(ABS Visite)","Medikationsprüfungen\npharmazeutische Visite",
                                             "ABP betreffend","Fehlerfolge (NCC MERP)","Korrigierende Maßnahmen",
                                             "Outcome Problem","Präventivmaßnahmen","Triple Whammy =\npotentielle Nierenschädigung",
                                             "Antibiotika bei\neGFR<30ml/min","Fehlerart","Status des ABP","Child-Pugh Score C",
                                             "Fehlerart Medikation\n(angelehnt an PCNE)")){
          names(input6_export)[name]<-capitalize(names(input6_export)[name])
          
        }
      }
      
      wb<-createWorkbook()
      addWorksheet(wb,sheetName="Datensatz")
      writeData(wb,"Datensatz",input6_export,rowNames = F)
      bold<-createStyle(textDecoration = "bold",wrapText=T)
      addStyle(wb,"Datensatz",bold,rows = 1,cols=c(1:ncol(input6_export)))
      
      width_vec1<-apply(input6_export,2,function(x){max(nchar(as.character(x)) + 2,1, na.rm = TRUE)})
      width_vec2<-apply(str_split_fixed(names(input6_export),pattern = "\n",Inf),1,function(x){max(nchar(x))})+2
      
      #width_vec2<-nchar(names(input6_export)) + 2
      width_vec<-apply(cbind(width_vec1,width_vec2),1,max)
      width_vec[width_vec>50]<-50
      setColWidths(wb,"Datensatz",cols = c(1:ncol(input6_export)),widths = width_vec)

      
      
      shinyjs::html("text", paste0("<br><br>&nbsp;&nbsp;&nbsp;&nbsp;2. Analyse: Patienten pro Tag<br>"), add = TRUE)
      
      output_proTag<-data.frame(Datum=as.Date(as.Date(input$DAY_Start):as.Date(input$DAY_End)),
                                Patienten.Tag=0,
                                Patienten.groesser.gleich.80.Jahre=0,
                                egfr.ohne.Dialyse=0,
                                k=0,
                                na=0,
                                Patienten.fuer.Medikationsanalyse=0,
                                Patienten.fuer.ABS.Visite=0,
                                Datum_intern=as.Date(as.Date(input$DAY_Start):as.Date(input$DAY_End)))
      
      wochentage<-c("Sonntag","Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag")
      
      helper<-str_split_fixed(output_proTag$Datum,pattern = "-",Inf)
      helper2<-paste0(helper[,3],".",helper[,2],".",helper[,1])
      output_proTag$Datum<-paste0(wochentage[1+as.POSIXlt(output_proTag$Datum)$wday],", ",helper2)
      ##tagesdarstellung optimieren
      
      for(i in 1:length(output_proTag$Datum_intern)){
        output_proTag$Patienten.Tag[i]<-sum(input6$prüfdatum==output_proTag$Datum_intern[i])
        output_proTag$Patienten.groesser.gleich.80.Jahre[i]<-sum(input6$prüfdatum==output_proTag$Datum_intern[i]&input6$alter>=80)
        output_proTag$Patienten.fuer.Medikationsanalyse[i]<-sum(input6$prüfdatum==output_proTag$Datum_intern[i]&input6$medikationsprüfungen.pharmazeutische.visite!="nicht geprüft")
        output_proTag$egfr.ohne.Dialyse[i]<-sum(input6$prüfdatum==output_proTag$Datum_intern[i]&is.na(input6$dialyse)&!is.na(input6$egfr)&input6$egfr<input$DAY_egfr)
        output_proTag$k[i]<-sum(input6$prüfdatum==output_proTag$Datum_intern[i]&is.na(input6$dialyse)&!is.na(input6$k)&(input6$k<input$DAY_k[1]|input6$k>input$DAY_k[2]))
        output_proTag$na[i]<-sum(input6$prüfdatum==output_proTag$Datum_intern[i]&is.na(input6$dialyse)&!is.na(input6$na)&(input6$k<input$DAY_na[1]|input6$k>input$DAY_na[2]))
        output_proTag$Patienten.fuer.ABS.Visite[i]<-sum(input6$prüfdatum==output_proTag$Datum_intern[i]&!is.na(input6$`antibiotika.x.tage.(abs.visite)`)&input6$`antibiotika.x.tage.(abs.visite)`>0)
        
      }
      output_proTag<-rbind(output_proTag,rep(NA,9),rep(NA,9))
      output_proTag[(nrow(output_proTag)-1),1]<-"Gesamtsumme"
      output_proTag[(nrow(output_proTag)-1),2:(ncol(output_proTag)-1)]<-colSums(output_proTag[1:(nrow(output_proTag)-2),2:(ncol(output_proTag)-1)])
      
      output_proTag[nrow(output_proTag),1]<-"Mittelwert pro Arbeitstag (ohne 0)"
      output_proTag[nrow(output_proTag),2:(ncol(output_proTag)-1)]<-apply(output_proTag[1:(nrow(output_proTag)-2),2:(ncol(output_proTag)-1)],MARGIN = 2,FUN = function(x){
        return(round(mean(x[x!=0])))
      })
      
      output_proTag2<-output_proTag[,1:8]
      names(output_proTag2)[2]<-"Patienten/Tag"
      names(output_proTag2)[3]<-"Patienten >=80\nJahre"
      names(output_proTag2)[4]<-paste0("Patienten eGFR <",input$DAY_egfr,"\nml/min ohne Dialyse")
      names(output_proTag2)[5]<-paste0("Patienten mit K <",input$DAY_k[1],"\noder >",input$DAY_k[2]," mmol/l")
      names(output_proTag2)[6]<-paste0("Patienten mit Na <",input$DAY_na[1],"\noder >",input$DAY_na[2]," mmol/l")
      names(output_proTag2)[7]<-paste0("Patienten für\nMedikationsanalyse")
      names(output_proTag2)[8]<-paste0("Patienten für\nABS Visite")
      #output_proTag2$Datum<-gsub(",",",\n",output_proTag2$Datum)
      #output_proTag2[nrow(output_proTag2),1]<-"Mittelwert pro\nArbeitstag (ohne 0)"
      output_proTag2[nrow(output_proTag2),as.character(output_proTag2[nrow(output_proTag2),])=="NaN"]<-0
      
      
      addWorksheet(wb,sheetName="Patienten pro Tag")
      writeData(wb,"Patienten pro Tag",output_proTag2)
      addStyle(wb,"Patienten pro Tag",bold,rows = 1,cols=c(1:ncol(output_proTag2)))
      width_vec1<-apply(output_proTag2,2,function(x){
        y<-str_split_fixed(x,"\n",Inf)
        return(max(nchar(y) + 2,1, na.rm = TRUE))
        })
      width_vec2<-apply(str_split_fixed(names(output_proTag2),pattern = "\n",Inf),1,function(x){max(nchar(x))})+2
      #width_vec<-apply(cbind(width_vec1,width_vec2),1,max)
      
      #setColWidths(wb,"Patienten pro Tag",cols = c(1:ncol(output_proTag2)),widths = width_vec)
      linie<-createStyle(border = "top")
      addStyle(wb,"Patienten pro Tag",linie,cols=c(1:ncol(output_proTag2)),rows=(nrow(output_proTag2)))
      
      
      
      ##eGFR Sonderanalyse
      egfr<-data.frame(Kategorie=c("eGFR >= 90ml/min","60ml/min <= eGFR < 90ml/min",
                                   "30ml/min <= eGFR < 60ml/min","15ml/min <= eGFR < 30ml/min",
                                   "eGFR < 15ml/min","Dialyse","keine Angabe"),
                       Abs=c(sum(!is.na(input6$egfr)&input6$egfr>=90&is.na(input6$dialyse)),
                             sum(!is.na(input6$egfr)&input6$egfr<90&input6$egfr>=60&is.na(input6$dialyse)),
                             sum(!is.na(input6$egfr)&input6$egfr<60&input6$egfr>=30&is.na(input6$dialyse)),
                             sum(!is.na(input6$egfr)&input6$egfr<30&input6$egfr>=15&is.na(input6$dialyse)),
                             sum(!is.na(input6$egfr)&input6$egfr<15&is.na(input6$dialyse)),
                             sum(!is.na(input6$dialyse)),
                             sum(is.na(input6$egfr&is.na(input6$dialyse))))
      )
      egfr$Rel1<-paste0(round(100*egfr$Abs/sum(egfr$Abs[1:6])),"%")
      egfr$Rel1[egfr$Rel1=="0%"&egfr$Fälle!=0]<-"<1%"
      egfr$Rel1[7]<-NA
      egfr$Rel1[egfr$Rel1=="NaN%"]<-"0%"
      names(egfr)<-c("Kategorie","Fälle","Relativ")
      
      writeData(wb,"Patienten pro Tag",egfr,startRow = max((nrow(output_proTag2)+5),62))
      addStyle(wb,"Patienten pro Tag",bold,rows = max((nrow(output_proTag2)+5),62),cols=c(1:ncol(egfr)))
      align<-createStyle(halign = "right")
      addStyle(wb,"Patienten pro Tag",align,
               rows = c((max((nrow(output_proTag2)+5),62)+1):(max((nrow(output_proTag2)+5),62)+8)),cols=3,stack=T)
      
      width_vec3<-apply(egfr,2,function(x){
        return(max(nchar(x) + 2,1, na.rm = TRUE))
      })
      width_vec3<-c(width_vec3,0,0,0,0,0)
      width_vec<-apply(cbind(width_vec1,width_vec2,width_vec3),1,max)
      
      setColWidths(wb,"Patienten pro Tag",cols = c(1:ncol(output_proTag2)),widths = width_vec)
      linie<-createStyle(border = "top")
      addStyle(wb,"Patienten pro Tag",linie,cols=c(1:ncol(output_proTag2)),rows=(nrow(output_proTag2)))
      
      
      ######################
      ##barplot 1
      auswahl<-input$barplot1
      auswahl<-gsub("Patienten eGFR unter Grenzwert ohne Dialyse",
                    paste0("Patienten eGFR <",input$DAY_egfr," ml/min ohne Dialyse"),auswahl)
      auswahl<-gsub("Patienten mit K außerhalb Grenzbereich",
                    paste0("Patienten mit K <",input$DAY_k[1]," oder >",input$DAY_k[2]," mmol/l"),auswahl)
      auswahl<-gsub("Patienten mit Na außerhalb Grenzbereich",
                    paste0("Patienten mit Na <",input$DAY_na[1]," oder >",input$DAY_na[2]," mmol/l"),auswahl)
      
      auswahl2<-input$barplot1
      auswahl2<-gsub("Patienten/Tag","Patienten.Tag",auswahl2)
      auswahl2<-gsub("Patienten >=80 Jahre","Patienten.groesser.gleich.80.Jahre",auswahl2)
      auswahl2<-gsub("Patienten eGFR unter Grenzwert ohne Dialyse","egfr.ohne.Dialyse",auswahl2)
      auswahl2<-gsub("Patienten mit K außerhalb Grenzbereich","k",auswahl2)
      auswahl2<-gsub("Patienten mit Na außerhalb Grenzbereich","na",auswahl2)
      auswahl2<-gsub("Patienten für Medikationsanalyse","Patienten.fuer.Medikationsanalyse",auswahl2)
      auswahl2<-gsub("Patienten für ABS Visite","Patienten.fuer.ABS.Visite",auswahl2)
      
      auswahl3<-auswahl
      auswahl3<-gsub("Patienten","",auswahl3)
      auswahl3<-gsub("/Tag","pro Tag",auswahl3,fixed=T)
      auswahl3<-gsub("^ ","",auswahl3)
      
      png(paste0(tempdir(), "/", "plot1.png"), width=3400, height=1400,units = "px")
        par(mar=c(18,6,5,1))
        if(sum(t(as.matrix(output_proTag[1:(nrow(output_proTag)-2),auswahl2])),na.rm=T)>0){
          if(length(auswahl)==3){
            barplot(t(as.matrix(output_proTag[1:(nrow(output_proTag)-2),auswahl2])),
                    beside=T,las=2,col=c("grey","skyblue","orange"),xaxt="n",
                    ylim=c(0,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1),
                    main=paste0("Anzahl Patienten ",paste(auswahl3,collapse = ", ")),cex.main=4,cex.axis=2,cex.lab=2)
            text(x=seq(2.5,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                 y=rep(-1*(max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1)/55,(nrow(output_proTag)-2)),cex=2,
                 labels=output_proTag$Datum[1:(nrow(output_proTag)-2)],srt=60,adj=1,xpd=T)
          }else{
            if(length(auswahl)==2){
              barplot(t(as.matrix(output_proTag[1:(nrow(output_proTag)-2),auswahl2])),
                      beside=T,las=2,col=c("grey","skyblue"),xaxt="n",
                      ylim=c(0,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1),
                      main=paste0("Anzahl Patienten ",paste(auswahl3,collapse = ", ")),cex.main=4,cex.axis=2,cex.lab=2)
              text(x=seq(2,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                   y=rep(-1*(max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1)/55,(nrow(output_proTag)-2)),cex=2,
                   labels=output_proTag$Datum[1:(nrow(output_proTag)-2)],srt=60,adj=1,xpd=T)
            }else{
              if(length(auswahl)==1){
                barplot(t(as.matrix(output_proTag[1:(nrow(output_proTag)-2),auswahl2])),
                        beside=T,las=2,col=c("grey"),xaxt="n",
                        ylim=c(0,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1),
                        main=paste0("Anzahl Patienten ",paste(auswahl3,collapse = ", ")),cex.main=4,cex.axis=2,cex.lab=2)
                text(x=seq(1.5,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                     y=rep(-1*(max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1)/55,(nrow(output_proTag)-2)),cex=2,
                     labels=output_proTag$Datum[1:(nrow(output_proTag)-2)],srt=60,adj=1,xpd=T)
              }
            }
          }
          
          for(i in 1:length(auswahl2)){
            if(nrow(output_proTag)<=6){
              text(x=seq(0.5+i,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                   y=output_proTag[1:(nrow(output_proTag)-2),auswahl2[i]]+max(0.4,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1*0.04),
                   output_proTag[1:(nrow(output_proTag)-2),auswahl2[i]],cex=4)
            }else{
              text(x=seq(0.5+i,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                   y=output_proTag[1:(nrow(output_proTag)-2),auswahl2[i]]+max(0.4,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1*0.04),
                   output_proTag[1:(nrow(output_proTag)-2),auswahl2[i]],cex=2.5/(nrow(output_proTag)-2)*15)
            }
          }
          
          if((max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1)>50){
            for(i in seq(50,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1,50)){
              abline(h=i,col="gray80")
            }
          }
          
          legend("topright",legend=auswahl[1],bty="n",fill="grey",cex=3)
          if(length(auswahl)>1)
            legend("top",legend=auswahl[2],bty="n",fill="skyblue",cex=3)
          if(length(auswahl)>2)
            legend("topleft",legend=auswahl[3],bty="n",fill="orange",cex=3)
        
        }else{
          if(length(auswahl)>0){
            if(length(auswahl)==3){
              barplot(t(as.matrix(output_proTag[1:(nrow(output_proTag)-2),auswahl2])),
                      beside=T,las=2,col=c("grey","skyblue","orange"),xaxt="n",
                      ylim=c(0,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1),
                      main=paste0("Anzahl Patienten ",paste(auswahl3,collapse = ", ")),cex.main=4,cex.axis=2,cex.lab=2)
              text(x=seq(2.5,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                   y=rep(-1*(max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1)/55,(nrow(output_proTag)-2)),cex=2,
                   labels=output_proTag$Datum[1:(nrow(output_proTag)-2)],srt=60,adj=1,xpd=T)
            }else{
              if(length(auswahl)==2){
                barplot(t(as.matrix(output_proTag[1:(nrow(output_proTag)-2),auswahl2])),
                        beside=T,las=2,col=c("grey","skyblue"),xaxt="n",
                        ylim=c(0,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1),
                        main=paste0("Anzahl Patienten ",paste(auswahl3,collapse = ", ")),cex.main=4,cex.axis=2,cex.lab=2)
                text(x=seq(2,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                     y=rep(-1*(max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1)/55,(nrow(output_proTag)-2)),cex=2,
                     labels=output_proTag$Datum[1:(nrow(output_proTag)-2)],srt=60,adj=1,xpd=T)
              }else{
                if(length(auswahl)==1){
                  barplot(t(as.matrix(output_proTag[1:(nrow(output_proTag)-2),auswahl2])),
                          beside=T,las=2,col=c("grey"),xaxt="n",
                          ylim=c(0,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1),
                          main=paste0("Anzahl Patienten ",paste(auswahl3,collapse = ", ")),cex.main=4,cex.axis=2,cex.lab=2)
                  text(x=seq(1.5,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                       y=rep(-1*(max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1)/55,(nrow(output_proTag)-2)),cex=2,
                       labels=output_proTag$Datum[1:(nrow(output_proTag)-2)],srt=60,adj=1,xpd=T)
                }
              }
            }
            
            for(i in 1:length(auswahl2)){
              if(nrow(output_proTag)<=6){
                text(x=seq(0.5+i,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                     y=output_proTag[1:(nrow(output_proTag)-2),auswahl2[i]]+max(0.4,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1*0.04),
                     output_proTag[1:(nrow(output_proTag)-2),auswahl2[i]],cex=4)
              }else{
                text(x=seq(0.5+i,(nrow(output_proTag)-2)*(length(auswahl2)+1),(length(auswahl2)+1)),
                     y=output_proTag[1:(nrow(output_proTag)-2),auswahl2[i]]+max(0.4,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1*0.04),
                     output_proTag[1:(nrow(output_proTag)-2),auswahl2[i]],cex=2.5/(nrow(output_proTag)-2)*15)
              }
            }
            
            if((max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1)>50){
              for(i in seq(50,max(round(max(output_proTag[1:(nrow(output_proTag)-2),auswahl2],na.rm=T),-1),5)*1.1,50)){
                abline(h=i,col="gray80")
              }
            }
            
            legend("topright",legend=auswahl[1],bty="n",fill="grey",cex=3)
            if(length(auswahl)>1)
              legend("top",legend=auswahl[2],bty="n",fill="skyblue",cex=3)
            if(length(auswahl)>2)
              legend("topleft",legend=auswahl[3],bty="n",fill="orange",cex=3)
          }else{
            plot(NULL,xlim=c(0,1),ylim=c(0,1),xaxt="n",yaxt="n",xlab="",ylab="",bty="n",
                 main=paste0("Anzahl Patienten ",paste(auswahl3,collapse = ", ")),cex.main=1.9)
            text(0.5,0.5,"Keine Daten vorhanden",cex=2) 
          }
        }
        
      dev.off()
      insertImage(wb, sheet="Patienten pro Tag", paste0(tempdir(), "/", "plot1.png"), width = 17,height = 7,
                  startRow = 1,startCol = ncol(output_proTag2)+2)
      
      
      ######################
      ##barplot 2
      auswahl_b<-input$barplot2
      auswahl_b<-gsub("Patienten eGFR unter Grenzwert ohne Dialyse",
                      paste0("Patienten eGFR <",input$DAY_egfr," ml/min\n ohne Dialyse"),auswahl_b)
      auswahl_b<-gsub("Patienten mit K außerhalb Grenzbereich",
                      paste0("Patienten mit\n K <",input$DAY_k[1]," oder >",input$DAY_k[2]," mmol/l"),auswahl_b)
      auswahl_b<-gsub("Patienten mit Na außerhalb Grenzbereich",
                      paste0("Patienten mit\n Na <",input$DAY_na[1]," oder >",input$DAY_na[2]," mmol/l"),auswahl_b)
      auswahl_b<-gsub("Patienten für Medikationsanalyse",
                      paste0("Patienten für\n Medikationsanalyse"),auswahl_b)
      auswahl_b<-gsub("Patienten für ABS Visite",
                      paste0("Patienten für\n ABS Visite"),auswahl_b)
      
      auswahl2b<-input$barplot2
      auswahl2b<-gsub("Patienten/Tag","Patienten.Tag",auswahl2b)
      auswahl2b<-gsub("Patienten >=80 Jahre","Patienten.groesser.gleich.80.Jahre",auswahl2b)
      auswahl2b<-gsub("Patienten eGFR unter Grenzwert ohne Dialyse","egfr.ohne.Dialyse",auswahl2b)
      auswahl2b<-gsub("Patienten mit K außerhalb Grenzbereich","k",auswahl2b)
      auswahl2b<-gsub("Patienten mit Na außerhalb Grenzbereich","na",auswahl2b)
      auswahl2b<-gsub("Patienten für Medikationsanalyse","Patienten.fuer.Medikationsanalyse",auswahl2b)
      auswahl2b<-gsub("Patienten für ABS Visite","Patienten.fuer.ABS.Visite",auswahl2b)
      
      auswahl3b<-auswahl_b
      auswahl3b<-gsub("Patienten","",auswahl3b)
      auswahl3b<-gsub("/Tag","pro Tag",auswahl3b,fixed=T)
      auswahl3b<-gsub("^ ","",auswahl3b)
      
      png(paste0(tempdir(), "/", "plot2.png"), width=1700, height=400,units = "px")
        par(mar=c(4,5,6,1))
        if(sum(t(as.matrix(output_proTag[nrow(output_proTag)-1,auswahl2b])))>0){
          layout(matrix(data = c(1,1,2,2,3),ncol = 5))
          barplot(t(as.matrix(output_proTag[nrow(output_proTag)-1,auswahl2b])),
                  beside=T,las=1,col=c("#3E7C90","#E69EBB","#EC952E","#F0D7E2","#9EBC9F","#E7EFAF","#6A6BE5","#EEED39"[1:length(auswahl_b)]),
                  ylim=c(0,max(6,round(max(output_proTag[nrow(output_proTag)-1,auswahl2b],na.rm=T),-1)*1.2)),
                  #names.arg=auswahl_b,
                  main=paste0("Gesamtsumme \n",
                              paste(strsplit(as.character(input$DAY_Start),"-")[[1]][c(3,2,1)],collapse = "."),
                              " bis ",
                              paste(strsplit(as.character(input$DAY_End),"-")[[1]][c(3,2,1)],collapse = ".")),
                  cex.main=2,cex.axis=1.5,names.arg="")
          abline(h=0)
          
          sd<-apply(as.data.frame(output_proTag[1:(nrow(output_proTag)-2),auswahl2b]),2,sd)
          sd[is.na(sd)]<-0
          
          barplot(t(as.matrix(output_proTag[nrow(output_proTag),auswahl2b])),
                  beside=T,las=1,col=c("#3E7C90","#E69EBB","#EC952E","#F0D7E2","#9EBC9F","#E7EFAF","#6A6BE5","#EEED39"[1:length(auswahl_b)]),
                  ylim=c(0,max(6,round(max(output_proTag[nrow(output_proTag),auswahl2b]+sd,na.rm = T),-1)*1.2)),xaxt="n",
                  #names.arg=auswahl_b,
                  main=paste0("Mittelwert pro Arbeitstag (ohne 0) \n",
                              paste(strsplit(as.character(input$DAY_Start),"-")[[1]][c(3,2,1)],collapse = "."),
                              " bis ",
                              paste(strsplit(as.character(input$DAY_End),"-")[[1]][c(3,2,1)],collapse = ".")),
                  cex.main=2,cex.axis=1.5)
          abline(h=0)
          
          for(i in 1:length(sd)){
            points(x=c(i+0.5,i+0.5),y=c(output_proTag[nrow(output_proTag),auswahl2b[i]]-sd[i],output_proTag[nrow(output_proTag),auswahl2b[i]]+sd[i]),
                   type="l",lwd=4)
          }
          par(mar=c(0.5,0.5,0.5,0.5))
          plot(NULL,xlim=c(0,1),ylim=c(0,2),xaxt="n",yaxt="n",xlab="",ylab="",bty="n")
          legend("center",legend=auswahl_b,cex=1.7,y.intersp = 2,
                 fill=c("#3E7C90","#E69EBB","#EC952E","#F0D7E2","#9EBC9F","#E7EFAF","#6A6BE5","#EEED39"[1:length(auswahl_b)]))
          
        }else{
          plot(NULL,xlim=c(0,1),ylim=c(0,1),xaxt="n",yaxt="n",xlab="",ylab="",bty="n",
               main=paste0("Gesamtsumme \n",
                           paste(strsplit(as.character(input$DAY_Start),"-")[[1]][c(3,2,1)],collapse = "."),
                           " bis ",
                           paste(strsplit(as.character(input$DAY_End),"-")[[1]][c(3,2,1)],collapse = ".")),cex.main=1.9)
          text(0.5,0.5,"Keine Daten vorhanden",cex=2)
        }
        
        dev.off()
        
        insertImage(wb, sheet="Patienten pro Tag", paste0(tempdir(), "/", "plot2.png"), width = 17,height = 4,
                    startRow = 39,startCol = ncol(output_proTag2)+2)
      
        
        
        ##barplot 3
        png(paste0(tempdir(), "/", "plot2b.png"), width=800, height=600,units = "px")
          par(mar=c(12,5,5,5.5))
          if(sum(egfr$Fälle[1:6])>0){
            barplot(egfr$Fälle[1:6],las=2,col=c(colorRampPalette(c("skyblue","orange"))(5),"grey"),xaxt="n",
                    ylim=c(0,round(max(egfr$Fälle[1:6],na.rm=T))*1.1),
                    main=paste0("Verteilung eGFR"),cex.main=1.9,ylab="Absolut",cex.lab=2,cex.axis=1.3)
            add_helper<-seq(0,round(max(egfr$Fälle[1:6],na.rm=T))*1.1,sum(egfr$Fälle[1:6])/10)
            axis(4,at=add_helper,
                 labels = paste0(seq(0,10*(length(add_helper)-1),10),"%"),las=2,cex.axis=1.4)
            mtext("Relativ",side = 4,line = 4.5,cex=2)
            text(x=seq(0.7,7,1.2),
                 y=rep(-1*(round(max(egfr$Fälle[1:6],na.rm=T))*1.1)/18,6),cex=1,
                 labels=egfr$Kategorie[1:6],srt=60,adj=1,xpd=T)
          }else{
            plot(NULL,xlim=c(0,1),ylim=c(0,1),xaxt="n",yaxt="n",xlab="",ylab="",bty="n",
                 main=paste0("Verteilung eGFR"),cex.main=1.9)
            text(0.5,0.5,"Keine Daten vorhanden",cex=2)
          }
          dev.off()
          
          insertImage(wb, sheet="Patienten pro Tag", paste0(tempdir(), "/", "plot2b.png"), width = 8,height = 6,
                      startRow =  max(nrow(output_proTag2),(39+20))+3,startCol = ncol(output_proTag2)+2)
        
      
        
        
        #################
        shinyjs::html("text", paste0("<br><br>&nbsp;&nbsp;&nbsp;&nbsp;3. Analyse: Medikationsprüfungen<br>"), add = TRUE)
        
        input6<-input5[input5$prüfdatum>=as.Date(input$MED_Start)&input5$prüfdatum<=as.Date(input$MED_End),]
        
        med_pruefung<-as.data.frame(table(input6$medikationsprüfungen.pharmazeutische.visite))
        med_pruefung$Rel<-round(100*med_pruefung$Freq/sum(med_pruefung$Freq))
        
        med_pruefung2<-med_pruefung[med_pruefung$Var1!="nicht geprüft",]
        med_pruefung2$Rel<-round(100*med_pruefung2$Freq/sum(med_pruefung2$Freq))
        label_med_pruefung2<-med_pruefung2$Var1
        label_med_pruefung2<-gsub("Anamnese und Umstellung prüfen","Anamnese und\n Umstellung prüfen",label_med_pruefung2)
        label_med_pruefung2<-gsub("Verlegung von IMESO nach MEDICO prüfen","Verlegung von IMESO\n nach MEDICO prüfen",label_med_pruefung2)
        label_med_pruefung2<-gsub("Klinikmedikation prüfen","Klinikmedikation\n prüfen",label_med_pruefung2)
        label_med_pruefung2<-gsub("Klinikmedikation ABS Visite prüfen","Klinikmedikation ABS\n Visite prüfen",label_med_pruefung2)
        label_med_pruefung2<-gsub("Entlassmedikation prüfen","Entlassmedikation\n prüfen",label_med_pruefung2)
        #ausprägungen:
        #nicht geprüft; Anamnese und Umstellung prüfen; Verlegung von IMESO nach MEDICO prüfen; Klinikmedikation prüfen
        #Klinikmedikation ABS Visite prüfen; Entlassmedikation prüfen
        
        med_abp<-as.data.frame(table(input6$abp.betreffend))
        med_abp$Rel<-round(100*med_abp$Freq/sum(med_abp$Freq))
        
        med_abp2<-med_abp[med_abp$Var1!="nicht geprüft"&med_abp$Var1!="keine ABP"&med_abp$Var1!="keine Medikation erfasst",]
        med_abp2$Rel<-round(100*med_abp2$Freq/sum(med_abp2$Freq))
        label_med_abp<-med_abp$Var1[c(which(med_abp$Var1=="nicht geprüft"),
                                      which(med_abp$Var1=="keine ABP"),
                                      which(med_abp$Var1=="keine Medikation erfasst"),
                                      which(med_abp$Var1!="nicht geprüft"&med_abp$Var1!="keine ABP"&med_abp$Var1!="keine Medikation erfasst"))]
        label_med_abp<-gsub("keine Medikation erfasst","keine Medikation\n erfasst",label_med_abp)
        label_med_abp<-gsub("Verlegung von IMESO nach MEDICO","Verlegung von\n IMESO nach MEDICO",label_med_abp)
        label_med_abp<-gsub("Umstellung auf Hausliste","Umstellung auf\n Hausliste",label_med_abp)
        label_med_abp<-gsub("Arzneimittelanamnese","Arzneimittelanamnese",label_med_abp)
        label_med_abp<-gsub("Ambulante Medikation","Ambulante Medikation",label_med_abp)
        #ausprägungen:
        #keine Medikation erfasst; keine ABP; Arzneimittelanamnese; Ambulante Medikation; Umstellung auf Hausliste
        #Verlegung von IMESO nach MEDICO; Klinikmedikation; Entlassmedikation; nicht geprüft
        
        
        med_fehler<-as.data.frame(table(input6$`fehlerfolge.(ncc.merp)`))
        med_fehler$Rel<-round(100*med_fehler$Freq/sum(med_fehler$Freq))
        med_fehler$Var1<-gsub("^A-.*","A-Fehler",med_fehler$Var1)
        med_fehler$Var1<-gsub("^B-.*","B-Fehler",med_fehler$Var1)
        med_fehler$Var1<-gsub("^C-.*","C-Fehler",med_fehler$Var1)
        med_fehler$Var1<-gsub("^D-.*","D-Fehler",med_fehler$Var1)
        med_fehler$Var1<-gsub("^E-.*","E-Fehler",med_fehler$Var1)
        med_fehler$Var1<-gsub("^F-.*","F-Fehler",med_fehler$Var1)
        med_fehler$Var1<-gsub("^G-.*","G-Fehler",med_fehler$Var1)
        med_fehler$Var1<-gsub("^H-.*","H-Fehler",med_fehler$Var1)
        med_fehler$Var1<-gsub("^I-.*","I-Fehler",med_fehler$Var1)
        med_fehler$Var1<-gsub("^W-.*","W-Fehler",med_fehler$Var1)
        
        med_fehler2<-med_fehler[med_fehler$Var1!="nicht geprüft",]
        med_fehler2$Rel<-round(100*med_fehler2$Freq/sum(med_fehler2$Freq))
        label_med_fehler<-med_fehler2$Var1
        
        
        output_proMed<-med_pruefung[c(which(med_pruefung$Var1=="nicht geprüft"),which(med_pruefung$Var1!="nicht geprüft")),]
        if(nrow(output_proMed)>0){
          if(nrow(med_pruefung2)==0){
            output_proMed<-cbind(output_proMed,NA)
          }else{
            if(length(which(output_proMed$Var1=="nicht geprüft"))>0){
              output_proMed<-cbind(output_proMed,rbind(NA,
                                                       med_pruefung2[c(which(med_pruefung2$Var1=="nicht geprüft"),which(med_pruefung2$Var1!="nicht geprüft")),])[,3])
            }else{
              output_proMed<-cbind(output_proMed,rbind(med_pruefung2[c(which(med_pruefung2$Var1=="nicht geprüft"),which(med_pruefung2$Var1!="nicht geprüft")),])[,3])
            }
          }
        }else{
          output_proMed[1,1]<-NA
          output_proMed<-cbind(output_proMed,NA,NA)
        }
 
        output_proMed[,3]<-paste0(output_proMed[,3],"%")
        output_proMed[,4]<-paste0(output_proMed[,4],"%")
        output_proMed[output_proMed$Var1=="nicht geprüft",4]<-NA
        names(output_proMed)<-c("Bereiche der Medikationsprüfung","Medikationsprüfungen pro Bereich","Relativ (alle)","Relativ (nur 'geprüft')")
        output_proMed$`Relativ (alle)`[output_proMed$`Relativ (alle)`=="0%"]<-"<1%"
        output_proMed$`Relativ (nur 'geprüft')`[output_proMed$`Relativ (nur 'geprüft')`=="0%"]<-"<1%"
        output_proMed[output_proMed=="NA%"]<-NA

        addWorksheet(wb,sheetName="Medikationsprüfungen")
        writeData(wb,"Medikationsprüfungen",output_proMed)
        addStyle(wb,"Medikationsprüfungen",bold,rows = 1,cols=c(1:ncol(output_proMed)))
        align<-createStyle(halign = "right")
        addStyle(wb,"Medikationsprüfungen",align,rows = c(2:(nrow(output_proMed)+1)),cols=c(3:4),gridExpand = T)
        
        width_vec1<-apply(output_proMed,2,function(x){
          y<-str_split_fixed(x,"\n",Inf)
          return(max(nchar(y) + 2,1, na.rm = TRUE))
        })
        width_vec2<-apply(str_split_fixed(names(output_proMed),pattern = "\n",Inf),1,function(x){max(nchar(x))})+2
        width_vec_proMed<-apply(cbind(width_vec1,width_vec2),1,max)
        
        
        
        png(paste0(tempdir(), "/", "plot3.png"), width=1700, height=400,units = "px")
        if(nrow(med_pruefung)>0){
          layout(matrix(c(1,1,2,2,3),ncol=5))
          par(mar=c(4,6,5,1))
          if(sum(med_pruefung$Var1=="nicht geprüft")>0){
            barplot(xlim=c(0,2.5),ylim=c(0,100),med_pruefung$Rel[med_pruefung$Var1=="nicht geprüft"],col="grey",
                    names.arg = "nicht geprüft",las=1,main="Medikationsprüfungen",yaxt="n",cex.names=2,cex.main=3)
          }
          if(sum(med_pruefung$Var1!="nicht geprüft")>0){
            if(sum(med_pruefung$Var1=="nicht geprüft")>0){
              barplot(xlim=c(0,2.5),ylim=c(0,100),as.matrix(med_pruefung$Rel[med_pruefung$Var1!="nicht geprüft"]),
                      add = T,space=1.5,col=c("skyblue1","orange","gold","green3","dodgerblue3","forestgreen")[1:nrow(med_pruefung2)],
                      names.arg="geprüft",yaxt="n",cex.names=2,main=3)              
            }else{
              barplot(xlim=c(0,2.5),ylim=c(0,100),as.matrix(med_pruefung$Rel[med_pruefung$Var1!="nicht geprüft"]),
                      add = F,space=1.5,col=c("skyblue1","orange","gold","green3","dodgerblue3","forestgreen")[1:nrow(med_pruefung2)],
                      names.arg="geprüft",yaxt="n",cex.names=2,main=3)
            }
          }
          
          axis(2,seq(0,100,20),labels=paste0(seq(0,100,20),"%"),las=1,cex.axis=2)
          
          if(nrow(med_pruefung2)>0){
            par(mar=c(0,1.5,5,1.5))
            own_pie(med_pruefung2$Freq,col = c("skyblue1","orange","gold","green3","dodgerblue3","forestgreen")[1:nrow(med_pruefung2)],
                    clockwise = T,labels = paste0(med_pruefung2$Rel,"%"),main="Pro Bereich")
            
            par(mar=c(0.5,0.5,0.5,0.5))
            plot(NULL,xlim=c(0,1),ylim=c(0,2),xaxt="n",yaxt="n",xlab="",ylab="",bty="n")
            legend("center",legend=label_med_pruefung2,cex=2,y.intersp = 2,
                   fill=c("skyblue1","orange","gold","green3","dodgerblue3","forestgreen")[1:nrow(med_pruefung2)])
          }else{
            plot(NULL,xlim=c(0,1),ylim=c(0,1),xaxt="n",yaxt="n",xlab="",ylab="",bty="n",
                 main="Pro Bereich",cex.main=3)
            text(0.5,0.5,"Keine Daten vorhanden",cex=2)
            plot(NULL,xlim=c(0,1),ylim=c(0,2),xaxt="n",yaxt="n",xlab="",ylab="",bty="n")
          }
        }else{
          plot(NULL,xlim=c(0,1),ylim=c(0,1),xaxt="n",yaxt="n",xlab="",ylab="",bty="n",
               main="Medikationsprüfungen",cex.main=3)
          text(0.5,0.5,"Keine Daten vorhanden",cex=2)
        }
        dev.off()
        
        insertImage(wb, sheet="Medikationsprüfungen", paste0(tempdir(), "/", "plot3.png"), width = 17,height = 4,
                    startRow = 1,startCol = ncol(output_proMed)+4)
        
        
        
        
        ###pro Fachrichtung
        
        lexikon_fach<-data.frame(Abk=c("KAIT","KHAU","KEND","KGHI","CHG","KMKG","KGYN","KHAE","KCHH","KKAR","CHK","KKJP","KNEP",
                                       "KCHN","KNEU","KNRAD","KNUK","KAUG","KORT","KHNO","KPAE","KCHP","KPNE","KPSY","KPSM","KDR",
                                       "KSTR","KCHU","KURO","KCHI","KPHO","CHT",
                                       "KCHK","PNE","KNRA"),
                                 Full=c("Anästhesiologie","Dermatologie","Endokrinologie","Gastro/Hepato/Infektiologie","Gefäßchirurgie",
                                        "Gesichtschirurgie","Gynäkologie","Hämato/Onkologie","Herz/Thoraxchirurgie","Kardiologie",
                                        "Kinderchirurgie","Kinderpsychiatrie","Nephrologie","Neurochirurgie","Neurologie","Neuroradiologie",
                                        "Nuklearmedizin","Ophtalmologie","Orthopädie","Oto/Rhino/Laryngologie","Pädiatrie","Plastische Chirurgie",
                                        "Pneumologie","Psychiatrie","Psychosomatik","Radiologie","Strahlentherapie","Unfallchirurgie",
                                        "Urologie","Viszeralchirurgie","Kinder Hämato/Onkologie","Thoraxchirurgie",
                                        "Kinderchirurgie","Pneumologie","Neuroradiologie"))
        
        med_fb<-as.data.frame(table(input6$fachrichtung))
        if(nrow(med_fb)>0){
          med_fb<-med_fb[order(med_fb$Freq,decreasing = T),]
          if(input$fb_in!="Alle"){
            auswahl<-as.numeric(substr(input$fb_in,start = 5,stop = nchar(input$fb_in)))
            med_fb<-med_fb[c(1:auswahl),]
          }
          
          med_fb$Rel<-round(100*med_fb$Freq/sum(med_fb$Freq))
          
          output_fb<-med_fb
        }else{
          output_fb<-data.frame(Var1=NA,Freq=NA,Rel=NA)
        }
        
        output_fb$Var1<-paste0(output_fb$Var1," ",lexikon_fach$Full[match(med_fb$Var1,lexikon_fach$Abk)])
        output_fb[,3]<-paste0(output_fb[,3],"%")
        names(output_fb)<-c("Fachrichtungen","Medikationsprüfungen pro Frachrichtung","Relativ")
        output_fb$Relativ[output_fb$Relativ=="0%"]<-"<1%"
        output_fb[output_fb=="NA%"]<-NA
        
        writeData(wb,"Medikationsprüfungen",output_fb,startRow = 21)
        addStyle(wb,"Medikationsprüfungen",bold,rows = 21,cols=c(1:ncol(output_fb)))
        align<-createStyle(halign = "right")
        addStyle(wb,"Medikationsprüfungen",align,rows = c(22:(nrow(output_fb)+22)),cols=c(3),gridExpand = T)
        
        width_vec1<-apply(output_fb,2,function(x){
          y<-str_split_fixed(x,"\n",Inf)
          return(max(nchar(y) + 2,1, na.rm = TRUE))
        })
        width_vec2<-apply(str_split_fixed(names(output_fb),pattern = "\n",Inf),1,function(x){max(nchar(x))})+2
        width_vec_fb<-apply(cbind(width_vec1,width_vec2),1,max)
        
        
        
        png(paste0(tempdir(), "/", "plot4.png"), width=1700, height=700,units = "px")
          par(mar=c(10,5,5,0.5))
          if(nrow(med_fb)>0){
            barplot(med_fb$Freq,names.arg =rep("",nrow(med_fb)),las=1,cex.main=2,cex.axis=1.5,
                    main="Medikationsprüfungen pro Fachrichtung",col=c("skyblue1","orange","gold","green3","dodgerblue3","forestgreen"))
            if(nrow(med_fb)>1){
              text(x=seq(0.75,(nrow(med_fb)*1.2),1.2),
                   y=rep(-3,nrow(med_fb)),
                   labels= paste0("",med_fb$Var1,"    \n",lexikon_fach$Full[match(med_fb$Var1,lexikon_fach$Abk)]),srt=60,adj=1,xpd=T)
            }else{
              text(x=0.75,
                   y=-1*max(med_fb$Freq)/30,
                   labels= paste0("",med_fb$Var1,"    \n",lexikon_fach$Full[match(med_fb$Var1,lexikon_fach$Abk)]),srt=0,xpd=T)
            }
          }else{
            plot(NULL,xlim=c(0,1),ylim=c(0,1),xaxt="n",yaxt="n",xlab="",ylab="",bty="n",
                 main="Medikationsprüfungen",cex.main=3)
            text(0.5,0.5,"Keine Daten vorhanden",cex=2)       
          }   
              
          dev.off()
          
          insertImage(wb, sheet="Medikationsprüfungen", paste0(tempdir(), "/", "plot4.png"), width = 17,height = 7,
                      startRow = 22,startCol = ncol(output_proMed)+4)
          
          
        
        ###mit ABP - Arzneimittelbezogenes Problem
        output_proMed2<-med_abp[c(which(med_abp$Var1=="nicht geprüft"),
                                  which(med_abp$Var1=="keine ABP"),
                                  which(med_abp$Var1=="keine Medikation erfasst"),
                                  which(med_abp$Var1!="nicht geprüft"&med_abp$Var1!="keine ABP"&med_abp$Var1!="keine Medikation erfasst")),]
        if(length(c(which(med_abp$Var1=="nicht geprüft"),which(med_abp$Var1=="keine ABP"),which(med_abp$Var1=="keine Medikation erfasst")))==3&&nrow(med_abp2)>0){
          output_proMed2<-cbind(output_proMed2,rbind(NA,NA,NA,med_abp2)[,3])
        }else{
          if(length(c(which(med_abp$Var1=="nicht geprüft"),which(med_abp$Var1=="keine ABP"),which(med_abp$Var1=="keine Medikation erfasst")))==2&&nrow(med_abp2)>0){
            output_proMed2<-cbind(output_proMed2,rbind(NA,NA,med_abp2)[,3])
          }else{
            if(length(c(which(med_abp$Var1=="nicht geprüft"),which(med_abp$Var1=="keine ABP"),which(med_abp$Var1=="keine Medikation erfasst")))==1&&nrow(med_abp2)>0){
              output_proMed2<-cbind(output_proMed2,rbind(NA,med_abp2)[,3])
            }else{
              if(nrow(output_proMed2)>0){
                output_proMed2<-cbind(output_proMed2,NA)
              }else{
                output_proMed2<-data.frame(Var1=NA,Freq=NA,Rel=NA,Rel2=NA)
              }
            }
          }
        }
        output_proMed2[,3]<-paste0(output_proMed2[,3],"%")
        output_proMed2[,4]<-paste0(output_proMed2[,4],"%")
        output_proMed2[c(which(output_proMed2$Var1=="nicht geprüft"),
                         which(output_proMed2$Var1=="keine ABP"),
                         which(output_proMed2$Var1=="keine Medikation erfasst")),4]<-NA
        names(output_proMed2)<-c("Bereiche mit ABP","ABP pro Bereich","Relativ (alle)","Relativ (nur 'mit ABP')")
        output_proMed2$`Relativ (alle)`[output_proMed2$`Relativ (alle)`=="0%"]<-"<1%"
        output_proMed2$`Relativ (nur 'mit ABP')`[output_proMed2$`Relativ (nur 'mit ABP')`=="0%"]<-"<1%"
        output_proMed2[output_proMed2=="NA%"]<-NA

        writeData(wb,"Medikationsprüfungen",output_proMed2,startRow = max(57,(21+nrow(output_fb)+3)))
        addStyle(wb,"Medikationsprüfungen",bold,rows = max(57,(21+nrow(output_fb)+3)),cols=c(1:ncol(output_proMed2)))
        
        addStyle(wb,"Medikationsprüfungen",align,rows = c((max(57,(21+nrow(output_fb)+3))+1):(max(57,(21+nrow(output_fb)+3))+1+nrow(output_proMed2))),
                 cols=c(3:4),gridExpand = T)
        
        
        width_vec1<-apply(output_proMed2,2,function(x){
          y<-str_split_fixed(x,"\n",Inf)
          return(max(nchar(y) + 2,1, na.rm = TRUE))
        })
        width_vec2<-apply(str_split_fixed(names(output_proMed2),pattern = "\n",Inf),1,function(x){max(nchar(x))})+2
        width_vec_proMed2<-apply(cbind(width_vec1,width_vec2),1,max)
        
        
        
        
        png(paste0(tempdir(), "/", "plot5.png"), width=1700, height=400,units = "px")
          layout(matrix(c(1,1,2,2,3),ncol=5))
          par(mar=c(4,6,5,1))
          if(nrow(as.matrix(med_abp$Rel[c(which(med_abp$Var1=="nicht geprüft"),
                                          which(med_abp$Var1=="keine ABP"),
                                          which(med_abp$Var1=="keine Medikation erfasst"))]))>0||
             nrow(as.matrix(med_abp$Rel[med_abp$Var1!="nicht geprüft"&med_abp$Var1!="keine ABP"&med_abp$Var1!="keine Medikation erfasst"]))>0){
            if(nrow(as.matrix(med_abp$Rel[c(which(med_abp$Var1=="nicht geprüft"),
                                            which(med_abp$Var1=="keine ABP"),
                                            which(med_abp$Var1=="keine Medikation erfasst"))]))>0){
              
              barplot(xlim=c(0,2.5),ylim=c(0,100),as.matrix(med_abp$Rel[c(which(med_abp$Var1=="nicht geprüft"),
                                                                          which(med_abp$Var1=="keine ABP"),
                                                                          which(med_abp$Var1=="keine Medikation erfasst"))]),
                      col=c("grey90","grey60","grey30"),beside = F,
                      names.arg = "ohne ABP",las=1,main="Medikationsprüfungen",yaxt="n",cex.names=2,cex.main=3)
              
              if(nrow(as.matrix(med_abp$Rel[med_abp$Var1!="nicht geprüft"&med_abp$Var1!="keine ABP"&med_abp$Var1!="keine Medikation erfasst"]))>0){
                barplot(xlim=c(0,2.5),ylim=c(0,100),as.matrix(med_abp$Rel[med_abp$Var1!="nicht geprüft"&med_abp$Var1!="keine ABP"&med_abp$Var1!="keine Medikation erfasst"]),
                        add = T,space=1.5,col=c("skyblue1","orange","gold","green3","dodgerblue3","forestgreen","darkorange3","goldenrod3","darkseagreen3")[1:(nrow(med_abp)-1)],
                        names.arg="mit ABP",yaxt="n",cex.names=2,main=3)
                axis(2,seq(0,100,20),labels=paste0(seq(0,100,20),"%"),las=1,cex.axis=2)
              }else{
                barplot(xlim=c(0,2.5),ylim=c(0,100),0,
                        add = T,space=1.5,col=c("skyblue1","orange","gold","green3","dodgerblue3","forestgreen","darkorange3","goldenrod3","darkseagreen3")[1:(nrow(med_abp)-1)],
                        names.arg="mit ABP",yaxt="n",cex.names=2,main=3)
                axis(2,seq(0,100,20),labels=paste0(seq(0,100,20),"%"),las=1,cex.axis=2)
              }
            }else{
              barplot(xlim=c(0,2.5),ylim=c(0,100),as.matrix(med_abp$Rel[med_abp$Var1!="nicht geprüft"&med_abp$Var1!="keine ABP"&med_abp$Var1!="keine Medikation erfasst"]),
                      space=1.5,col=c("skyblue1","orange","gold","green3","dodgerblue3","forestgreen","darkorange3","goldenrod3","darkseagreen3")[1:(nrow(med_abp)-1)],
                      names.arg="mit ABP",main="Medikationsprüfungen",yaxt="n",cex.names=2,cex.main=3)
              axis(2,seq(0,100,20),labels=paste0(seq(0,100,20),"%"),las=1,cex.axis=2)
            }
          }else{
            plot(NULL,xlim=c(0,1),ylim=c(0,1),xaxt="n",yaxt="n",xlab="",ylab="",bty="n",
                 main="Medikationsprüfungen",cex.main=3)
            text(0.5,0.5,"Keine Daten vorhanden",cex=2)
          }
          
          par(mar=c(0,1.5,5,1.5))
          if(nrow(med_abp2)>0){
            own_pie(med_abp2$Freq,col = c("skyblue1","orange","gold","green3","dodgerblue3","forestgreen","darkorange3","goldenrod3","darkseagreen3")[1:nrow(med_abp2)],
                    clockwise = T,labels = paste0(med_abp2$Rel,"%"),main="Mit ABP")
          }
          else{
            plot(NULL,xlim=c(0,1),ylim=c(0,1),xaxt="n",yaxt="n",xlab="",ylab="",bty="n",
                 main="Mit ABP",cex.main=3)
            text(0.5,0.5,"Keine Daten vorhanden",cex=2)
          }
          
          par(mar=c(0.5,0.5,0.5,0.5))
          plot(NULL,xlim=c(0,1),ylim=c(0,2),xaxt="n",yaxt="n",xlab="",ylab="",bty="n")
          if(length(c(which(med_abp$Var1=="nicht geprüft"),which(med_abp$Var1=="keine ABP"),which(med_abp$Var1=="keine Medikation erfasst")))==3&&nrow(med_abp)>0){
            if(nrow(med_abp)>9){
              legend("center",legend=label_med_abp,cex=1.7,y.intersp = 1.3,
                     fill=c("grey90","grey60","grey30","skyblue1","orange","gold","green3","dodgerblue3","forestgreen","darkorange3","goldenrod3","darkseagreen3")[1:nrow(med_abp)])
            }else{
              legend("center",legend=label_med_abp,cex=2,y.intersp = 1.5,
                     fill=c("grey90","grey60","grey30","skyblue1","orange","gold","green3","dodgerblue3","forestgreen","darkorange3","goldenrod3","darkseagreen3")[1:nrow(med_abp)])
            }
          }else{
            if(length(c(which(med_abp$Var1=="nicht geprüft"),which(med_abp$Var1=="keine ABP"),which(med_abp$Var1=="keine Medikation erfasst")))==2&&nrow(med_abp)>0){
              legend("center",legend=label_med_abp,cex=2,y.intersp = 1.5,
                     fill=c("grey90","grey60","skyblue1","orange","gold","green3","dodgerblue3","forestgreen","darkorange3","goldenrod3","darkseagreen3")[1:nrow(med_abp)])
            }else{
              if(length(c(which(med_abp$Var1=="nicht geprüft"),which(med_abp$Var1=="keine ABP"),which(med_abp$Var1=="keine Medikation erfasst")))==1&&nrow(med_abp)>0){
                legend("center",legend=label_med_abp,cex=2,y.intersp = 1.5,
                       fill=c("grey90","skyblue1","orange","gold","green3","dodgerblue3","forestgreen","darkorange3","goldenrod3","darkseagreen3")[1:nrow(med_abp)])
                }
            }
          }
          dev.off()
          
          insertImage(wb, sheet="Medikationsprüfungen", paste0(tempdir(), "/", "plot5.png"), width = 17,height = 4,
                      startRow = max(58,(21+nrow(output_fb)+3)),startCol = ncol(output_proMed)+4)
        
        
        
        ###Fehlerfolgenzuordnung
        lexikon_fehler<-data.frame(Fehlerfolge=c("A-Fehler","B-Fehler","C-Fehler","D-Fehler","E-Fehler","F-Fehler",
                                                 "G-Fehler","H-Fehler","I-Fehler","W-Fehler","keine Fehlerfolge","nicht geprüft"),
                                   Beschreibung=c("Umstände oder Ereignisse, die zu\nFehlern führen können",
                                                  "Fehler ist aufgetreten, hat jedoch\nden Patienten nicht erreicht",
                                                  "Fehler ist aufgetreten, der den\nPatienten zwar erreicht, diesem\njedoch keinen Schaden zugefügt hat",
                                                  "Fehler, der intensiviertes Monitoring\nund/ der Intervention erforderte, um\nschädliche Folgen für Patienten zu\nvermeiden",
                                                  "Fehler, der zu vorübergehenden\nschädlichen Folgen geführt oder\nbeigetragen hat",
                                                  "Fehler, der zu vorübergehenden\nschädlichen Folgen geführt oder\nbeigetragen hat und Hospitalisation\noder verlängerte Hospitalisation erforderte",
                                                  "Fehler, der zu permanentem Schaden\nführte oder dazu beigetragen hat",
                                                  "Fehler, der lebensrettende\nInterventionen erforderte",
                                                  "Fehler, der zum Tod des Patienten\nführte",
                                                  "fehlende Wirtschaftlichkeit:\nunnötige Kosten",
                                                  NA,NA),
                                   Schweregrad=c("ohne schädliche\nFolgen","ohne schädliche\nFolgen","ohne schädliche\nFolgen","potentiell\nschädliche Folgen","schädliche Folgen",
                                                 "schädliche Folgen","schädliche Folgen","schädliche Folgen","schädliche Folgen",NA,NA,NA))
        
        output_fehler<-med_fehler[c(which(med_fehler$Var1=="nicht geprüft"),
                                    which(med_fehler$Var1!="nicht geprüft"&med_fehler$Var1!="keine Fehlerfolge"),
                                    which(med_fehler$Var1=="keine Fehlerfolge")),]
        if(sum(output_fehler$Var1=="nicht geprüft")>0){
          output_fehler<-cbind(output_fehler,rbind(NA,
                                                   med_fehler2[c(which(med_fehler2$Var1=="nicht geprüft"),
                                                                 which(med_fehler2$Var1!="nicht geprüft"&med_fehler2$Var1!="keine Fehlerfolge"),
                                                                 which(med_fehler2$Var1=="keine Fehlerfolge")),])[,3])
        }else{
          output_fehler<-cbind(output_fehler,med_fehler2[c(which(med_fehler2$Var1=="nicht geprüft"),
                                                           which(med_fehler2$Var1!="nicht geprüft"&med_fehler2$Var1!="keine Fehlerfolge"),
                                                           which(med_fehler2$Var1=="keine Fehlerfolge")),][,3])
        }
        if(nrow(output_fehler)==0){
          output_fehler[1,1]<-NA
        }
        
        output_fehler[output_fehler[,1]=="keine Fehlerfolge",1]<-"keine\nFehlerfolge"
        output_fehler[,3]<-paste0(output_fehler[,3],"%")
        output_fehler[,4]<-paste0(output_fehler[,4],"%")
        output_fehler[output_fehler$Var1=="nicht geprüft",4]<-NA
        output_fehler$Beschreibung<-lexikon_fehler$Beschreibung[match(output_fehler$Var1,lexikon_fehler$Fehlerfolge)]
        output_fehler$Schweregrad<-lexikon_fehler$Schweregrad[match(output_fehler$Var1,lexikon_fehler$Fehlerfolge)]
        output_fehler<-output_fehler[,c(1,5,6,2,3,4)]
        names(output_fehler)<-c("Bereich/ABP","Beschreibung","Schweregrad","Prüfungen mit ABP\nund Fehlerfolgenzuordnung","Relativ (alle)","Relativ (nur 'geprüft')")
        output_fehler$`Relativ (alle)`[output_fehler$`Relativ (alle)`=="0%"]<-"<1%"
        output_fehler$`Relativ (nur 'geprüft')`[output_fehler$`Relativ (nur 'geprüft')`=="0%"]<-"<1%"
        output_fehler[output_fehler=="NA%"]<-NA
        
        writeData(wb,"Medikationsprüfungen",output_fehler,startRow = max(78,(max(57,(21+nrow(output_fb)+3))+1+nrow(output_proMed2)+1+2)))
        addStyle(wb,"Medikationsprüfungen",bold,rows = max(78,(max(57,(21+nrow(output_fb)+3))+1+nrow(output_proMed2)+1+2)),
                 cols=c(1:ncol(output_fehler)))

        addStyle(wb,"Medikationsprüfungen",align,rows = c(max(78,(max(57,(21+nrow(output_fb)+3))+1+nrow(output_proMed2)+1+3)):(max(78,(max(57,(21+nrow(output_fb)+3))+1+nrow(output_proMed2)+1+3))+nrow(output_fehler))),
                 cols=c(5:6),gridExpand = T,stack = T)
        
        width_vec1<-apply(output_fehler,2,function(x){
          y<-str_split_fixed(x,"\n",Inf)
          return(max(nchar(y) + 2,1, na.rm = TRUE))
        })
        width_vec2<-apply(str_split_fixed(names(output_fehler),pattern = "\n",Inf),1,function(x){max(nchar(x))})+2
        width_vec<-apply(cbind(width_vec1,width_vec2,c(width_vec_proMed,NA,NA),c(width_vec_proMed2,NA,NA),c(width_vec_fb,NA,NA,NA)),1,function(x){
          max(x,na.rm=T)
        })
        setColWidths(wb,"Medikationsprüfungen",cols = c(1:ncol(output_fehler)),widths = width_vec)
        
        wrap<-createStyle(wrapText = T)
        addStyle(wb,"Medikationsprüfungen",wrap,rows=c(1:(max(78,(max(57,(21+nrow(output_fb)+3))+1+nrow(output_proMed2)+1+3))+nrow(output_fehler))),cols=1:ncol(output_fehler),gridExpand = T,stack = T)
        
        med_fehler2<-med_fehler2[c(which(med_fehler2$Var1=="keine Fehlerfolge"),which(med_fehler2$Var1!="keine Fehlerfolge")),]
        
        png(paste0(tempdir(), "/", "plot6.png"), width=1700, height=400,units = "px")
        if(length(med_fehler2$Freq[med_fehler2$Rel!=0])>0){
          layout(matrix(c(1,1,1,2),ncol=4))
          par(mar=c(0,1.5,5,1.5))
          if(sum(label_med_fehler=="keine Fehlerfolge")==1){
            own_pie(med_fehler2$Freq[med_fehler2$Rel!=0],col = c("grey","skyblue1","orange","gold","green3","dodgerblue3",
                                                                 "forestgreen","darkorange3","goldenrod3","darkseagreen3","royalblue4")[1:nrow(med_fehler2[med_fehler2$Rel!=0,])],
                    clockwise = T,labels = paste0(med_fehler2$Rel,"%"),main="Prüfungen mit ABP und Anteil Fehlerfolgenzuordnung")          
          }else{
            own_pie(med_fehler2$Freq[med_fehler2$Rel!=0],col = c("skyblue1","orange","gold","green3","dodgerblue3",
                                                                 "forestgreen","darkorange3","goldenrod3","darkseagreen3","royalblue4")[1:nrow(med_fehler2[med_fehler2$Rel!=0,])],
                    clockwise = T,labels = paste0(med_fehler2$Rel,"%"),main="Prüfungen mit ABP und Anteil Fehlerfolgenzuordnung")
          }
          
          par(mar=c(0.5,0.5,0.5,0.5))
          plot(NULL,xlim=c(0,1),ylim=c(0,2),xaxt="n",yaxt="n",xlab="",ylab="",bty="n")
          if(sum(label_med_fehler=="keine Fehlerfolge")==1){
            legend("center",legend=label_med_fehler[c(which(label_med_fehler=="keine Fehlerfolge"),which(label_med_fehler!="keine Fehlerfolge"))],
                   cex=2,y.intersp = 1.5,
                   fill=c("grey","skyblue1","orange","gold","green3","dodgerblue3",
                          "forestgreen","darkorange3","goldenrod3","darkseagreen3","royalblue4")[1:length(label_med_fehler)])
          }else{
            legend("center",legend=label_med_fehler[c(which(label_med_fehler=="keine Fehlerfolge"),which(label_med_fehler!="keine Fehlerfolge"))],
                   cex=2,y.intersp = 1.5,
                   fill=c("skyblue1","orange","gold","green3","dodgerblue3",
                          "forestgreen","darkorange3","goldenrod3","darkseagreen3","royalblue4")[1:length(label_med_fehler)])
          }
        }else{
          plot(NULL,xlim=c(0,1),ylim=c(0,1),xaxt="n",yaxt="n",xlab="",ylab="",bty="n",
               main="Prüfungen mit ABP und Anteil Fehlerfolgenzuordnung",cex.main=3)
          text(0.5,0.5,"Keine Daten vorhanden",cex=2)
        }
          dev.off()
          
          insertImage(wb, sheet="Medikationsprüfungen", paste0(tempdir(), "/", "plot6.png"), width = 17,height = 4,
                      startRow = max(80,(max(57,(21+nrow(output_fb)+3+5))+1+nrow(output_proMed2)+1+2)),startCol = ncol(output_proMed)+4)

          
          
          start_anders_med2<-paste(strsplit(as.character(values$MED_Start),"-",fixed = T)[[1]][c(3,2,1)],collapse = ".")        
          ende_max_anders_med2<-paste(strsplit(as.character(values$MED_End),"-",fixed = T)[[1]][c(3,2,1)],collapse = ".")
          
          if(!is.null(input$day_type)&&input$day_type==F){
            help_klinik1_kurz<-strsplit(help_klinik1," ",fixed=T)[[1]][1]
            name_export<-paste0("Export Analyse Pharmazeutische Kurvenvisite ",start_anders_med2,"-",ende_max_anders_med2," ",help_klinik1_kurz,".xlsx")
          }else{
            name_export<-paste0("Export Analyse Pharmazeutische Kurvenvisite ",start_anders_med2,"-",ende_max_anders_med2,".xlsx")
          }
          

        #pfad<-paste0("~",paste(unlist(input$export_dir2[[1]]),collapse="/"))
        if(paste(unlist(input$export_dir2[[1]]),collapse="/")==0||paste(unlist(input$export_dir2[[1]]),collapse="/")==1){
          #pfad<-"~"
          pfad<-"C:/"
        }else{
          #pfad<-paste0("~",paste(unlist(input$export_dir2[[1]]),collapse="/"))
          pfad<-paste0("C:",paste(unlist(input$export_dir2[[1]]),collapse="/"))
        }
          #if(file.opened(paste0(pfad,"/",name_export))){
          #  shinyjs::html("text", paste0("<br><br>Datei ",paste0(pfad,"/",name_export)," ist noch geöffnet. Bitte schließen Sie zuerst die Datei.<br><br>"), add = TRUE)
          #  return()
          #}
        #saveWorkbook(wb,file=paste0(pfad,"/",name_export),overwrite = T)
        
        tt <- tryCatch(saveWorkbook(wb,file=paste0(pfad,"/",name_export),overwrite = T),
                       warning=function(w) w)
        if(is.list(tt)==F){
          saveWorkbook(wb,file=paste0(pfad,"/",name_export),overwrite = T)
        }else{
          if(length(grep("Permission denied",tt$message))>0){
            shinyjs::html("text", paste0("<br><br>Zugriff verweitert.<br><br>
                                         Falls Sie die exportierte Datei aus einer vorherigen Analyse noch geöffnet haben: bitte schließen Sie diese oder ändern Sie den Dateinamen, damit ein Update exportiert werden kann. 
                                         <br>Bitte stellen Sie weiterhin sicher, dass Sie für den ausgewählten Ordner Schreibrechte besitzen (für C:/ häufig nicht der Fall)."), add = TRUE)
          }else{
            shinyjs::html("text", paste0("<br><br>",tt$message), add = TRUE)
          }
          return()
        }
        
        
        shinyjs::html("text", paste0("<br><br>Daten und Analyse-Ergebnisse erfolgreich exportiert.<br><br>"), add = TRUE)
        shinyjs::html("text", paste0("&nbsp;&nbsp;&nbsp;&nbsp;",paste0(pfad,"/",name_export),"<br>"), add = TRUE)
      
    })

  })
  
  
  
  session$onSessionEnded(function() {
    stopApp()
  })
  
})
