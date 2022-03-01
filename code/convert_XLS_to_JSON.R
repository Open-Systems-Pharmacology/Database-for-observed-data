options("warn"=0) # default
rm(list = ls()) # empty workspace
closeAllConnections()
require(readxl)
require(jsonlite)

#copy paste in der cmd:
#  cd "<path to OSP installation folder>\Open Systems Pharmacology\PK-Sim 7.3"
#  PKSim.exe/dev 

#fdir <- dirname(sys.frame(1)$ofile) # get location of this script when script is sourced
fdir <- dirname(rstudioapi::getActiveDocumentContext()$path) # get location of this script using R studio
setwd(file.path(fdir,".."))
source(file.path("code","auxiliaryFunctions.R"))

sharePoint <- getwd()
DBfileName <- "ObsDataPK_OSP.xlsx"

# read in database
studies <- read_xlsx(file.path(sharePoint,DBfileName), sheet = "Studies",col_types = "text", skip = 1)
df <- read_xlsx(file.path(sharePoint,DBfileName),sheet="PK-Profiles",col_types = "text")
MW <- read_xlsx(file.path(sharePoint,DBfileName),sheet="Analyte",col_types = "text")
proj <- read_xlsx(file.path(sharePoint,DBfileName),sheet="Projects",col_types = "text")

# convert NA to empty string
studies <- data.frame(apply(studies,MARGIN = 2,replaceNAbyEmptyString),stringsAsFactors = FALSE)
df      <- data.frame(apply(df,MARGIN = 2,replaceNAbyEmptyString),stringsAsFactors = FALSE)
MW      <- data.frame(apply(MW,MARGIN = 2,replaceNAbyEmptyString),stringsAsFactors = FALSE)
proj    <- data.frame(apply(proj,MARGIN = 2,replaceNAbyEmptyString),stringsAsFactors = FALSE)

# remove white space (left and right)
studies <- data.frame(apply(studies,MARGIN = 2,trimws,"both"),stringsAsFactors = FALSE)
df      <- data.frame(apply(df,MARGIN = 2,trimws,"both"),stringsAsFactors = FALSE)
MW      <- data.frame(apply(MW,MARGIN = 2,trimws,"both"),stringsAsFactors = FALSE)
proj    <- data.frame(apply(proj,MARGIN = 2,trimws,"both"),stringsAsFactors = FALSE)

ids <- unique(df$ID)
PROJECTS <- unique(trimws(strsplit(paste(proj$Projects[is.element(proj$ID,ids)],collapse = ","),split = ",", fixed = TRUE)[[1]],which = "both"))
mapProjToIDs <- lapply(strsplit(paste(proj$Projects[is.element(proj$ID,ids)]),split = ",", fixed = TRUE),trimws,"both")
names(mapProjToIDs) <- proj$ID[is.element(proj$ID,ids)]

# Create subdirectory
if (!dir.exists(file.path(getwd(),"buildingBlocks"))){
  dir.create(file.path(getwd(),"buildingBlocks"))
}

# Run conversion
{
  BB_all <- data.frame(stringsAsFactors = FALSE)
  eachID_collector <- NULL
  for(eachProject in PROJECTS[PROJECTS!=""]){
    counter <- 0
    BB <- data.frame(stringsAsFactors = FALSE)
    idsInProject <- as.numeric(names(mapProjToIDs[unlist(lapply(mapProjToIDs,FUN = function(x) any(is.element(x,eachProject))))]))
    for(eachID in idsInProject){
      counter <- counter+1
      SELECT_study <- studies$ID == eachID
      SELECT_df    <- df$ID == eachID
      SELECT_projects <- proj$ID == eachID
      if(!(!any(studies$Study[SELECT_study] != c(df$Study[SELECT_df][1],proj$Study[SELECT_projects])) &
           !any(studies$Reference[SELECT_study] != c(df$Reference[SELECT_df][1],proj$Reference[SELECT_projects])) &
           !any(studies$Grouping[SELECT_study] != c(df$Grouping[SELECT_df][1],proj$Grouping[SELECT_projects])) &
           !any(studies$Compound.Analyte[SELECT_study] != c(df$Analyte[SELECT_df][1],proj$Analyte[SELECT_projects])) &
           !any(studies$Compartment[SELECT_study] != c(df$Compartment[SELECT_df][1])))){
        stop("Inconsistency over sheets. No matching for ",eachID)
      }
      # get x and y
      rawAvgUnit <- unique(tolower(df$AvgUnit[SELECT_df & df$Avg!=""]))
      if(length(rawAvgUnit)>1) stop("Multiple Avg units found within one data series for ",eachID)
      rawAvgType <- unique(tolower(df$AvgType[SELECT_df & df$Avg!=""]))
      if(length(rawAvgType)>1) stop("Multiple Avg types found within one data series for ",eachID)
      rawVarUnit <- unique(tolower(df$VarUnit[SELECT_df & df$Var!=""]))
      if(length(rawVarUnit)>1) stop("Multiple Var units found within one data series for ",eachID)
      rawVarType <- unique(tolower(df$VarType[SELECT_df  & df$Var!=""]))
      if(length(rawVarType)>1) stop("Multiple Var types found within one data series for ",eachID)
      
      
      avgColumn  <-getAverage(avgValue = df$Avg[SELECT_df],avgUnit = rawAvgUnit,avgType = rawAvgType,
                              timeValue = df$Time[SELECT_df],df$LLOQ[SELECT_df],eachID)
      timeColumn  <-getTime(timeValue = df$Time[SELECT_df],timeUnit = df$TimeUnit[SELECT_df][1],avgValue = df$Avg[SELECT_df],df$LLOQ[SELECT_df],eachID)
      varColumn <- getVariance(varValue = df$Var[SELECT_df],varUnit =rawVarUnit,varType = rawVarType,
                               avgValue = df$Avg[SELECT_df],avgUnit = rawAvgUnit,
                               timeValue = df$Time[SELECT_df],n = as.numeric(studies$N[SELECT_study]),df$LLOQ[SELECT_df],eachID)
      
      if(tolower(studies$Data.type[SELECT_study])=="aggregated"){
        dataType <- "aggregated"
        dataTypeNaming <- paste0("agg.",ifelse(as.character(studies$N[SELECT_study])!="",paste0(" (n=",as.character(studies$N[SELECT_study]),")"),""))
      } else if(tolower(studies$Data.type[SELECT_study])=="typical"){
        dataType <- "typical representative"
        dataTypeNaming <- paste0("typical",ifelse(as.character(studies$N[SELECT_study])!="",paste0(" (n=",as.character(studies$N[SELECT_study]),")"),""))
      } else if(tolower(studies$Data.type[SELECT_study])=="individual"){
        dataType <- "individual"
        dataTypeNaming <- "indiv."
      } else {
        stop("Unknown data type (neither aggregated nor individual) specified for ",eachID)
      }
      
      datasetName <- paste(studies$Study[SELECT_study],
                           studies$Grouping[SELECT_study],
                           studies$Compound.Analyte[SELECT_study],
                           studies$Route[SELECT_study],
                           paste(studies$Dose[SELECT_study],studies$Dose.Unit[SELECT_study]),
                           studies$Compartment[SELECT_study],
                           dataTypeNaming,
                           sep=" - ")
      # Name
      {
        BB[counter,"Name"] <- datasetName
      }
      # ExtendedProperties
      {
        extProp <- as.data.frame(matrix(c("DB Version","OSP DATABASE",
                                          "ID",as.character(studies$ID[SELECT_study]),
                                          "Study Id",studies$Study[SELECT_study],
                                          "Reference",as.character(studies$Reference[SELECT_study]),
                                          "Source",studies$Source[SELECT_study],
                                          "Grouping",studies$Grouping[SELECT_study],
                                          "Data type", dataType,
                                          "N",ifelse(as.character(studies$N[SELECT_study])!="",as.character(studies$N[SELECT_study]),"n/a"),
                                          "Molecule",studies$Compound.Analyte[SELECT_study],
                                          "Species",studies$Species[SELECT_study],
                                          "Organ",getOrgan(studies$Compartment[SELECT_study],eachID),
                                          "Compartment",getCompartment(studies$Compartment[SELECT_study],eachID),
                                          "Route",studies$Route[SELECT_study],
                                          "Dose",paste(studies$Dose[SELECT_study],studies$Dose.Unit[SELECT_study]),
                                          paste0("Times of Administration [",as.character(studies$Times.Unit[SELECT_study]),"]"),as.character(studies$Times.of.Administration[SELECT_study]),
                                          "Formulation",ifelse(studies$Formulation.type[SELECT_study]!="",studies$Formulation.type[SELECT_study],"."),
                                          "Food state",ifelse(studies$Fasted.Fed.state[SELECT_study]!="",studies$Fasted.Fed.state[SELECT_study],"."),
                                          "Comment",varColumn$comment),
                                        ncol=2,byrow = TRUE),stringsAsFactors = FALSE)
        colnames(extProp) <- c("Name","Value")
        BB$ExtendedProperties[[counter]] <- extProp
      }
      
      # Columns
      {
        BB$Columns[[counter]] <- data.frame(stringsAsFactors = FALSE)
        # Name
        {
          BB$Columns[[counter]][1,"Name"] <- "Avg"
        }
        # QuantityInfo
        {
          BB$Columns[[counter]]$QuantityInfo <- data.frame(
            Name="Avg",
            Path=paste(datasetName,
                       "ObservedData",
                       getOrgan(studies$Compartment[SELECT_study],eachID),
                       getCompartment(studies$Compartment[SELECT_study],eachID),
                       studies$Compound.Analyte[SELECT_study],
                       avgColumn$type,sep="|"),
            stringsAsFactors = FALSE)
        }
        # DataInfo
        {
          curMolWeight <- as.numeric(MW[tolower(MW[,1])==tolower(BB$ExtendedProperties[[counter]]$Value[BB$ExtendedProperties[[counter]]$Name == "Molecule"]),2])
          if(length(curMolWeight)!=1){
            stop("Molecular weight not specified for ",BB$ExtendedProperties[[counter]]$Value[BB$ExtendedProperties[[counter]]$Name == "Molecule"])
          }
          BB$Columns[[counter]]$DataInfo <- data.frame(
            Origin="Observation",
            AuxiliaryType=getAuxiliaryType(avgColumn$type,eachID),
            Source="",
            MolWeight=curMolWeight,
            stringsAsFactors = FALSE)
          if(length(avgColumn$lloq)>0) BB$Columns[[counter]]$DataInfo$LLOQ=as.numeric(avgColumn$lloq)*getUnitFactorToBaseUnit(unit = avgColumn$unit,eachID)
        }
        # Dimension
        {
          BB$Columns[[counter]]$Dimension <- avgColumn$dim
        }
        # Unit
        {
          BB$Columns[[counter]]$Unit <- avgColumn$unit
        }
        # Values
        {
          if(any(is.na(avgColumn$value))) avgColumn$value[is.na(avgColumn$value)] <- "NaN"
          BB$Columns[[counter]]$Values <- list(I(avgColumn$value))
        }
        # RelatedColumns
        {
          if(varColumn$dim!=""){
            BB$Columns[[counter]]$RelatedColumns[[1]] <- list()
            BB$Columns[[counter]]$RelatedColumns[[1]] <- data.frame(stringsAsFactors = FALSE)
            # Name
            {
              BB$Columns[[counter]]$RelatedColumns[[1]][1,"Name"] <- "Var"
            }
            # QuantityInfo
            {
              BB$Columns[[counter]]$RelatedColumns[[1]]$QuantityInfo <- data.frame(
                Name="Var",
                Path=paste(datasetName,
                           "ObservedData",
                           getOrgan(studies$Compartment[SELECT_study],eachID),
                           getCompartment(studies$Compartment[SELECT_study],eachID),
                           studies$Compound.Analyte[SELECT_study],
                           varColumn$type,sep="|"),
                stringsAsFactors = FALSE)
            }
            # DataInfo
            {
              BB$Columns[[counter]]$RelatedColumns[[1]]$DataInfo <- data.frame(
                Origin="ObservationAuxiliary",
                AuxiliaryType=getAuxiliaryType(varColumn$type,eachID),
                Source="",
                MolWeight=curMolWeight,
                stringsAsFactors = FALSE)
            }
            # Dimension
            {
              BB$Columns[[counter]]$RelatedColumns[[1]]$Dimension <- varColumn$dim
            }
            # Unit
            {
              if(varColumn$unit!=""){
                BB$Columns[[counter]]$RelatedColumns[[1]]$Unit <- varColumn$unit
              }
            }
            # Values
            {
              if(any(is.na(varColumn$value))) varColumn$value[is.na(varColumn$value)] <- "NaN"
              BB$Columns[[counter]]$RelatedColumns[[1]]$Values <- list(I(varColumn$value))
            }
          }
        }
      }
      # BaseGrid
      {
        BB$BaseGrid[[counter]] <- list()
        # Name
        {
          #BB$BaseGrid[[counter]][1,"Name"] <- "Time"
          BB$BaseGrid[[counter]]$Name <- "Time"
        }
        
        # QuantityInfo
        {
          BB$BaseGrid[[counter]]$QuantityInfo <- list(
            Name="Time",
            Path=paste(datasetName,"Time",sep="|"),
            Type="Time")
        }
        # DataInfo
        {
          BB$BaseGrid[[counter]]$DataInfo <- list(
            Origin="BaseGrid",
            AuxiliaryType="Undefined",
            Source="")
        }
        # Dimension
        {
          BB$BaseGrid[[counter]]$Dimension <- timeColumn$dim
        }
        # Unit
        {
          BB$BaseGrid[[counter]]$Unit <- timeColumn$unit
        }
        # Values
        {
          if(any(is.na(timeColumn$value))) timeColumn$value[is.na(timeColumn$value)] <- "NaN"
          BB$BaseGrid[[counter]]$Values <- I(timeColumn$value)
        }
      }
      # add to BB_all
      {
        if(!(eachID %in% eachID_collector)){
            idx <- nrow(BB_all)+1
            for(jj in 1:ncol(BB)){
              BB_all[idx,jj] <- ifelse(is.list(BB[nrow(BB),jj]),list(BB[nrow(BB),jj]),BB[nrow(BB),jj])
            }
            if(idx == 1) colnames(BB_all) <- colnames(BB)
        }
        eachID_collector <- unique(c(eachID_collector,eachID))
      }
    }
    cat("Writing: ",paste0("ObsData_",eachProject,".json"),"\n")
    curFile <- file(file.path(getwd(),paste0("ObsData_",eachProject,".json")),encoding="UTF-8")
    write(toJSON(BB, digits=I(10), pretty=T, auto_unbox=T),file = curFile)
    close(curFile)
  }
  cat("Writing: ",paste0("ObsData_","ALL",".json"),"\n")
  curFile <- file(file.path(getwd(),paste0("ObsData_","ALL",".json")),encoding="UTF-8")
  write(toJSON(BB_all, digits=I(10), pretty=T, auto_unbox=T),file = curFile)
  close(curFile)
}
for(eachBB in 1:dim(BB_all)[1]){
  fileName <- paste(BB_all[eachBB,]$ExtendedProperties[[1]]$Value[2],BB_all[eachBB,]$Name,sep=" - ")
  fileName <- gsub("[?*\"<>|\\/]", "_", fileName)
  curFile <- file(file.path(getwd(),"buildingBlocks",paste0(fileName,".json")),encoding="UTF-8")
  write(toJSON(BB_all[eachBB,], digits=I(10), pretty=T, auto_unbox=T),file = curFile)
  close(curFile)
}




