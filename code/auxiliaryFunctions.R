replaceNAbyEmptyString <- function(x){
  x[is.na(x)] <- ""
  return(x)
}

getTime <- function(timeValue,timeUnit,avgValue,lloqValue,eachID){
  timeUnit <- tolower(timeUnit)
  if(length(timeValue) != length(avgValue)) stop("Unequal length of time and avg values"," for ",eachID)
  idLLOQ <- grepl(pattern = "<",avgValue,fixed = TRUE)
  if(length(idLLOQ)>0 | any(lloqValue != "")){
    avgValue <- gsub(pattern = "<",replacement = "",x = avgValue,fixed = TRUE)
    LLOQ <- unique(c(as.numeric(unique(avgValue[idLLOQ])),as.numeric(lloqValue[lloqValue != ""])))
    LLOQ <- LLOQ[!is.na(LLOQ)]
    if(length(LLOQ)>1) stop("Multiple LLOQ detected within one dataset"," for ",eachID)
    avgValue <- as.numeric(avgValue)
    avgValue[idLLOQ] <- LLOQ/2
  } else {
    LLOQ <- ""
    avgValue <- as.numeric(avgValue)
  }
  timeValue <- as.numeric(timeValue)
  timeValue <- timeValue[!is.na(avgValue) & !is.na(timeValue) & !avgValue <= 0]
  timeUnit <- trimws(timeUnit,which = "both")
  timeDim <- getDimension(timeUnit,eachID)
  if(timeDim != "Time") stop("Incorrect unit for time for ",eachID)
  return(list(value=timeValue,unit=timeUnit,dim=timeDim))
}

getAverage <- function(avgValue,avgUnit,avgType,timeValue,lloqValue,eachID){
  avgUnit <- tolower(avgUnit)
  if(length(timeValue) != length(avgValue)) stop("Unequal length of time and avg values"," for ",eachID)
  if(length(avgType)==0 || tolower(avgType)==""){
    avgType <- ""
    avgDim <- ""
    avgUnit <- ""
    avgValue <- rep(NaN,length(avgValue))
    stop("No average type specified for ",eachID)
  } else {
    idLLOQ <- grepl(pattern = "<",avgValue,fixed = TRUE)
    if(length(idLLOQ)>0 | any(lloqValue != "")){
      avgValue <- gsub(pattern = "<",replacement = "",x = avgValue,fixed = TRUE)
      LLOQ <- unique(c(as.numeric(unique(avgValue[idLLOQ])),as.numeric(lloqValue[lloqValue != ""])))
      LLOQ <- LLOQ[!is.na(LLOQ)]
      if(length(LLOQ)>1) stop("Multiple LLOQ detected within one dataset"," for ",eachID)
      avgValue <- as.numeric(avgValue)
      avgValue[idLLOQ] <- LLOQ/2
    } else {
      LLOQ <- ""
      avgValue <- as.numeric(avgValue)
    }
    timeValue <- as.numeric(timeValue)
    avgValue <- avgValue[!is.na(avgValue) & !is.na(timeValue) & !avgValue <= 0]
    avgUnit <- trimws(avgUnit,which = "both")
    avgType <- trimws(avgType,which = "both")
    avgDim <- getDimension(avgUnit,eachID)
    if(tolower(avgType)=="arith. mean"){
      avgType <- "ArithmeticMean"
    } else if(tolower(avgType)=="geo. mean" | tolower(avgType)=="geom. mean"){
      avgType <- "GeometricMean"
    } else if(tolower(avgType)=="individual"){
      avgType <- "Individual"
    } else if(tolower(avgType)=="median"){
      avgType <- "Median"
    } else {
      stop("Unknown average type: ",avgType," for ",eachID)
    }
  }
  return(list(value=avgValue,unit=avgUnit,type=avgType,dim=avgDim,lloq=LLOQ))
}

getVariance <- function(varValue,varUnit,varType,avgValue,avgUnit,timeValue,n,lloqValue,eachID){
  varUnit <- tolower(varUnit)
  avgUnit <- tolower(avgUnit)
  
  if(length(varValue) != length(avgValue)) stop("Unequal length of variance and avg values"," for ",eachID)
  if(length(varType)==0 || tolower(varType)==""){
    varType <- ""
    varDim <- ""
    varUnit <- ""
    varValue <- rep(NaN,length(varValue))
    varComment <- "."
    
  } else {
    idLLOQ <- grepl(pattern = "<",avgValue,fixed = TRUE)
    if(length(idLLOQ)>0 | any(lloqValue != "")){
      avgValue <- gsub(pattern = "<",replacement = "",x = avgValue,fixed = TRUE)
      LLOQ <- unique(c(as.numeric(unique(avgValue[idLLOQ])),as.numeric(lloqValue[lloqValue != ""])))
      LLOQ <- LLOQ[!is.na(LLOQ)]
      if(length(LLOQ)>1) stop("Multiple LLOQ detected within one dataset"," for ",eachID)
      avgValue <- as.numeric(avgValue)
      avgValue[idLLOQ] <- LLOQ/2
    } else {
      LLOQ <- ""
      avgValue <- as.numeric(avgValue)
    }
    timeValue <- as.numeric(timeValue)
    #avgValue <- avgValue[!is.na(avgValue) & !is.na(timeValue)]
    
    varValue <- varValue[!is.na(avgValue) & !is.na(timeValue) & !avgValue <= 0]
    avgValue <- avgValue[!is.na(avgValue) & !is.na(timeValue) & !avgValue <= 0]
    varValue <- as.numeric(varValue)
    varUnit <- trimws(varUnit,which = "both")
    varType <- trimws(varType,which = "both")
    if(tolower(varType)=="arith. sd"){
      varType <- "ArithmeticStdDev"
      varDim <- getDimension(varUnit,eachID)
      if(varDim != getDimension(avgUnit,eachID)) stop("Wrong dimension for ",varType," for ",eachID)
      varValue[is.na(varValue) | varValue<0] <- NaN
      varComment <- "."
    } else if(tolower(varType)=="geo. sd" | tolower(varType)=="geom. sd"){
      varType <- "GeometricStdDev"
      if(varUnit==""){
        varDim <- "Dimensionless"
      } else {
        stop("Unknown unit for ",varType," for ",eachID)
      }
      varValue[is.na(varValue) | varValue<1] <- NaN
      varComment <- "."
    } else if(tolower(varType)=="arith. sem"){
      if(!is.na(as.numeric(n)) & as.numeric(n) > 1){
        varType <- "ArithmeticStdDev"
        varDim <- getDimension(varUnit,eachID)
        if(varDim != getDimension(avgUnit,eachID)) stop("Wrong dimension for ",varType," for ",eachID)
        varValue <- varValue * sqrt(as.numeric(n))
        varValue[is.na(varValue) | varValue<0] <- NaN
        varComment <- "Arith. SEM converted to arith. SD"
      } else {
        varType <- ""
        varDim <- ""
        varValue[1:length(varValue)] <- NaN
        varComment <- "Arith. SEM not converted to arith. SD (N not specified)"
      }
    } else if(tolower(varType)=="geo. cv"){
      varType <- "GeometricStdDev"
      if(varUnit=="%"){
        varDim <- "Dimensionless"
        varValue <- varValue/100
      } else if(varUnit==""){
        varDim <- "Dimensionless"
      } else{
        stop("Unknown unit for ",varType," for ",eachID)
      }
      varUnit <- ""
      varValue <- exp(sqrt(log(varValue^2+1)))
      varValue[is.na(varValue) | varValue<0] <- NaN
      varComment <- "Geo. CV converted to geo. SD"
    } else if(tolower(varType)=="arith. cv"){
      varType <- "ArithmeticStdDev"
      if(varUnit=="%"){
        varDim <- getDimension(avgUnit,eachID)
        varValue <- varValue/100
      } else if(varUnit==""){
        varDim <- getDimension(avgUnit,eachID)
      } else{
        stop("Unknown unit for ",varType," for ",eachID)
      }
      varUnit <- avgUnit
      varValue <- avgValue[!is.na(avgValue)] * varValue
      varValue[is.na(varValue) | varValue<0] <- NaN
      varComment <- "Arith. CV converted to arith. SD"
    } else{
      varType <- ""
      varDim <- ""
      varUnit <- ""
      varValue <- rep(NaN,length(varValue))
      varComment <- "."
      warning("Unknown variance type: ",varType," for ",eachID, ". Values ignored.")
    }
    varValue[is.na(avgValue) | is.na(varValue)] <- NaN
  }
  return(list(value=varValue,unit=varUnit,dim=varDim,type=varType,comment=varComment))
}



getDimension <- function(unit,eachID){
  unit <- tolower(unit)
  if(is.element(trimws(tolower(unit),which = "both"),tolower(c("mg/l","µg/l","ng/l","pg/l","mg/dl","µg/dl","ng/dl","pg/dl","mg/ml","µg/ml","ng/ml","pg/ml")))){
    return("Concentration (mass)")
  } else if(is.element(trimws(tolower(unit),which = "both"),tolower(c("mmol/l","µmol/l","nmol/l","pmol/l","mmol/dl","µmol/dl","nmol/dl","pmol/dl","mmol/ml","µmol/ml","nmol/ml","pmol/ml")))){
    return("Concentration (molar)")
  } else if(is.element(trimws(tolower(unit),which = "both"),tolower(c("","%")))){
    return("Fraction")
  } else if(is.element(trimws(tolower(unit),which = "both"),tolower(c("h","min","s","day(s)","week(s)","month(s)","year(s)")))){
    return("Time")
  } else if(is.element(trimws(tolower(unit),which = "both"),tolower(c("mmol","µmol","nmol","pmol")))){
    return("Amount")
  }else if(is.element(trimws(tolower(unit),which = "both"),tolower(c("mg","µg","ng","pg")))){
    return("Mass")
  } else {
    stop("Unknown Dimension"," for ",eachID)
  }
}

getCompartment <- function(compartment="plasma",eachID){
  temp <- unlist(strsplit(compartment,"|",fixed = TRUE))
  if(length(temp)==2){
    compartment <- temp[2]
  } else if(length(temp)==1){
    compartment <- temp[1] 
  } else{
    stop("Too many entries for organ|compartment: ",compartment," for ",eachID)
  }
  
  if(tolower(compartment)=="plasma"){
    return("Plasma")
  } else if(tolower(compartment)=="arterial plasma"){
    return("Plasma")
  } else if(tolower(compartment)=="whole blood"){
    return("Whole Blood")
  } else if(tolower(compartment)=="serum"){
    return("Plasma")
  } else if(tolower(compartment)=="urine"){
    return("Urine")
  } else if(tolower(compartment)=="feces"){
    return("Feces")
  } else if(tolower(compartment)=="bile"){
    return("Undefined")
  } else if(tolower(compartment)=="fraction"){
    return("Undefined")
  } else if(is.element(compartment,c("Plasma","Blood Cells","Interstitial","Intracellular"))){
    return(compartment)
  } else {
    return("Undefined")
    #stop("Unknown Compartment: ",compartment," for ",eachID)
  }
}

getOrgan <- function(compartment="plasma",eachID){
  temp <- unlist(strsplit(compartment,"|",fixed = TRUE))
  if(length(temp)<=2){
    compartment <- temp[1]
  } else{
    stop("Too many entries for organ|compartment: ",compartment," for ",eachID)
  }
  
  if(tolower(compartment)=="plasma"){
    return("Peripheral Venous Blood")
  } else if(tolower(compartment)=="arterial plasma"){
    return("Arterial Blood")
  } else if(tolower(compartment)=="whole blood"){
    return("Peripheral Venous Blood")
  } else if(tolower(compartment)=="serum"){
    return("Peripheral Venous Blood")
  } else if(tolower(compartment)=="urine"){
    return("Kidney")
  } else if(tolower(compartment)=="feces"){
    return("Lumen")
  } else if(tolower(compartment)=="bile"){
    return("Liver")
  } else if(is.element(compartment,c("Bone","Brain","Fat","Gonads","Heart","Kidney",
                                     "Large Intestine","Liver","Lung","Muscle","Pancreas",
                                     "Saliva","Skin","Small Intestine","Spleen","Stomach",
                                     "Duodenum","UpperJejunum","LowerJejunum","UpperIleum","LowerIleum",
                                     "Cecum","Colon Ascendens","Colon Transversum","Colon Descendens","Colon Sigmoid","Rectum"))){
    return(compartment)
  } else {
    return("Undefined")
    #stop("Unknown Compartment: ",compartment," for ",eachID)
  }
}

getAuxiliaryType <- function(type,eachID){
  if(tolower(type)=="arithmeticmean"){
    return("Undefined")
  } else if(tolower(type)=="geometricmean"){
    return("Undefined")
  } else if(tolower(type)=="arithmeticstddev"){
    return("ArithmeticStdDev")
  } else if(tolower(type)=="geometricstddev"){
    return("GeometricStdDev")
  } else if(tolower(type)=="individual"){
    return("Undefined")
  } else if(tolower(type)=="median"){
    return("Undefined")
  } else{
    warning("Unknown type: no auxilary type could be determined for ",type," for ",eachID)
    return("Undefined")
  }
}


getUnitFactorToBaseUnit <- function(unit,eachID){
  unit <- tolower(unit)
  dimension <- getDimension(unit,eachID)
  if(dimension=="Concentration (mass)"){
    unitFactor <- switch(unit, "mg/l"={1e-6},"µg/l"={1e-9},"ng/l"={1e-12},"pg/l"={1e-15},
                         "mg/dl"={1e-5},"µg/dl"={1e-8},"ng/dl"={1e-11},"pg/dl"={1e-14},
                         "mg/ml"={1e-3},"µg/ml"={1e-6},"ng/ml"={1e-9},"pg/ml"={1e-12},NA)
  } else if(dimension=="Concentration (molar)"){
    unitFactor <- switch(unit, "mmol/l"={1e3},"µmol/l"={1},"nmol/l"={1e-3},"pmol/l"={1e-6},
                         "mmol/dl"={1e4},"µmol/dl"={1e1},"nmol/dl"={1e-2},"pmol/dl"={1e-5},
                         "mmol/ml"={1e06},"µmol/ml"={1e3},"nmol/ml"={1},"pmol/ml"={1e-3},NA)
  } else if(dimension=="Mass"){
    unitFactor <- switch(unit, "mg"={1e-6},"µg"={1e-9},"ng"={1e-12},"pg"={1e-15},NA)
  } else if(dimension=="Amount"){
    unitFactor <- switch(unit, "mmol/l"={1e3},"µmol/l"={1},"nmol/l"={1e-3},"pmol/l"={1e-6},NA)
  } else if(dimension=="Fraction"){
    unitFactor <- switch(unit, "%"={1e-2},{1})
  } else{
    stop("Unknown Dimension"," for ",eachID)
  }
  return(unitFactor)
}
