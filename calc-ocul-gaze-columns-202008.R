Sys.setenv(JAVA_HOME="c:/Program Files/Java/jdk-11.0.1/")
dataset.dir <- 'e:/projects/archive2017/multidiscourse/text/'
pears.ocul.gaze.flist <- readLines(paste(dataset.dir, 'pears-ocul-gaze-list.txt', sep=''))
setwd(dataset.dir)
#log.fname <- file("_report.log", open='wt')
new.sheet.name <- 'calc-ocul-gaze'

#FIXME R CMD BATCH calc-ocul-gaze-columns-202008.R
#NOTE single-runtime version; R runtime gets a filelist, and runs a for loop
#  contrary to multiple-runtime, when multiple R instances are called from .bat file in parallel
require(XLConnect)
#NOTE milli-second accuracy; very important option for time parsing
#NOTE  may also use %OS3 format specifier instead
#NOTE options -> encoding='UTF-8'
options(digits.secs=3, scipen=15)
options(tibble.pillar.subtle=FALSE, tibble.pillar.sigfig=9, tibble.pillar.min_title_chars=10)
#options(XLConnect.dateTimeFormat='%M:%OS')

print(paste('#### excel file processing script; pears-ocul-gaze; 202008 ####', sep=''))
print(paste('#### ', Sys.Date(), ' ####', sep=''))
for(fname in pears.ocul.gaze.flist) {
  t.start <- Sys.time()
  print(paste('#### ', format(t.start, '%T %z'), ' ####', sep=''))
  wb <- loadWorkbook(fname, create=FALSE, password=NULL)
  print(summary(wb))
  #NOTE we treat custom time format as character, then manually convert to s.f format (like in pandas)
  #  then, on saving, format as human-readable MM:SS.000 again
  #FIXME col B to strptime
  #NOTE trim newlines, #SEE python same problem occured
  #read
  df = readWorksheet(wb, sheet=1, header=FALSE, rownames=NULL,
                     colTypes=c('numeric', 'character', 'numeric', 'factor', 'factor', 'character'),
                     dateTimeFormat='%M:%OS')
  print('#### sanity check: ####')
  print(paste('col A: all indices are sequential: ', all(diff(df[,1])==1), sep=''))
  #TODO print(paste('col B: all times are rising: ', '#TODO', sep=''))
  print(paste('col D: unique names: ', paste(unique(df[,4]), collapse=','), sep=''))
  print(paste('col E: unique names: ', paste(unique(df[,5]), collapse=','), sep=''))
  #process
  col.B <- strptime(df[,2], format='%M:%OS', tz='Europe/Moscow')
  col.D.change.ind <- c(1, which(diff(as.numeric(as.factor(df[,4])))!=0)+1)
  col.D.change.ind.second <- c(col.D.change.ind[-1]-1, nrow(df))
  ind.seq <- seq_along(col.D.change.ind)
  df[col.D.change.ind, 6] <- ind.seq
  df[col.D.change.ind, 7] <- df[col.D.change.ind, 4]
  for(n in ind.seq) {
    df[col.D.change.ind[n], 8] <- sum(df[col.D.change.ind[n]:col.D.change.ind.second[n], 3])
  }
  print('#### all calculation done ####')
  #write
  my.df <- createSheet(wb, name=new.sheet.name)
  #existsSheet()
  #setActiveSheet()
  #setDataFormatForType(wb, type=XLC$"DATA_TYPE.NUMERIC", format="0.0000")
  #setFillForegroundColor(header, color=XLC$COLOR.GREY_25_PERCENT)
  #setSheetPos()
  writeWorksheet(wb, data=df, sheet=new.sheet.name, startRow=1, startCol=1, header=FALSE, rownames=NULL)
  saveWorkbook(wb)
  print('#### workbook updated! ####')
  print(paste('time elapsed: ', format(difftime(Sys.time(), t.start, units=c('secs')), digits=4), sep=''))
}
#sink(log.fname, type=c('output', 'message'), split=TRUE, append=FALSE)
#close(log.fname)
