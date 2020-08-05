Sys.setenv(JAVA_HOME="c:/Program Files/Java/jdk-11.0.1/")
dataset.dir <- 'e:/projects/archive2017/multidiscourse/text/'
setwd(dataset.dir)
pears.ocul.gaze.flist <- readLines(paste(dataset.dir, 'pears-ocul-gaze-list.txt', sep=''))

#require(data.table); require(tibble)
require(XLConnect)
require(readxl)
require(openxlsx)
#NOTE milli-second accuracy; very important option for time parsing
#NOTE  may also use %OS3 format specifier instead
options(digits.secs=3, scipen=15)
options(tibble.pillar.subtle=FALSE, tibble.pillar.sigfig=9, tibble.pillar.min_title_chars=10)
options(openxlsx.numFmt=NULL,
        openxlsx.dateFormat='yyyy-mm-dd', openxlsx.datetimeFormat='mm:ss.000',
        openxlsx.borderColour='#ff0000')

fname <- pears.ocul.gaze.flist[1]
print(paste('excel file processing script; pears-ocul-gaze; 202008', sep=''))
print(paste('reading file: ', fname, sep=''))
print(paste('format is: ', excel_format(fname), sep=''))
print(paste('converting time, origin is: ', getDateOrigin(fname), sep=''))
print(paste('file contains sheets: ', excel_sheets(fname), sep=''))

df <- read.xlsx(fname, sheet=1,
                startRow=1, colNames=FALSE, rowNames=FALSE,
                detectDates=FALSE, check.names=FALSE, na.strings='',
                skipEmptyRows=FALSE, skipEmptyCols=FALSE)
#NOTE we treat custom time format as numeric, then manually convert to s.f format (like in pandas)
#  then, on saving, format as human-readable MM:SS.000 again
df = read_excel(fname, sheet=1, col_names=FALSE,
                col_types=c('numeric', 'numeric', 'numeric', 'text', 'text', 'numeric', 'text', 'numeric', 'text'),
                trim_ws=TRUE)
wb <- loadWorkbook(fname)

wb <- write.xlsx(df, file='myexcel.xlsx', asTable=FALSE,
                 startRow=1, startCol=1,
                 rowNames=FALSE, colNames=FALSE, keepNA=FALSE,
                 overwrite=TRUE)
#setColWIdths(wb, sheet=1, cols=1:5, widths=20)
#saveWorkbook(wb, fname, overwrite=TRUE)

#addWorksheet(wb, sheetName='', gridLines=TRUE)
#freezePane(wb, sheet=1, firstRow=TRUE, firstCol=TRUE)
#writeData(wb, sheet=1, x=df, startRow, startCol)
#writeDataTable(wb, sheet=1, x=df, rowNames=TRUE, colnames=TRUE, tableStyle='TableStyleLight9')
#  startRow, startCol
#insertPlot(wb, sheet, xy=c('C', 4))