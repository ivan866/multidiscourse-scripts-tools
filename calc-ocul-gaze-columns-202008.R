dataset.dir <- 'e:/projects/archive2017/multidiscourse/text/'
setwd(dataset.dir)
pears.ocul.gaze.flist <- readLines(paste(dataset.dir, 'pears-ocul-gaze-list.txt', sep=''))

require(openxlsx)
options(openxlsx.numFmt=NULL,
        openxlsx.dateFormat='yyyy-mm-dd', openxlsx.datetimeFormat='mm:ss,sss',
        openxlsx.borderColour='#c0')
#openxlsx.timeFormat='mm:ss'

fname <- pears.ocul.gaze.flist[1]
df <- read.xlsx(fname, sheet=1,
                startRow=1, colNames=FALSE, rowNames=FALSE,
                detectDates=TRUE, check.names=FALSE, na.strings='',
                skipEmptyRows=FALSE, skipEmptyCols=FALSE)

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