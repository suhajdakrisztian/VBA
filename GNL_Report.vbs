'grab dataset from oasys
sub GNL_Report()
	Dim PT As PivotTable
	'open dld file, open destination, paste in dld
	application.screenupdating=false	
	downloaded_dataset=workbooks.open(application.getopenfilename())
	downloaded_dataset.sheets(1).range(cells(),cells()).copy
	reporting_workbook=workbooks.open("FILEPATH")
	reporting_workbook.worksheets("NAME").cells(1,1).pastespecial xlvalues
	reporting_workbook.worksheets("NAME").activate
	For Each PT In ActiveWorkbook.Sheets(1).PivotTables
		PT.RefreshTable
	Next PT
	application.screenupdating=true
end sub