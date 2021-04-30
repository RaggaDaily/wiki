# Set Path to Source Excel-File
$filename = 'path/to/excelSourceFile.xlsx'

# Set Parameters for PDF export File
$xlFixedFormat = “Microsoft.Office.Interop.Excel.xlFixedFormatType” -as [type]
$xlQuality = "Microsoft.Office.Interop.Excel.xlQualityStandard" -as [type]

# Start Pagenumber
$xlFromPage = 1
# Last Pagenumber
$xlToPage = 7 
# Sleeping time during script execution
$sleepTime = 2

# Create Excel-Application Object
$objExcel = New-Object -ComObject Excel.Application

# If you execute the script and you want to make Excel visible during the execution
$objExcel.Visible = $true
# Display Errors if you want
$objExcel.DisplayAlerts = $false

# Set target path for the pdf export
$pdfFilepath='path/to/excelTargetPdfFile.pdf'

# Open the workbook
$workBook = $objExcel.Workbooks.Open($filename)

# Wait for x seconds then update the spreadsheet
Start-Sleep -s $sleepTime

# Focus on the first row in spreadsheet of the "Data" worksheet
# Note: This is only for visibility, it does not affect the data refresh
$workSheet = $workBook.Sheets.Item("sheet1")
$workSheet.Select()

#Refresh all data in this workbook. This means, that all data-source for pivot-tables and pivot charts will be refresh 
$workBook.RefreshAll()

# Wait for x seconds then update the spreadsheet
Start-Sleep -s $sleepTime

# Export Excel file to pdf with the parameters in the begin of the script
$workbook.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $pdfFilepath, $xlQuality,$false,$false, $xlFromPage, $xlToPage)

# Save any changes done by the refresh
$objExcel.ActiveWorkbook.Save()
$workBook.Saved = $true

# Write to command line that the workbook was saved
write-host "saving $filename"

# Close Workbooks
$objExcel.Workbooks.close()

# Exit Excel
$objExcel.Quit()
$objExcel = $null

# Write to host that the spreadsheet is updated.
write-host "Finished updating the spreadsheet" -foregroundcolor "green"
