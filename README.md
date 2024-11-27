What you'll need: 
 uneditied excel running files 

What to do: 
  go to the app (runratrun.streamlit.app) 
  you'll use the upload files option on the app. Select every excel file you want the app to run. 
    (why? well, the use local directory option won't work unless you're running the app driectly from python on your desktop because in the app cloud it doesn't have access to your local directory) 
  press start processing 
  wait.... patience, it takes a while 
You should get 5 exports: active, inactive, hourly active, hourly inactive, and debug. You should only need to download the first 4 unless you're running into problems with the app and then the debug info might be useful. 

To convert the hourly data to the format we want I built a macro for everyone to run in excel. this is what that looks like:

 Sub ReorganizeAndSaveHourlyDataInOneFile()
    Dim ws As Worksheet
    Dim newWB As Workbook
    Dim newWS As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim metrics(1 To 7) As String
    Dim metricIndex As Integer
    Dim ratRow As Long
    Dim dayIndex As Integer
    Dim currentCol As Long, newCol As Long
    Dim metricName As String
    
    ' Define the metrics order (adjust as needed for your dataset)
    metrics(1) = "Total_Bouts"
    metrics(2) = "Minutes_Running"
    metrics(3) = "Total_Wheel_Turns"
    metrics(4) = "Distance_m"
    metrics(5) = "Avg_Distance_per_Bout"
    metrics(6) = "Avg_Bout_Length"
    metrics(7) = "Speed"

    ' Create a new workbook for all reorganized data
    Set newWB = Application.Workbooks.Add

    ' Loop through each worksheet in the original workbook
    For Each ws In ThisWorkbook.Sheets
        ' Create a new worksheet in the new workbook for the current hour
        Set newWS = newWB.Sheets.Add(After:=newWB.Sheets(newWB.Sheets.Count))
        newWS.Name = ws.Name & " Reorganized"

        ' Get the last row and column of the current sheet
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

        ' Copy the Rat column to the new worksheet
        ws.Range("A1:A" & lastRow).Copy Destination:=newWS.Range("A1")
        newWS.Cells(1, 1).Value = "Rat"

        ' Reorganize the data by grouping metrics across days
        newCol = 2 ' Start from column 2 in the new worksheet
        For metricIndex = 1 To UBound(metrics)
            metricName = metrics(metricIndex)
            For dayIndex = 1 To (lastCol - 1) / 7 ' Calculate the number of days
                currentCol = 2 + (dayIndex - 1) * 7 + (metricIndex - 1)
                newWS.Cells(1, newCol).Value = metricName & " Day " & dayIndex
                For ratRow = 2 To lastRow
                    newWS.Cells(ratRow, newCol).Value = ws.Cells(ratRow, currentCol).Value
                Next ratRow
                newCol = newCol + 1
            Next dayIndex
        Next metricIndex
    Next ws

    ' Delete the default empty sheet
    Application.DisplayAlerts = False
    newWB.Sheets(1).Delete
    Application.DisplayAlerts = True

    ' Save the new workbook
    Dim savePath As String
    savePath = Application.GetSaveAsFilename(InitialFileName:="Reorganized_Hourly_Data.xlsx", FileFilter:="Excel Files (*.xlsx), *.xlsx")
    If savePath <> "False" Then
        newWB.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
        MsgBox "Reorganized data saved to: " & savePath
    Else
        MsgBox "Save canceled. Data not saved."
    End If

    ' Close the new workbook
    newWB.Close SaveChanges:=False
 End Sub

You will run that macro by going into visual basic by opening your hourly excel > visual basic > module > paste this macro > back to excel > developer > macro > run macro you just created 
it will create a new excel with the reorganized data 
