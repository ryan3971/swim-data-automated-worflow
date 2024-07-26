Attribute VB_Name = "SplitData"
Sub SplitDataBySource()
    Dim sourceCol As Range
    Dim sourceCell As Range
    Dim lastRow As Long
    Dim wsSource As Worksheet
    Dim wsTemplate As Worksheet
    Dim wsNew As Worksheet
    Dim templatePath As String
    Dim headerRow As Range
    Dim copyRange As Range
    Dim targetRange As Range
    Dim sourceData As Variant
    Dim i As Long, j As Long
    Dim dataExists As Boolean
    
    ' Define the source column and last row
    Set wsSource = ThisWorkbook.Sheets("Meet Results") ' Change "Sheet1" to your source sheet name
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    Set sourceCol = wsSource.Range("A2:A" & lastRow) ' Assuming data starts from row 2
    
    ' Path to your template sheet
    templatePath = "C:\Users\rbt7r\OneDrive\Desktop\Adult Life\Work\Contract Work\template_metrics.xlsx" ' Change this to your template file path
    
    ' Loop through each unique source in the source column
    For Each sourceCell In sourceCol.SpecialCells(xlCellTypeVisible).Cells
        ' Check if a worksheet with this source name already exists
        If WorksheetExists(sourceCell.Value) Then
            Set wsNew = ThisWorkbook.Sheets(sourceCell.Value)
        Else
            ' Copy the template sheet to create a new sheet
            Set wsTemplate = Workbooks.Open(templatePath).Sheets("template_sheet") ' Change "Template" to your template sheet name
            wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            Set wsNew = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            wsNew.Name = sourceCell.Value
            wsTemplate.Parent.Close False
            
            ' Copy header row from source sheet to the new sheet (row 18)
            Set headerRow = wsSource.Rows(1)
            Set copyRange = wsSource.Range(wsSource.Cells(1, 2), wsSource.Cells(1, 6)) ' Columns B:F
            Set targetRange = wsNew.Range("A18")
            copyRange.Copy Destination:=targetRange
        End If
        
        ' Check if the data already exists in the new sheet
        dataExists = False
        sourceData = wsSource.Range(wsSource.Cells(sourceCell.Row, 2), wsSource.Cells(sourceCell.Row, 6)).Value
        lastRow = wsNew.Cells(wsNew.Rows.Count, "A").End(xlUp).Row
        If lastRow >= 19 Then
            For i = 19 To lastRow
                dataExists = True
                For j = 1 To UBound(sourceData, 2)
                    If wsNew.Cells(i, j).Value <> sourceData(1, j) Then
                        dataExists = False
                        Exit For
                    End If
                Next j
                If dataExists Then Exit For
            Next i
        End If
        
        ' Copy selected columns data from source sheet to the new sheet (starting at row 19) if it does not already exist
        If Not dataExists Then
            Set copyRange = wsSource.Range(wsSource.Cells(sourceCell.Row, 2), wsSource.Cells(sourceCell.Row, 6)) ' Columns B:D
            Set targetRange = wsNew.Range("A" & wsNew.Cells(wsNew.Rows.Count, "A").End(xlUp).Row + 1)
            
            ' Ensure data starts at row 19
            If targetRange.Row < 19 Then
                Set targetRange = wsNew.Range("A19")
            End If
            
            copyRange.Copy Destination:=targetRange
        End If
    Next sourceCell
End Sub

Function WorksheetExists(wsName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(wsName)
    On Error GoTo 0
    WorksheetExists = Not ws Is Nothing
End Function

