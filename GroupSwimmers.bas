Attribute VB_Name = "GroupSwimmers"
Sub GroupAndCollapseAllSheets()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim groupRange As Range
    Dim currentGroup As String
    Dim firstSheetSkipped As Boolean
    
    firstSheetSkipped = False
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Skip the first worksheet
        If Not firstSheetSkipped Then
            firstSheetSkipped = True
            GoTo SkipSheet
        End If
        
        ' Determine the last row of data in column A of the current worksheet
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Start from the second row (assuming headers are in the first row)
        For i = 19 To lastRow
            ' Check if the value in column B is empty
            If ws.Cells(i, 2).Value = "" Then
                ' If it's empty, collapse the previous group (if any)
                If Not groupRange Is Nothing Then
                    groupRange.Rows.Group
                    Set groupRange = Nothing
                End If
            Else
                ' If it's not empty, add the cell to the current group range
                If groupRange Is Nothing Then
                    Set groupRange = ws.Cells(i, 2)
                Else
                    Set groupRange = Union(groupRange, ws.Cells(i, 2))
                End If
            End If
        Next i
        
        ' Collapse the last group (if any)
        If Not groupRange Is Nothing Then
            groupRange.Rows.Group
        End If
        
        ' Reset groupRange for the next worksheet
        Set groupRange = Nothing
        
SkipSheet:
    Next ws
End Sub

