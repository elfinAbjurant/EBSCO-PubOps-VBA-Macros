Attribute VB_Name = "Module4"
Sub purge_dates()
    ' Deletes rows (on the ActiveSheet) where column V does NOT end with "9999".
    ' Robust approach: loop bottom->top and delete table rows via ListRows.Delete
    ' when the row belongs to a ListObject; otherwise delete the worksheet row.
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim lo As ListObject
    Dim deletedThisRow As Boolean
    
    Set ws = ActiveSheet
    
    ' Find last used row in column V
    lastRow = ws.Cells(ws.Rows.Count, "V").End(xlUp).row
    If lastRow < 2 Then Exit Sub ' nothing to delete (only header or empty)
    
    ' Loop upward to avoid skipping rows after deletion
    For i = lastRow To 2 Step -1
        cellValue = Trim(CStr(ws.Cells(i, "V").Value))
        
        If Right(cellValue, 4) <> "9999" Then
            deletedThisRow = False
            
            ' If the row is inside any table's DataBodyRange, delete via ListRows
            For Each lo In ws.ListObjects
                If Not lo.DataBodyRange Is Nothing Then
                    If Not Intersect(ws.Rows(i), lo.DataBodyRange) Is Nothing Then
                        On Error Resume Next
                        ' Compute the index of the ListRow within the table
                        lo.ListRows(i - lo.HeaderRowRange.row).Delete
                        If Err.Number = 0 Then
                            deletedThisRow = True
                        Else
                            Err.Clear
                        End If
                        On Error GoTo 0
                        Exit For
                    End If
                End If
            Next lo
            
            ' If not part of a table (or table delete failed), delete the worksheet row
            If Not deletedThisRow Then
                On Error Resume Next
                ws.Rows(i).Delete
                ' If deletion fails (rare), clear the row contents as a fallback
                If Err.Number <> 0 Then
                    Err.Clear
                    ws.Rows(i).ClearContents
                End If
                On Error GoTo 0
            End If
        End If
    Next i
End Sub

