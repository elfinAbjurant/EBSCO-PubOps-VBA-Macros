Attribute VB_Name = "Module3"
Sub purge_packages()
    ' ================================================================
    ' purge_packages
    ' ---------------------------------------------------------------
    ' Deletes any rows in the active sheet that contain the value "P"
    ' in column G.
    ' ================================================================
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ActiveSheet
    
    ' Find the last used row in column G
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).row
    
    ' Loop upward to safely delete matching rows
    For i = lastRow To 1 Step -1
        If ws.Cells(i, "G").Value = "P" Then
            ws.Rows(i).Delete
        End If
    Next i

End Sub

