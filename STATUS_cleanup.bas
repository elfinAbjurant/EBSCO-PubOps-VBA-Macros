Attribute VB_Name = "Module5"
Sub STATUS_Cleanup()

    Dim wsE As Worksheet
    Dim wsT As Worksheet
    Dim tblE As ListObject
    Dim tblT As ListObject
    Dim spidCol As Long
    Dim statusCol As Long
    Dim i As Long
    Dim XRMap As Object
    Dim moveStatuses As Variant
    Dim statusValue As String
    Dim spidValue As String
    Dim row As ListRow

    ' Sheets and tables from PAC_sheet_prep
    Set wsE = Worksheets("EBSCO")
    Set wsT = Worksheets("TBR")
    Set tblE = wsE.ListObjects("EBSCO_Table")
    Set tblT = wsT.ListObjects("TBR_Table")

    ' Column indices
    spidCol = tblE.ListColumns("SPID NAME").Index
    statusCol = tblE.ListColumns("STATUS").Index

    ' -------------------------------
    ' Step 1: Move rows with specific STATUS values
    ' -------------------------------
    moveStatuses = Array("AU", "CS", "DD", "DU", "MM", "NA", "NC", "OD", "SP", "SQ", "SS")
    
    For i = tblE.ListRows.Count To 1 Step -1
        Set row = tblE.ListRows(i)
        statusValue = Trim$(row.Range.Columns(statusCol).Value)
        
        If Not IsError(Application.Match(statusValue, moveStatuses, 0)) Then
            ' Move to TBR
            tblT.ListRows.Add.Range.Value = row.Range.Value
            row.Delete
        End If
    Next i

    ' -------------------------------
    ' Step 2: Move fully-XR SPIDs
    ' -------------------------------
    Set XRMap = CreateObject("Scripting.Dictionary")
    XRMap.CompareMode = vbTextCompare

    ' First pass: check which SPIDs are fully XR
    For Each row In tblE.ListRows
        spidValue = CStr(row.Range.Columns(spidCol).Value)
        statusValue = Trim$(CStr(row.Range.Columns(statusCol).Value))
        
        If Not XRMap.Exists(spidValue) Then XRMap(spidValue) = True
        If statusValue <> "XR" Then XRMap(spidValue) = False
    Next row

    ' Second pass: move rows for fully-XR SPIDs
    For i = tblE.ListRows.Count To 1 Step -1
        Set row = tblE.ListRows(i)
        spidValue = CStr(row.Range.Columns(spidCol).Value)
        
        If XRMap(spidValue) = True Then
            tblT.ListRows.Add.Range.Value = row.Range.Value
            row.Delete
        End If
    Next i

    MsgBox "STATUS cleanup complete."

End Sub

