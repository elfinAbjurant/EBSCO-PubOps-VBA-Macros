Attribute VB_Name = "Module5"
Sub STATUS_Cleanup()

    Dim tblE As ListObject
    Dim tblT As ListObject
    Dim row As ListRow
    Dim i As Long
    Dim SPID As String
    Dim statusValue As String
    Dim XRMap As Object
    Dim targetDict As Object
    Dim v As Variant  ' For looping status array
    
    ' Tables created by PAC_Sheet_Prep
    Set tblE = Worksheets("EBSCO").ListObjects("EBSCO_Table")
    Set tblT = Worksheets("TBR").ListObjects("TBR_Table")
    
    ' Build dictionary for fully-XR detection
    Set XRMap = CreateObject("Scripting.Dictionary")
    XRMap.CompareMode = vbTextCompare

    ' Pass 1: Determine which SPIDs are fully XR
    For Each row In tblE.ListRows
        SPID = CStr(row.Range.Columns(tblE.ListColumns("SPID NAME").Index).Value)
        statusValue = UCase$(Trim$(row.Range.Columns(tblE.ListColumns("STATUS").Index).Value))

        If Not XRMap.Exists(SPID) Then XRMap.Add SPID, True
        If statusValue <> "XR" Then XRMap(SPID) = False
    Next row

    ' Pass 2: Move rows for SPIDs that are fully XR
    For i = tblE.ListRows.Count To 1 Step -1
        Set row = tblE.ListRows(i)
        SPID = CStr(row.Range.Columns(tblE.ListColumns("SPID NAME").Index).Value)

        If XRMap.Exists(SPID) Then
            If XRMap(SPID) = True Then
                tblT.ListRows.Add.Range.Value = row.Range.Value
                row.Delete
            End If
        End If
    Next i

    ' Dictionary of statuses that should move individually
    Set targetDict = CreateObject("Scripting.Dictionary")
    targetDict.CompareMode = vbTextCompare
    
    Dim arr
    arr = Array("AU", "CS", "DD", "DU", "MM", "NA", "NC", "OD", "SP", "SQ", "SS")

    ' FIXED: v is Variant, so For Each works
    For Each v In arr
        targetDict(v) = True
    Next v

    ' Pass 3: Move rows with targeted status values
    For i = tblE.ListRows.Count To 1 Step -1
        Set row = tblE.ListRows(i)
        statusValue = Trim$(row.Range.Columns(tblE.ListColumns("STATUS").Index).Value)

        If targetDict.Exists(UCase$(statusValue)) Then
            tblT.ListRows.Add.Range.Value = row.Range.Value
            row.Delete
        End If
    Next i

    MsgBox "STATUS cleanup complete."

End Sub

