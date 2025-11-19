Attribute VB_Name = "Module2"
Sub PAC_VLOOKUP()
    Dim wsE As Worksheet
    Dim wsP As Worksheet
    Dim colName As String
    Dim lastRow As Long
    
    ' Reference sheets
    Set wsE = ActiveWorkbook.Sheets("EBSCO")
    Set wsP = ActiveWorkbook.Sheets("PUBLISHER")
    
    ' Step 1: Insert two blank columns to the right of column O
    wsE.Columns("P:Q").Insert Shift:=xlToRight
    
    ' Step 2: Ask user what Column Q should be titled
    colName = InputBox("What is the Publisher?", "Column Q Title")
    If colName = "" Then colName = "Publisher"
    wsE.Range("Q1").Value = colName
    
    ' Step 3: Use Text-to-Columns on Column O with "-" delimiter
    ' Output starts at column O so the split parts flow into the new Column P
    wsE.Columns("O:O").TextToColumns Destination:=wsE.Range("O1"), _
        DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, _
        Other:=True, OtherChar:="-"
    
    ' Step 4: Apply formulas
    ' Q2 (EBSCO sheet)
    wsE.Range("Q2").Formula = "=VLOOKUP([@[PUBL PRODUCT CODE]],PUBLISHER!A:A,1,0)"
    
    ' C2 (PUBLISHER sheet)
    wsP.Range("C2").Formula = "=VLOOKUP([@[Journal Code]],EBSCO!O:O,1,0)"
    
    MsgBox "Columns inserted, text split, and formulas applied.", vbInformation
End Sub

