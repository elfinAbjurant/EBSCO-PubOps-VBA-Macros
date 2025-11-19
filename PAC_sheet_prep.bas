Attribute VB_Name = "Module1"
Sub PAC_sheet_prep()
    ' ================================================================
    ' PAC_sheet_prep
    ' ---------------------------------------------------------------
    ' This macro prepares a workbook for a PAC QC workflow.
    ' It renames the first sheet, creates new sheets, converts data
    ' into tables, copies headers, and builds a publisher reference
    ' key table for later use.
    '
    ' Steps:
    '   1. Rename first sheet to "EBSCO"
    '   2. Delete and recreate sheets "PUBLISHER" and "TBR"
    '   3. Convert data on EBSCO sheet to an Excel table
    '   4. Copy header row from EBSCO to TBR & create a TBR table
    '   5. Create a 3x2 table on PUBLISHER sheet
    '   6. Add a color-coded key on PUBLISHER for reference
    ' ================================================================
    
    Dim ws As Worksheet           ' EBSCO sheet reference
    Dim wsP As Worksheet          ' PUBLISHER sheet reference
    Dim wsTBR As Worksheet        ' TBR sheet reference
    Dim tbl As ListObject         ' Table object for EBSCO
    Dim pubTable As ListObject    ' Table object for PUBLISHER
    Dim tbrTable As ListObject    ' Table object for TBR
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range              ' Range used for table creation
    
    ' ------------------------------------------------
    ' Step 1: Rename the first sheet to "EBSCO"
    ' ------------------------------------------------
    Set ws = ActiveWorkbook.Sheets(1)
    ws.Name = "EBSCO"
    
    ' ------------------------------------------------
    ' Step 2: Add new sheets "PUBLISHER" and "TBR"
    ' ------------------------------------------------
    ' First, delete existing ones if they already exist
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets("PUBLISHER").Delete
    ActiveWorkbook.Sheets("TBR").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Then recreate fresh sheets at the end of the workbook
    ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    ActiveSheet.Name = "PUBLISHER"
    Set wsP = ActiveWorkbook.Sheets("PUBLISHER")
    
    ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    ActiveSheet.Name = "TBR"
    Set wsTBR = ActiveWorkbook.Sheets("TBR")
    
    ' ------------------------------------------------
    ' Step 3: Convert the data on the EBSCO sheet into a table
    ' ------------------------------------------------
    With ws
        ' Identify last used row and column for the data range
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        Set rng = .Range(.Cells(1, 1), .Cells(lastRow, lastCol))
        
        ' Remove any existing tables to prevent conflicts
        For Each tbl In .ListObjects
            tbl.Unlist
        Next tbl
        
        ' Create a new structured table using the data range
        Set tbl = .ListObjects.Add(xlSrcRange, rng, , xlYes)
        tbl.Name = "EBSCO_Table"
        tbl.TableStyle = "TableStyleLight1"  ' White table style
    End With
    
    ' ------------------------------------------------
    ' Step 4: Copy the header row from EBSCO to TBR sheet & create a table
    ' ------------------------------------------------
    ws.Rows(1).Copy Destination:=wsTBR.Rows(1)
    
    With wsTBR
        ' Determine the last used column on the TBR sheet (after copy)
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        
        ' Create the table on the TBR sheet properly from the worksheet
        Set tbrTable = .ListObjects.Add(xlSrcRange, .Range(.Cells(1, 1), .Cells(2, lastCol)), , xlYes)
        tbrTable.Name = "TBR_Table"
        tbrTable.TableStyle = "TableStyleLight1"
    End With
    
    ' ------------------------------------------------
    ' Step 5: Create a small 3x2 table on the PUBLISHER sheet
    ' ------------------------------------------------
    With wsP
        ' Define table headers
        .Range("A1").Value = "Journal Code"
        .Range("B1").Value = "Journal Name"
        .Range("C1").Value = "EBSCO"
        ' Leave second row blank
        .Range("A2:C2").ClearContents
        
        ' Add the new table
        Set pubTable = .ListObjects.Add(xlSrcRange, .Range("A1:C2"), , xlYes)
        pubTable.Name = "PUBLISHER_Table"
        pubTable.TableStyle = "TableStyleLight1"
    End With
    
    ' ------------------------------------------------
    ' Step 6: Add color-coded key for reference values
    ' ------------------------------------------------
    With wsP
        ' Label column for key
        .Range("F3").Value = "KEY"
        .Range("F4").Value = "OKAY"
        
        ' "ADD" status with teal fill (Accent 1, Lighter 80%)
        .Range("F5").Value = "ADD"
        .Range("F5").Interior.Color = RGB(198, 239, 255)
        
        ' "REMOVE" status with red fill
        .Range("F6").Value = "REMOVE"
        .Range("F6").Interior.Color = RGB(255, 0, 0)
        
        ' "UNSURE" status with yellow fill
        .Range("F7").Value = "UNSURE"
        .Range("F7").Interior.Color = RGB(255, 255, 0)
    End With

End Sub
