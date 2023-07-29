'*****************************************************
Sub PivotCreating()

'Disable Screen flashing when running code
Application.ScreenUpdating = False

'Declare Variables
Dim PSheet As Worksheet
Dim DSheet As Worksheet
Dim PCache As PivotCache
Dim PTable As PivotTable
Dim PRange As Range
Dim LastRow As Long
Dim LastCol As Long
Dim Io As ListObject

'Delete existed "PaperSummary" before Launching
On Error Resume Next
Application.DisplayAlerts = False
Worksheets("PaperSummary").Delete
Sheets.Add Before:=ActiveSheet
ActiveSheet.Name = "PaperSummary"
Application.DisplayAlerts = True
Set PSheet = Worksheets("PaperSummary")
Set DSheet = Worksheets("InStock")

'Chosing DataSource
Set Io = DSheet.ListObjects(1)

'Define Data Range
LastCol = Io.ListColumns.Count
LastRow = Io.ListRows.Count
Set PRange = DSheet.Cells(5, 1).Resize(LastRow, LastCol)

'Define Pivot Cache
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(2, 2), _
TableName:="PaperPivotTable")

'Insert Blank Pivot Table
Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(1, 1), TableName:="PaperPivotTable")

    
'Setting things up
    'Insert Column Fields
    With ActiveSheet.PivotTables("PaperPivotTable").PivotFields("Width")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    'Insert Row Fields
    With ActiveSheet.PivotTables("PaperPivotTable").PivotFields("Specs")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    'Insert Data Field
    With ActiveSheet.PivotTables("PaperPivotTable").PivotFields("Remaining")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "Sum of Remaining "
    End With
    
'Polishing PivotTable
With ActiveSheet.PivotTables("PaperPivotTable")
    .DataPivotField.PivotItems("Sum of Remaining ").Caption = "Paper Summary"
    .CompactLayoutColumnHeader = "Width"
    .CompactLayoutRowHeader = "Specs"
    .RowGrand = False
    .ColumnGrand = False
    .MergeLabels = True
    .PivotFields("Slitting").PivotItems("(blank)").Caption = " "
    .ShowTableStyleRowStripes = True
End With
    
    'Delete SubTotal
    Call NoSubtotals
    
    'Select PivotTable and All border
    Call AllBorder
    
    'Minimize col A
    Columns("A:A").ColumnWidth = 0.53
    
    'Add last update timestone B1
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "=LastSavedTimeStamp()"
    Selection.NumberFormat = "mmm/dd/yyyy h:mm;@"
    
    'Freeze Pane
    Range("C5").Select
    ActiveWindow.FreezePanes = True
    
    'Set Zoom ratio
    ActiveWindow.Zoom = 100
    
    'Refresh screen for update calculated data range
    Application.ScreenUpdating = True
    
End Sub

'*****************************************************
Sub NoSubtotals()

Dim pt As PivotTable
Dim pf As PivotField

On Error Resume Next
For Each pt In ActiveSheet.PivotTables
  pt.ManualUpdate = True
  For Each pf In pt.PivotFields
    'First, set index 1 (Automatic) to True,
    'so all other values are set to False
    pf.Subtotals(1) = True
    pf.Subtotals(1) = False
  Next pf
  pt.ManualUpdate = False
Next pt

End Sub

'*****************************************************
Sub AllBorder()

    ActiveSheet.PivotTables("PaperPivotTable").PivotSelect "", xlDataAndLabel, True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
End Sub

'*****************************************************
'For sorting INK table, require add some code at specific (INK) sheet at Microsoft Excel Objects
Sub FastInkFilter(sch As String)
    
    Dim lo As ListObject
    Set lo = ActiveSheet.ListObjects(1)
    LastCol = lo.ListColumns.Count
    
    If lo.AutoFilter.FilterMode Then
        lo.AutoFilter.ShowAllData
        lo.Range.AutoFilter Field:=LastCol, Criteria1:= _
            Array("*" + sch + "*"), Operator:=xlFilterValues
        Else
        lo.Range.AutoFilter Field:=LastCol, Criteria1:= _
            Array("*" + sch + "*"), Operator:=xlFilterValues
        End If
        
    Call ScrollToTop
    Range("search_string").Select
End Sub

'*****************************************************
'For sorting InStock table, require add some code at specific (InStock) sheet at Microsoft Excel Objects
Sub FastPaperFilter(sch As String)
    
    Dim lo As ListObject
    Set lo = ActiveSheet.ListObjects(1)
    LastCol = lo.ListColumns.Count
    
    If lo.AutoFilter.FilterMode Then
        lo.AutoFilter.ShowAllData
        lo.Range.AutoFilter Field:=LastCol, Criteria1:= _
            Array("*" + sch + "*"), Operator:=xlFilterValues
        Else
        lo.Range.AutoFilter Field:=LastCol, Criteria1:= _
            Array("*" + sch + "*"), Operator:=xlFilterValues
        End If
        
     Call ScrollToTop
    Range("Paper_String").Select
End Sub

'*****************************************************
'Adding LastSavedtTimeStamp
Function LastSavedTimeStamp() As Date
  LastSavedTimeStamp = ActiveWorkbook.BuiltinDocumentProperties("Last Save Time")
End Function

'*****************************************************
'This macro scrolls to the top of your spreadsheet
Sub ScrollToTop()
ActiveWindow.ScrollRow = 1 'the row you want to scroll to
End Sub

'*****************************************************
'This macro scrolls to the bottom of your spreadsheet for easily insert data in 'Paper' Sheet
Sub ScrollToBotPaper()

Dim Io As ListObject
Dim LastPaperRow As Long
Dim DSheet As Worksheet

'Choosing DataSheet
Set DSheet = Worksheets("InStock")
'Chosing DataSource
Set Io = DSheet.ListObjects(1)

'Finding the last row of list
LastPaperRow = Io.ListRows.Count

ActiveWindow.ScrollRow = LastPaperRow - 4  'the row you want to scroll to
End Sub
