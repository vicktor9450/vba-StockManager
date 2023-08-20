'Micro code for sorting by keyword

'Require to rename the SearchBox cell name into >Paper_String<
'There is additional column name Index(hidden) which contain:
'[@Specs] [@Width] for searching
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim KeyCells As Range
    Set KeyCells = Range("Paper_String")
    
    If Not Application.Intersect(KeyCells, Range(Target.Address)) _
           Is Nothing Then
            FastPaperFilter (KeyCells.Value)
    End If
End Sub
