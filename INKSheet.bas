'Micro code for sorting by keyword

'Require to rename the SearchBox cell name into >search_string<
'There is additional column name Index(hidden) which contain:
'[@[Location Index]] [@NAME]

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim KeyCells As Range
    Set KeyCells = Range("search_string")
    
    If Not Application.Intersect(KeyCells, Range(Target.Address)) _
           Is Nothing Then
            FastInkFilter (KeyCells.Value)
    End If
End Sub

