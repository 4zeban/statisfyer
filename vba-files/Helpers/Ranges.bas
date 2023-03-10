Attribute VB_Name = "Ranges"
'namespace=vba-files\Helpers

Public Sub ApplyPatternColor(ByVal range As Range, ByVal color As Variant, Optional ByVal entireRow As Boolean = True)
    On Error Resume Next
    Dim r as Range
    For Each r In range
        if entireRow Then
            r.EntireRow.Interior.PatternColor = color
        Else
            r.Interior.PatternColor = color
        End If
        DoEvents
    Next r
    On Error GoTo 0
End sub

Public Sub ApplyPattern(ByVal range As Range, Optional ByVal pattern As XlPattern = xlPatternNone)
    On Error Resume Next
    Dim r as Range
    For Each r In range
        r.EntireRow.Interior.Pattern = pattern
        DoEvents
    Next r
    On Error GoTo 0
End sub
