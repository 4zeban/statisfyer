Attribute VB_Name = "Ranges"
'namespace=vba-files\Helpers

Public Sub ApplyPatternColor(ByVal range As range, ByVal color As Variant, Optional ByVal entireRow As Boolean = True)
    On Error Resume Next
    Dim r As range
    For Each r In range
        If entireRow Then
            r.entireRow.Interior.PatternColor = color
        Else
            r.Interior.PatternColor = color
        End If
        DoEvents
    Next r
    On Error GoTo 0
End Sub

Public Sub ApplyPattern(ByVal range As range, Optional ByVal pattern As XlPattern = xlPatternNone)
    On Error Resume Next
    Dim r As range
    For Each r In range
        r.entireRow.Interior.pattern = pattern
        DoEvents
    Next r
    On Error GoTo 0
End Sub
