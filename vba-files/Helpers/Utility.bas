Attribute VB_Name = "Utility"
'namespace=vba-files\Helpers

Sub SearchInAllSheets(ByVal searchValue As Variant)   
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        
        On Error Resume Next
        Dim foundCell As range
        Set foundCell = ws.Columns(1).Cells.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
        If Not foundCell Is Nothing And GetName(foundCell.offset(0,1)) = "" Then
            ws.Activate
            foundCell.offset(0,1).Select
            On Error GoTo 0
            Exit For
        End If

    Next ws
    On Error GoTo 0
End Sub

Function RemoveSpecialChars(ByVal str As String) As String
    Dim i As Integer
    Dim newStr As String
    Dim allowedChars As String

    allowedChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZÅÄÖabcdefghijklmnopqrstuvwxyzåäö "

    For i = 1 To Len(str)
        If InStr(allowedChars, Mid(str, i, 1)) > 0 Then
            newStr = newStr & Mid(str, i, 1)
        End If
    Next i

    RemoveSpecialChars = newStr
End Function

Function GetKeyValues(ByVal key As String, Optional rowNumber As Long = 1) As Variant
    Dim ws As Worksheet
    Dim valuesString As String
    Set ws = ThisWorkbook.Sheets("_key-values")
    Dim Values As Variant
    Values = Array("", "", "", "", "")

    Set keyCell = ws.Rows(rowNumber).Find(key, LookIn:=xlValues, LookAt:=xlWhole)

    For i = LBound(Values) To UBound(Values)
        If Not keyCell Is Nothing Then
            Values(i) = keyCell.offset(i + 1, 0).Value
        End If
    Next i

    GetKeyValues = Values
End Function

Sub SetKeyValues(ByVal key As String, ByVal Values As Variant, Optional rowNumber As Long = 1)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("_key-values")
    Dim keyCell As Range

    Set keyCell = ws.Rows(rowNumber).Find(key, LookIn:=xlValues, LookAt:=xlWhole)
    If keyCell Is Nothing Then
        For i = 1 To 500
            If ws.Rows(rowNumber).Cells(1).offset(0, i).Value = "" Then
                Set keyCell = ws.Rows(rowNumber).Cells(1).offset(0, i)
                Exit For
            End If
        Next i
        keyCell.Value = key
    End If
    For i = 0 To UBound(Values)
        keyCell.offset(i + 1, 0).Value = Values(i)
    Next i
End Sub

Function GetRowValuesFromColumn(cell As Range, Optional StartColumn As Long = 4) As Variant
    Dim RowValues() As Variant
    Dim i As Long

    ReDim RowValues(30)

    For i = 0 To 30
        RowValues(i) = cell.offset(0, i + StartColumn).Value
    Next i

    GetRowValuesFromColumn = RowValues
End Function

Function GetOffsetForYear(ByVal year As String) As Integer
    Dim searchRange As Range
    Dim foundCell As Range

    Set searchRange = Range("_mappings!A2:A40")
    Set foundCell = searchRange.Find(year, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        GetOffsetForYear = foundCell.offset(0, 1).Value
    Else
        MsgBox "Offset for Year '" & year & "' not found :("
    End If
End Function

Function GetHeadersForYear(ByVal year As String) As Variant
    Dim searchRange As Range
    Dim foundCell As Range
    Dim ret As Range
    Set searchRange = Range("_mappings!A2:A40")
    Set foundCell = searchRange.Find(year, LookIn:=xlValues, LookAt:=xlWhole)

    '=_mappings!E4:V4
    If Not foundCell Is Nothing Then
        GetHeadersForYear = GetRowValuesFromColumn(foundCell)
    Else
        MsgBox "Headers for Year '" & year & "' not found :("
    End If
End Function

Function GetLastWord(ByVal sText As String) As String
    Dim arrWords() As String
    arrWords = Split(sText, " ")
    GetLastWord = arrWords(UBound(arrWords))
End Function

Function CapitalizeAfterSpace(ByVal str As String) As String
    Dim words() As String
    words = Split(str, " ")
    For i = 0 To UBound(words)
        If Not words(i) = "" Then
            words(i) = UCase(Left(words(i), 1)) & Right(words(i), Len(words(i)) - 1)
        End If
    Next i
    CapitalizeAfterSpace = Join(words, "")
End Function
