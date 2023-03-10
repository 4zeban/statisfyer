Attribute VB_Name = "taggingController"
'namespace=vba-files\Controllers

Public Enum Values
    MappedKey = 0
    mappedPrefix = 1
End Enum

Public Sub CheckTagGenForRow()
Attribute CheckTagGenForRow.VB_ProcData.VB_Invoke_Func = "g\n14"
    MsgBox ActiveSheet.Rows(ActiveCell.row).Cells(1).Value & vbCrLf & BuildTagForRow(ActiveCell.row, False, False, False)
End Sub

Public Sub FindNextNonTagged()
Attribute FindNextNonTagged.VB_ProcData.VB_Invoke_Func = "n\n14"
    Utility.SearchInAllSheets(ActiveCell.Value)
End Sub

Public Sub TagSelection()
Attribute TagSelection.VB_ProcData.VB_Invoke_Func = "t\n14"
    Call TagRows
End Sub

Public Sub TagSelection_YOLO()
Attribute TagSelection_YOLO.VB_ProcData.VB_Invoke_Func = "y\n14"
   Call TagRows(True)
End Sub

Public Sub MarkTaggedRows()
    Dim rows As Range
    Set rows = GetRowsBetweenValues("SAMTLIGA BROTT", "�vriga f�rfattningar")
    Call MarkTaggedRange(rows)
End Sub

Public Sub UnMarkRange()
    Dim rows As Range
    Set rows = GetRowsBetweenValues("SAMTLIGA BROTT", "�vriga f�rfattningar")
    Call ResetPattern(rows)
End Sub

Private Sub MarkTaggedRange(ByVal rng As Range)
    Dim cell as Range 
    For each cell in rng
        If GetName(cell) = "" Then
            Call Ranges.ApplyPattern(cell)
        Else
            Call Ranges.ApplyPattern(cell, xlPatternChecker)
            Call Ranges.ApplyPatternColor(cell, RGB(0,204,0))
        End if
    Next cell
End Sub

Private Function GetRowsBetweenValues(ByVal startingValue As String, ByVal endingValue As String) As Range
    Dim startRow as Range
    Dim lastRow as Range
    
    Set startRow = ActiveSheet.Columns(1).Cells.Find(startingValue, LookIn:=xlValues, LookAt:=xlWhole)
    Set endRow = ActiveSheet.Columns(1).Cells.Find(endingValue, LookIn:=xlValues, LookAt:=xlWhole)
   
    If startRow is Nothing Then
        MsgBox "Starting row '" & startingValue & "' not found :/"
        End 
    End If

    If endRow is Nothing Then
        MsgBox "Ending row '" & endingValue & "' not found :/"
        End
    End If

    Set GetRowsBetweenValues = Range(startRow.offset(0,1), endRow.offset(0,1))
End Function
    
Private Sub ResetPattern(ByVal range As Range)
    Dim cell as Range 
    For each cell in range
        Call Ranges.ApplyPattern(cell)
    Next cell
End Sub

Private Sub TagRows(Optional ByVal YOLO As Boolean = False)

    Dim tag As String
    Dim selectedRange As Range
    Dim currentRow As Range
    Set selectedRange = Selection
    Dim originalStyle As Variant
    originalStyle = ActiveCell.EntireRow.Interior.Color

    Call Ranges.ApplyPattern(selectedRange.Rows, xlPatternChecker)

    For Each currentRow In selectedRange.Rows
        Call Ranges.ApplyPatternColor(currentRow, RGB(204, 255, 204))
        If Not YOLO Then
            tag = BuildTagForRow(currentRow.row, True, False)
        Else
            tag = BuildTagForRow(currentRow.row, False, False, False)
        End If
        Call TagRow(currentRow.row, tag)
        Call Ranges.ApplyPatternColor(currentRow, RGB(0, 255, 0))
    Next currentRow

    'Call Ranges.ApplyPattern(selectedRange.Rows)

End Sub

Private Sub TagRow(ByVal rowNumber As Long, ByVal tag As String)

    Dim year As String
    Dim year_offset As Integer
    Dim headers As Variant
    Dim offset As Long
    
    Application.ScreenUpdating = False
        
    year = GetLastWord(ActiveSheet.name)
    year_offset = GetOffsetForYear(year)
    headers = GetHeadersForYear(year)
    offset = 0
    
    Set root = ActiveSheet.Cells(rowNumber, 2)
    
    Dim cc As range

    For i = LBound(headers) To UBound(headers)
        Set cc = root.offset(0, offset)
        If Len(headers(i)) > 0 Then
            If Len(cc.Value) > 0 Then
                cc.name = tag & "_" & headers(i)
            Else
                MsgBox tag & "_" & headers(i) & " not set - found no value for " & offset
            End If
            offset = offset + year_offset
        End If
    Next i
    
    Application.ScreenUpdating = True
    DoEvents
End Sub

Function GetName(Byval cell As Range) As String 
    On Error Resume Next
        GetName = cell.name
        Exit Function
    On Error Goto 0
        GetName = ""
End Function
    
Private Function FormatKey(ByVal key As String) As String
    
    key = Replace(key, "kap.", "")
    key = Replace(key, "p.", "")
    
    For i = 1 To 12
        key = Replace(key, i & "a", "")
        key = Replace(key, i & " a", "")
        key = Replace(key, i & "b", "")
        key = Replace(key, i & "c", "")
        key = Replace(key, i & "d", "")
        key = Replace(key, i & " p", "")
    Next i

    key = RemoveSpecialChars(key)
    key = CapitalizeAfterSpace(key)

    FormatKey = key
    
End Function

Private Sub ClearRows()
    Dim selectedRange As Range
    Dim cell As Range
    For Each cell In Selection.Cells
        Call DeleteNameFromCell(cell)
    Next cell
End Sub

Private Sub DeleteNameFromCell(ByVal cell As Range)
    On Error Resume Next ' ignore errors if name doesn't exist
    ActiveWorkbook(cell.name.name).Delete
    On Error GoTo 0 ' stop ignoring errors
End Sub

Private Function BuildTagForActiveRow(Optional ByVal confirmMappings As Boolean = True, Optional ByVal writeValues As Boolean = True, Optional ByVal confirmWriteValues As Boolean = True) As String
    Call BuildTagForRow(ActiveCell.row, confirmMappings, writeValues, confirmWriteValues)
End Function

Private Function BuildTagForRow(ByVal rowNumber As Long, Optional ByVal confirmMappings As Boolean = True, Optional ByVal writeValues As Boolean = True, Optional ByVal confirmWriteValues As Boolean = True) As String

    If InStr(1, ActiveSheet.name, "420") = 0 Then
        MsgBox "Active sheet not 420!"
        End
    End If
    
    Dim key As String
    Dim keyValues As Variant
    Dim parentKey As String
    Dim keyCell As Range
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Set keyCell = ws.Rows(rowNumber).Cells(1)
    
     If keyCell.IndentLevel > 0 Then
        For i = 1 To 25
            If keyCell.offset(i * -1, 0).IndentLevel = 0 Then
                parentKey = FormatKey(keyCell.offset(i * -1, 0).Value) & "_"
                Exit For
            End If
        Next i
     End If
     
    key = parentKey & FormatKey(keyCell.Value)
    keyValues = GetKeyValues(key)
    
    If confirmMappings Then
        If keyValues(Values.mappedPrefix) = "" Then
            keyValues(Values.mappedPrefix) = InputBox("Prefix Mapping for " & keyCell.Value, "No mapped prefix found in values. Prefix verification required:", [LastPrefix])
        Else
            keyValues(Values.mappedPrefix) = InputBox("Prefix Mapping for " & keyCell.Value, "Prefix verification required:", keyValues(Values.mappedPrefix))
        End If
         If keyValues(Values.MappedKey) = "" Then
            keyValues(Values.MappedKey) = InputBox("Key Mapping for " & keyCell.Value, "No mapped key found in values. Key verification required:", key)
        Else
            keyValues(Values.MappedKey) = InputBox("Key Mapping for " & keyCell.Value, "Key verification required:", key)
         End If
    Else
         If keyValues(Values.mappedPrefix) = "" Then
            keyValues(Values.mappedPrefix) = InputBox("Prefix Mapping for " & keyCell.Value, "No mapped prefix found in values. Prefix verification required:", [LastPrefix])
        End If
    End If
    
    [LastPrefix] = keyValues(Values.mappedPrefix)
    
    Debug.Print "BuildTagForRow() | row: " & keyCell.row & " | keyCell " & keyCell.Value & " | key: " & key & " |  mappedPrefix: " & keyValues(Values.mappedPrefix) & " |  mappedKey: " & keyValues(Values.MappedKey) & " | lastPrefix: " & [LastPrefix]
    
    Dim tag As String
    tag = Join(Array("T420", keyValues(Values.mappedPrefix), GetLastWord(ActiveSheet.name), keyValues(Values.MappedKey)), "_")
    
    If Right(tag, 1) = "_" Then
        tag = Left(tag, Len(tag) - 1)
    End If
        
    Debug.Print "BuildTagForRow() | row: " & keyCell.row & "�| tag: " & tag
        
    BuildTagForRow = tag
    
    If writeValues Then
        If confirmWriteValues Then
            Dim response As VbMsgBoxResult
            response = MsgBox("Write Values for " & keyCell.Value & " | Tag values: " & tag, vbYesNo, "Write key values?")
             If response = vbYes Then
                Call SetKeyValues(key, keyValues)
            End If
        End If
        Call SetKeyValues(key, keyValues)
    End If
    
End Function

Private Sub SetTags(ByVal rootCell As Range, ByVal tag As String, ByVal suffixes As Variant)
    For i = LBound(suffixes) To UBound(suffixes)
        If Len(suffixes(i)) > 0 Then
            If Len(rootCell.offset(0, offset).Value) > 0 Then
                rootCell.offset(0, offset).name = tag & suffixes(i)
            Else
                MsgBox tag & suffixes(i) & " not set - found no value for " & offset
            End If
            offset = offset + year_offset
        End If
   Next i
End Sub
