Attribute VB_Name = "tagView"
'namespace=vba-files\Views

Public Function TagRef() As Variant
  Application.Volatile
  On Error GoTo ErrorHandler
  TagRef = range(BuildRef()).Value
  Exit Function
ErrorHandler:
    TagRef = "-"
End Function

Public Function BuildRow(ByVal table As String, ByVal pref As String, Optional year As Long = 1995, Optional ByVal key As String = "") As Variant
    Dim row As Variant
    Dim yearValues() As Variant
    
    ReDim yearValues(0 To 35)
    
    On Error Resume Next ' ignore errors if name doesn't exist
       
    For i = 0 To 35
      If (key = "") Then
            yearValues(i) = range("T" & table & "_" & pref & "_" & year + i & "_Summa")
        Else
            yearValues(i) = range("T" & table & "_" & pref & "_" & year + i & "_" & key)
        End If
    Next i
    On Error GoTo 0 ' stop ignoring errors
    
    BuildRow = yearValues
End Function

Public Function BuildRef() As String
  Application.Volatile
  Dim PREFIX As String
  Dim name As String
  Dim year As String

  Set ws = ActiveSheet

  Dim nameString As String
  Dim row As Long
  row = Application.Caller.row

  PREFIX = ws.Cells(row, 1)
  year = ws.Cells(1, Application.Caller.Column)
  tag = ws.Cells(row, 2)
  suffix = ws.Cells(row, 3)

  If PREFIX = "" Or year = "" Then
    BuildRef = ""
  Else
    name = "T420_" & PREFIX & "_" & year
    If Not tag = "" Then
      name = name & "_" & tag
      If Not suffix = "" Then
        name = name & "_" & suffix
      Else
        name = name & "_" & "Summa"
      End If
    Else
      name = name & "_" & "Summa"
    End If
    Debug.Print "BuildRef() | row: " & row & " namestring: " & name
    BuildRef = name
  End If

End Function

Function GetPrefixes(ByVal table As String) As Variant
  Dim names As Variant
  names = GetTNames(table)
  
  Dim nameList As Object
  Set nameList = CreateObject("Scripting.Dictionary")
  
  For i = 0 To UBound(names)
    nameList(Split(names(i), "_")(0)) = 1
  Next i
   
  GetPrefixes = nameList.Keys()
End Function

Function GetTagsForPrefix(ByVal rng As range, ByVal table As String, Optional excludeHeaders As Boolean = False) As Variant

  Dim names As Variant
  names = GetTNames(table)
  
  Dim nameList As Object
  Set nameList = CreateObject("Scripting.Dictionary")
  
  For i = 0 To UBound(names)
    If Split(names(i), "_")(0) = rng.Value Then
        nameList(Join(RemoveNthValueFromArray(Split(names(i), "_"), 0), "_")) = 1
    End If
  Next i
  
  If excludeHeaders Then
    Dim filterList As Object
    Set filterList = CreateObject("Scripting.Dictionary")
    
    Dim tagList As Variant
    tagList = RemoveValuesInRange(nameList.Keys(), range("_mappings!E28:AC28"))
    Dim s As String
    Dim index As Long
        
    For i = 1 To UBound(tagList)
        s = tagList(i)
        
        index = InStr(1, tagList(i), "_")
        If index = 0 Then
            filterList(tagList(i)) = 1
        Else
            If range("_mappings_" & table & "!E28:AC28").Find(What:=Mid(tagList(i), index + 1), LookIn:=xlValues, LookAt:=xlWhole) Is Nothing Then
                filterList(Left(tagList(i), index - 1)) = 1
            End If
        End If

    Next i
    
    GetTagsForPrefix = filterList.Keys()
  Else
    GetTagsForPrefix = nameList.Keys()
  End If
  
End Function

Function GetTextAfterUnderscore(ByVal myString As String) As String
    Dim index As Long
    index = InStr(1, myString, "_")
    
    If index = 0 Then
        GetTextAfterUnderscore = ""
    Else
        GetTextAfterUnderscore = Mid(myString, index + 1)
    End If
End Function


Private Function GetTNames(ByVal table As String, Optional ByVal stripYear As Boolean = True) As Variant
 Dim wb As Workbook
    Dim ws As Worksheet
    Dim nm As name
    Dim nameList As Object
    Set nameList = CreateObject("Scripting.Dictionary")
    Dim sn As Variant
    
    Set wb = ThisWorkbook
    For Each nm In wb.names
        If Left(nm.name, 4) = "T" & table Then
            Dim s As String
            s = Join(RemoveNthValueFromArray(Split(Replace(nm.name, "T" & table & "_", ""), "_"), 1), "_")
            nameList(s) = 1
        End If
    Next nm
    
    GetTNames = nameList.Keys()
End Function

Function RemoveNthValueFromArray(myArray As Variant, n As Long) As Variant
    Dim newArray() As Variant
    ReDim newArray(LBound(myArray) To UBound(myArray) - 1)
    
    Dim i As Long, j As Long
    j = LBound(newArray)
    
    For i = LBound(myArray) To UBound(myArray)
        If i <> n Then
            newArray(j) = myArray(i)
            j = j + 1
        End If
    Next i
    
    RemoveNthValueFromArray = newArray
End Function


Function RemoveValuesInRange(myArray As Variant, myRange As range) As Variant
    Dim i As Long, j As Long
    Dim newArray() As Variant
    Dim tempArray() As Variant

    ReDim newArray(LBound(myArray) To UBound(myArray))
    
    For i = LBound(myArray) To UBound(myArray)
        Dim found As Boolean
        found = False
        
        
        For j = 1 To myRange.Cells.Count
            If myArray(i) = myRange.Cells(j).Value Then
                found = True
                Exit For
            End If
        Next j
        
        If Not found Then
            newArray(i) = myArray(i)
        End If
    Next i
    
    Dim k As Long, n As Long
    n = 0
    
    For k = LBound(newArray) To UBound(newArray)
        If Not IsEmpty(newArray(k)) Then
            n = n + 1
            ReDim Preserve tempArray(1 To n)
            tempArray(n) = newArray(k)
        End If
    Next k
    
    RemoveValuesInRange = tempArray
End Function

