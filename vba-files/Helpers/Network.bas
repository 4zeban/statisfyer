Attribute VB_Name = "Network"
'namespace=vba-files\Helpers

Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Function ImportWorkbookFromURL(url As String)
    ' Download the file from the URL
    Dim tempFilePath As String
    tempFilePath = ThisWorkbook.Path & "/" & fso.GetFileName(url)
    URLDownloadToFile 0, url, tempFilePath, 0, 0
    
    ' Import the workbook and add its sheets to the current workbook
    Dim wb As Workbook
    Set wb = Workbooks.Open(tempFilePath)
    
    Dim sheet As Worksheet
    For Each sheet In wb.Sheets
        sheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Next sheet
    
    ' Close the downloaded workbook and delete the temporary file
    wb.Close SaveChanges:=False
    fso.DeleteFile tempFilePath
End Function

Sub TestImportWorkbookFromURL()
    Dim url As String
    url = "https://bra.se/download/18.22a7170813a0d141d2180004866/1371914755403/100La-1997.xls"
    ImportWorkbookFromURL url
End Sub



