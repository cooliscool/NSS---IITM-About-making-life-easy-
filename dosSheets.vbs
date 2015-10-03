Dim i As Integer
Dim sheetNum As Integer
Dim beginRow As Integer
Dim endRow As Integer
Dim incr As Integer
Dim column As String

Sub dos()
'
' Specify Column of house list in 'column'
' Specify Starting Row number and End row number in respective variables
' Specify increment - the required number of rows in each page 
'
    beginRow = 3
    endRow = 43
    incr = 10
    column = "B"
    For i = beginRow To endRow Step incr
    
        j = i + 9
        sheetNum = (i - 3) / 10 + 1
        Sheets("Sheet2").Select
        Range(column & i & ":" & column & i + 9).Select
        Selection.Copy
        Sheets("Sheet1").Select
        Range("A2:F2") = "Sheet - " & sheetNum
        Range("A4").Select
        ActiveSheet.Paste
        Call PDFActiveSheet(sheetNum)
        
    Next i
    
End Sub

Function PDFActiveSheet(num As Integer)
Dim ws As Worksheet
Dim strPath As String
Dim myFile As Variant
Dim strFile As String
Dim folderName As String
folderName = "dos_she"
On Error GoTo errHandler

Set ws = ActiveSheet

'enter name and select folder for file
' start in current workbook folder
strFile = num & ".pdf"
strFile = ThisWorkbook.Path & "\" & folderName & "\" & strFile

If myFile <> "False" Then
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=strFile, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

End If

exitHandler:
    Exit Function
errHandler:
    MsgBox "Could not create PDF file"
    Resume exitHandler
End Function


