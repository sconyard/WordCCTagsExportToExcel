Sub getWordData()
Dim wdApp As New Word.Application
Dim myDoc As Word.Document
Dim CCtl As Word.ContentControl
Dim myFolder As String, strFile As String
Dim myWkSht As Worksheet, i As Long, j As Long

myFolder = "Word Document Repository"
Application.ScreenUpdating = False

If myFolder = "" Then Exit Sub
Set myWkSht = ActiveSheet
ActiveSheet.Cells.Clear

i = myWkSht.Cells(myWkSht.Rows.Count, 1).End(xlUp).Row
strFile = Dir(myFolder & "\*.docx", vbNormal)

While strFile <> ""
i = i + 1

Set myDoc = wdApp.Documents.Open(Filename:=myFolder & "\" & strFile, AddToRecentFiles:=False, Visible:=False)

With myDoc
j = 0
For Each CCtl In .ContentControls
If CCtl.Tag = "Export" Then
j = j + 1
myWkSht.Cells(i, j) = CCtl.Range.Text
Else
End If
Next
myWkSht.Columns.AutoFit
End With
myDoc.Close SaveChanges:=False
strFile = Dir()
Wend
wdApp.Quit
Set myDoc = Nothing: Set wdApp = Nothing: Set myWkSht = Nothing
Application.ScreenUpdating = True

End Sub
