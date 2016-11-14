Private Sub CommandButton1_Click()
Dim sPath As String, sName As String
Dim bk As Workbook, r As Range
Dim r1 As Range, sh As Worksheet
Set BaseWks = Workbooks.Add(xlWBATWorksheet).Worksheets(1)
Set sh = BaseWks
'Set sh = ActiveSheet  ' this is the summary sheet
sPath = "C:\Email Attachments\"
sName = Dir(sPath & "*.xls?")
Do While sName <> ""
Set bk = Workbooks.Open(sPath & sName)
Set r = bk.Worksheets(1).UsedRange
Set r1 = sh.Cells(sh.Rows.Count, 1).End(xlUp).Offset(1, 0)
r.Copy
r1.PasteSpecial xlValues
r1.PasteSpecial xlFormats
Application.CutCopyMode = False  'Empty Clipboard
bk.Close SaveChanges:=False
sName = Dir()
Loop
End Sub
