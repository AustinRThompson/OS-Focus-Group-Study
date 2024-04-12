Attribute VB_Name = "NewMacros"
Sub ExportComments()
' Note: A reference to the Microsoft Excel # Object Library is required, set via Tools|References in the Word VBE.
Dim StrCmt As String, StrTmp As String, i As Long, j As Long, xlApp As Object, xlWkBk As Object
StrCmt = "Page,Line,Author,Date & Time,Comment,Reference Text"
StrCmt = Replace(StrCmt, ",", vbTab)
With ActiveDocument
  ' Process the Comments
  For i = 1 To .Comments.Count
    With .Comments(i)
      StrCmt = StrCmt & vbCr & .Reference.Information(wdActiveEndAdjustedPageNumber) & vbTab
      StrCmt = StrCmt & .Reference.Information(wdFirstCharacterLineNumber) & vbTab & .Author & vbTab
      StrCmt = StrCmt & .Date & vbTab & Replace(Replace(.Range.Text, vbTab, "<TAB>"), vbCr, "<P>")
      StrCmt = StrCmt & vbTab & Replace(Replace(.Scope.Text, vbTab, "<TAB>"), vbCr, "<P>")
    End With
  Next
End With
' Test whether Excel is already running.
On Error Resume Next
Set xlApp = GetObject(, "Excel.Application")
'Start Excel if it isn't running
If xlApp Is Nothing Then
  Set xlApp = CreateObject("Excel.Application")
  If xlApp Is Nothing Then
    MsgBox "Can't start Excel.", vbExclamation
    Exit Sub
  End If
End If
On Error GoTo 0
With xlApp
  Set xlWkBk = .Workbooks.Add
  ' Update the workbook.
  With xlWkBk.Worksheets(1)
    For i = 0 To UBound(Split(StrCmt, vbCr))
      StrTmp = Split(StrCmt, vbCr)(i)
        For j = 0 To UBound(Split(StrTmp, vbTab))
          .Cells(i + 1, j + 1).Value = Split(StrTmp, vbTab)(j)
        Next
    Next
    .Columns("A:D").AutoFit
  End With
  ' Tell the user we're done.
  MsgBox "Workbook updates finished.", vbOKOnly
  ' Switch to the Excel workbook
  .Visible = True
End With
' Release object memory
Set xlWkBk = Nothing: Set xlApp = Nothing
End Sub
