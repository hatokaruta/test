Option Explicit

Sub Main()
  'On Error Resume Next
 
  ' Excelアプリケーションのインスタンス生成
  Dim objXls : Set objXls = CreateObject("Excel.Application")
  'If Not objXls Then Exit Sub

 dim fs
 Set fs = CreateObject("WScript.Shell")

  ' Excelの表示
  'objXls.Visible = True
  'objXls.ScreenUpdating = False

	dim objWorkbook,objSheet
  Set objWorkbook = objXls.Workbooks.Open(fs.CurrentDirectory & "\aaaa.xlsx") 
  Set objSheet = objXls.Sheets("Sheet1")
  'objSheet.activate

  objXls.ActiveSheet.Range("$A:$H").AutoFilter 2, "asd"

  objWorkbook.Save
  objWorkbook.Close
 
  objXls.Quit

  set objWorkbook = nothing
  set fs = nothing
  Set objXls = Nothing
End Sub

Main

'    ActiveSheet.Range("$A$1:$H$29").AutoFilter(2, "adf")

