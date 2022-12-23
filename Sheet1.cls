Sub Worksheet_Change(ByVal Target As Range)
Dim wb1 As Workbook, wb2 As Workbook
Dim ws1 As Worksheet, ws2 As Worksheet
Dim iRow As Integer
Dim icol As Integer
Dim lastRow As Long
Dim today As String
today = Format(Date, "mm-dd")
'当前选择的行
iRow = Selection.Row
icol = Selection.Column
Set wb1 = ActiveWorkbook
Set ws1 = wb1.ActiveSheet

'7为点击到当前行的第7列时执行插入到另一张表的方法，可根据需求更改
If icol = 7 Then
'替换成目标文件的位置
Set wb2 = Workbooks.Open("C:\Users\Administrator\Desktop\新建 XLS 工作表 (2).xls")
Set ws2 = wb2.ActiveSheet
With ActiveSheet
'求另一张表有内容的最后一行，前提需要第1列需要有数据，否则获取不到
lastRow = .Cells(ws2.Rows.Count, "A").End(xlUp).Row
End With
MsgBox "有内容的最后一行是：" & today
'默认在目标表的第1列插入当前日期
ws2.Cells(lastRow + 1, 1).Value = today
'2为从当前行的第2列开始插入一直到第6列，可根据需求更改
For i = 2 To 6
ws2.Cells(lastRow + 1, i).Value = ws1.Cells(iRow, i).Value

Next i
wb2.Save
End If
End Sub
