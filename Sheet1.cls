VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub Worksheet_Change(ByVal Target As Range)
Dim wb1 As Workbook, wb2 As Workbook
Dim ws1 As Worksheet, ws2 As Worksheet
Dim iRow As Integer
Dim icol As Integer
Dim lastRow As Long
Dim today As String
today = Format(Date, "mm-dd")
'��ǰѡ�����
iRow = Selection.Row
icol = Selection.Column
Set wb1 = ActiveWorkbook
Set ws1 = wb1.ActiveSheet

If icol = 7 Then
Set wb2 = Workbooks.Open("C:\Users\Administrator\Desktop\�½� XLS ������ (2).xls")
Set ws2 = wb2.ActiveSheet
With ActiveSheet
'����һ�ű������ݵ����һ��
lastRow = .Cells(ws2.Rows.Count, "A").End(xlUp).Row
End With
MsgBox "�����ݵ����һ���ǣ�" & today
ws2.Cells(lastRow + 1, 1).Value = today
For i = 2 To 6
ws2.Cells(lastRow + 1, i).Value = ws1.Cells(iRow, i).Value

Next i
wb2.Save
End If
End Sub