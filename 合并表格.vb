把要合并的文件放同一个文件夹下面。
再新建的excel表，并右键单击sheet1，找到“查看代码”，单击进去。
代码复制进去，然后找到工具栏上面的“运行”下的“运行子过程/用户窗体”

Sub 合并当前目录下所有工作簿的全部工作表()
Dim MyPath, MyName, AWbName
Dim Wb As Workbook, WbN As String
Dim G As Long
Dim Num As Long
Dim BOX As String
Application.ScreenUpdating = False
MyPath = ActiveWorkbook.Path
MyName = Dir(MyPath & "\" & "*.xls")
AWbName = ActiveWorkbook.Name
Num = 0
Do While MyName <> ""
If MyName <> AWbName Then
Set Wb = Workbooks.Open(MyPath & "\" & MyName)
Num = Num + 1
With Workbooks(1).ActiveSheet
.Cells(.Range("B65536").End(xlUp).Row + 2, 1) = Left(MyName, Len(MyName) - 4)
For G = 1 To Sheets.Count
Wb.Sheets(G).UsedRange.Copy .Cells(.Range("B65536").End(xlUp).Row + 1, 1)
Next
WbN = WbN & Chr(13) & Wb.Name
Wb.Close False
End With
End If
MyName = Dir
Loop
Range("B1").Select
Application.ScreenUpdating = True
MsgBox "共合并了" & Num & "个工作薄下的全部工作表。如下：" & Chr(13) & WbN, vbInformation, "提示"
End Sub
