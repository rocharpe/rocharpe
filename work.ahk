;调用英文输入法
SendMessage,0x50,0,67699721,,A

;弹出提示框，并输入需要的
InputBox ,time,标题,***F4启动***ESC暂停***`r`n***F5刷新***F11退出***`r`n`r`n***请关闭输入法***`r`n输入,,200,200

;绝对值坐标需要的参数，好像每个参数需要换行才生效
#SingleInstance,force
CoordMode,mouse

;调用Excel需要的参数
excel:= ComObjActive("Excel.Application")
FileName := "" ;无需指定文件名，;~ filepath:=A_ScriptDir . "\Desktop\cdms.xlsm"
Sh:=excel.Worksheets["Sheet2"]
;添加表格里的参数方便引用 
zhi:=sh.Range("B3").value 
num:= sh.Range("B3").text  

;插入其他ahk需要用到的参数
#Include C:\test_game\find.ahk

;持续找图，直到找到的命令，在循环前加一点时间可能会更好，反正时需要等待
#Include C:\test_game\find.ahk
Sleep,300
quer:="|<确认>23.0UEFtsEUYE111024Ti4Cd48py8ceYFFLseXeFcZIWF1311U"
loop
  {
  if 查找文字(859,674,50,50,quer,"*113",X,Y,OCR,0,0)
  Break
  }
 
;
