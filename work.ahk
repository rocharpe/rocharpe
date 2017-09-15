;调用英文输入法，使用if语句是，后面记得加括号，如if （i>1）
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

;获取存储变量长度的函数，当初这个困扰了很久
InputVar := zhi
ro:= StrLen(InputVar) 

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
 
;同样的，还有循环找颜色直到找到为止的命令，是否需要第一个参数，不是很清楚
CoordMode,Pixel
Loop
{
 PixelGetColor,tuy,631,525
 PixelGetColor,tue,535,532
 PixelGetColor,tus,890,536 
 if (tuy = "0xFFFFFF" and tue = "0xFFFFFF" and tus = "0xFFFFFF" )
 break
 } 
 
 ;一个强大的找text函数，feiyue老师提供
 S1:=获取屏幕坐标处的文本(985,703)
   获取屏幕坐标处的文本(x, y) {
  BlockInput, MouseMove
  CoordMode, Mouse
  MouseGetPos, 初始X, 初始Y
  ;-- 瞬间移动
  MouseMove, x, y, 0
  MouseGetPos,,,, cid, 2
  ControlGetText, s,, ahk_id %cid%
  ;-- 瞬间移动
  MouseMove, 初始x, 初始y, 0
  BlockInput, MouseMoveOff
  return, s
}
