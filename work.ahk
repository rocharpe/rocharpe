;调用英文输入法，使用if语句是，后面记得加括号，如if （i>1）
;[!A-Za-z0-9一-﨩] 一键替换有标点符号的文档
SendMessage,0x50,0,134481924,,A ;中文键盘
SendMessage,0x50,0,67699721,,A ;英文键盘
SendMessage,0x50,0,-534640636,,A ;搜狗输入法
;还有一个方法比较灵活，通过查询注册表的键值来调用输入法，比如搜狗的为E0220804
HKL:=DllCall("LoadKeyboardLayout", Str,"E02208048",UInt,1)
ControlGetFocus,ctl,A
SendMessage,0x50,0,HKL,%ctl%,A

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
;根据观察，发现光简单的循环不能满足当前的需求，因为需要的窗口会变动，所以在if前面增加了几个按键，应该可以解决当前的需求
#Include C:\test_game\find.ahk
Sleep,300
quer:="|<确认>23.0UEFtsEUYE111024Ti4Cd48py8ceYFFLseXeFcZIWF1311U"
loop
  {
  Click,705,90 ;会跳出的窗口
  Sleep ,1000
  Click,838,90 ;当前检测的窗口
  Sleep,1000
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

;-- 利用快捷键抓点并直接生成代码，feiyue老师提供
F2::
CoordMode, Mouse
MouseGetPos, x, y
s=%x%, %y%
Clipboard:=s
ToolTip, 抓点成功！可直接粘贴代码！
Sleep, 500
ToolTip
return

;选中需要的文字，并删除
MouseClickDrag, L, 255,120, 52,120
Sleep,50
send , {bs}

SendEvent {Click 6, 52, down}{click 45, 52, up} ;同样的效果，更具兼容性

;直接修改框架里的内容
#IfWinActive ahk_exe ;先指定需要操作控件的ClassNN，可以用spy查看
ControlSetText , edit1 , %var% ;同样的eiit1可以使用spy查看，后面可以用变量

;鼠标操作
Click, right ;右击鼠标
Click  ; 在鼠标光标的当前位置点击鼠标左键.
Click 100, 200  ; 在指定坐标处点击鼠标左键.
Click 100, 200, 0  ; 移动而不点击鼠标.
Click 100, 200 right  ; 点击鼠标右键.
Click 2  ; 执行双击.

;网页搜索
Click
Sleep 100
send ^A
Sleep 100
Clipboard =  https://www.baidu.com/s?wd=%Clipboard%
Sleep 100
send ^v
send {enter}
Sleep 100
Clipboard=
Sleep 100

;添加字库
zk.="|<1>3.KGGLU"
zk.="|<2>5.R68W8Vy"
zk.="|<3>5.R49UVWu"
zk.="|<4>5.4NGdD4C"
zk.="|<5>5.z27UVWu"
zk.="|<6>5.R+7clWu"
zk.="|<7>5.z8F248G"
zk.="|<8>5.R6/clWu"
zk.="|<9>5.R6ALVGu"
zk.="|<0>5.R6AMlWu"
zk.="|<.>5.000000W"

;-- 以小数点前一位作为坐标点
one:=two:=thr:=""
查找文字(985,703,26,8,zk,"6D6D6D-000000",X,Y,one, 0, 0)
查找文字(985,725,26,8,zk,"6D6D6D-000000",X,Y,two, 0, 0)
查找文字(985,746,26,8,zk,"6D6D6D-000000",X,Y,thr, 0, 0)


