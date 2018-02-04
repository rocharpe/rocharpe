/*
感谢feiyue老师提供帮助
by rocharpe 2018.2.4
*/
F4::
InputBox ,time,CDMS 2.8_2018-2,***F4 启动***`r`n***ESC暂停***`r`n***F5 退出***`r`n`r`n输入单量,,200,200
i:=0
loop %time%
{ 
#SingleInstance,force
CoordMode,Pixel
CoordMode,mouse
i:=i+1
;设置
excel:= ComObjActive("Excel.Application")
FileName := "" ;无需指定文件名
Sh:=excel.Worksheets["Sheet2"]

;添加表格里的参数方便引用 
zhi:=num:=hsc:=pinm:=gongs:=fee:=bctype:=cft:=S1:=S2:=S3:=S4:=S5:=S6:=""
zhi:=sh.Range("B3").value 
num:= sh.Range("B3").text  
hsc:= sh.Range("E1").text  
pinm:= sh.Range("E2").text 
gongs:= sh.Range("E3").text 
fee:= sh.Range("E7").text
bctype:=sh.Range("E5").text 
cft:=sh.Range("E6").text 
#Include C:\test_game\find.ahk

;~判断单号是否为空
if (zhi<10000000)
    {
    gosub stop
    }
else 
    {
    BlockInput, MouseMove
    ;~ 输入单号
    Click , 800,10 ;CDMS位置
    Sleep,200
    SendMessage,0x50,0,67699721,,A
    Sleep,50
    Clipboard:=num
    Sleep,100
    Click,395,58 ;运单号的位置
    gosub double
    send ,{enter}
    Sleep,200
    ;~ 点击确定
    quer:="|<确认>23.0UEFtsEUYE111024Ti4Cd48py8ceYFFLseXeFcZIWF1311U"
    loop
        {
        if 查找文字(859,674,150,150,quer,"*113",X,Y,OCR,0,0)
        Break
        }
    CoordMode, Mouse
    MouseMove, X, Y
    sleep,150
    Click
    sleep,50
    Click
    sleep,50
    MouseMove, 489,83
    sleep,300
    } 
    Loop
        {
        PixelGetColor,tuy,631,525
        PixelGetColor,tue,535,532
        PixelGetColor,tus,890,536 
        if (tuy = "0xFFFFFF" and tue = "0xFFFFFF" and tus = "0xFFFFFF" )
        break
        } 
    sleep,300
    Click
    sleep,500         
    ;~ 判断加载情况
    hsc:="|<>23.TzzxU00C000AzzyN004mn09YY0H980aHnlAYY2N964mG29ZawH800aTzzA000Q001jzzyU"
    loop
        {
        if 查找文字(445,722,150,150,hsc,"**50",X,Y,OCR,0,0)
        Break
        }
	sleep,300
    ;~如果遇到木包装
    mbz:="|<木包装>63.xzzrryzzzzzjzqyzr0M31s0zI0SvvpvjTryyzrTSj1ryzbrk830Ph0ruyzrTOf1Pqwq0yvvJSzSryjzp0OfLvqzyzyTzIMD0rU03a09urvuzuzmzTTSzTjyPjrPs31vzj/fyv3TSzTxbKTrPvvrw0DtwOpT0E7zzzTzhk/vzU"  
    if 查找文字(1028,828,150,150,mbz,"*194",X,Y,OCR,0,0) 
        {
        CoordMode, Mouse
        MouseMove, x+226, y-68
        sleep,50
        Click
        sleep,200
        MouseMove, 360,83
        sleep,200
        Click
        gosub errorp
        send , 木包装！
        Sleep,50
        send ,{enter}
        gosub outer
        }
    ;~如果看到最惠国
    zhuihg:="|<最惠国>35.Ds10zyEFxx04zU42zd11r48TzWG8Ed07QLRSw9+VGYdxx2JtE43zuN2YY0TZ52jzVFFxE1U"
    if 查找文字(467,867,50,50,zhuihg,"**50",X,Y,OCR,0,0)
        {
        ;~ hs商品编码
        sleep,50
        Clipboard:=hsc 
        Sleep,100
        Click,417,697 ;hs的位置 
        gosub double
        send,{enter}
        Sleep,100
        }
    else
    ;~ 如果没有看到最惠国 
        {
        gosub errorp
        send , 非最惠国！或者其他不可抗因素
        Sleep,50
        send ,{enter} 
        gosub outer
        }
    ;~如果税金跳出错误框
    hs:="|<hs>15.T4TAqkZY4sUr44ckZaAaT4Q" 
    if 查找文字(704,570,50,50,hs,"**50",X,Y,OCR,0,0)
        {
        CoordMode, Mouse
        MouseMove, x, y
        sleep,50
        Click
        sleep,100
        gosub errorp
        send , HS错误！
        Sleep,50
        send ,{enter}
        gosub outer
        }

S5:=获取屏幕坐标处的文本(986,769)
Sleep,50
    ;~ 如果税金同时为零
    if (S5<0.01)
        {
        gosub errorp
        send , 税金为零！未做
        Sleep,100
        send ,{enter}
        gosub outer
        }
    ;~否则输入品名
    i:=i-1
    Clipboard:=pinm  
    Sleep,100
    send ,^v
    Sleep,100
    clipboard:=
    Sleep,100
    send,{enter}
    sleep,100
    Send,无规格
    sleep,100
    
    ;~跳转到公司的位置，并输入公司
	;~ MouseClickDrag,L,996,365,664,365
    click,830,365
    sleep,100
    Clipboard:=gongs 
    gosub double

    S4:=获取屏幕坐标处的文本(803,782)
    Sleep,50
    if (InStr(cft,"fob") and S4<400)  ;如果是FOB且金额小于400rmb
        {
        Click,476,218  ;c lei
        sleep,100
        send,b
        Sleep,50
        Click,797,282
        sleep,50
        Click
        sleep,100
        send,cif
        sleep,100
        send,fob
        sleep,100
        Click,408,279
        sleep,50
        Click
        sleep,100
        send,^{x}
        sleep,50
        Click,476,218  ;c lei
        sleep,100
        send,c
        sleep,50
        send,{enter}
        sleep,100
        Click,408,279
        sleep,50
        Click
        sleep,100
        send,^{v}
        Sleep,100
        clipboard:=
        Sleep,100                        
        }
    if (InStr(cft,"fob") and S4>=400) ;如果是FOB且金额大于400rmb
        {
        Click,797,282
        sleep,50
        Click
        sleep,100
        send,cif
        sleep,100
        send,fob
        sleep,100
        Click,415,279
        sleep,50
        Click
        sleep,100
        Clipboard:=fee 
        Sleep,100
        send ,^v
        Sleep,100
        clipboard:=
        Sleep,100
        }
    if (InStr(cft,"CIF")) ;如果是CIF
        {
            ;看看这里能不能加上cif判断，并做修改
            ;=================================
        Sleep,100
        }

S1:=获取屏幕坐标处的文本(985,703)
S2:=获取屏幕坐标处的文本(985,725)
S3:=获取屏幕坐标处的文本(985,746)
Sleep,100
if (S1<50 and S2<50 and S3<50)
    {
    Sleep,50
    Click,476,218  ;c lei
    sleep,100
    send,b
    sleep,50
    send,{enter}
    sleep,50	
    }

;------------------------------------------------------------------------------------------------------------------------------------------------------

    change:="|<>51.Ak0M0D061bnv01g0kNa3M7nrq2rUPwlUBDVb2n6A1aOFXvMggAnGMMP6ZVaOCn2kpMAnGQPA6alwOyOnkhutXMCIn6NEMPT3ABa66CU" 
    send, {F9}
    sleep,100
    loop
        {
        if 查找文字(535,106,501,50,change,"**50",X,Y,OCR,0,0)
        Break
        }
    send, {F8}
    sleep,200
                
    ; 跳出后，点击单号
    outer:
    Sleep,100
    Click , 800,10 ;CDMS位置
    Sleep,200
    tiaocu:="|<>51.SE8000002+0000004kE000000X2tDbz7DfCN9YnNBgUO98oFDcY3F96W914UO9AYFA8YSF9wW8x4s0080000000100000000800000U"
    if 查找文字(326,77,50,50,tiaocu,"**50",X,Y,OCR,0,0)
        {
        CoordMode, Mouse
        MouseMove, X, Y
        sleep,100
        Click
        sleep,100
        }
    Click,1700,10 ;点击表格
    sleep,200

go:="|<>15.D3W8WV8A11U8A11XcA91F4FkQU"

if 查找文字(1456,209,100,100,go,"000000-000000",X,Y,OCR,0,0)
{
  CoordMode, Mouse
  MouseMove, X, Y
  Click ;下一票
  sleep,301
}
			
}
return

;===================================================================
  获取屏幕坐标处的文本(x, y) {
  CoordMode, Mouse
  MouseGetPos, 初始X, 初始Y
  ;-- 瞬间移动
  MouseMove, x, y, 0
  MouseGetPos,,,, cid, 2
  ControlGetText, s,, ahk_id %cid%
  ;-- 瞬间移动
  MouseMove, 初始x, 初始y, 0
  return, s
}
;===================================================================
double:
    Sleep,100
    send ,^a
    Sleep,100
    send ,^v
    Sleep,100
    clipboard:=
    Sleep,100
    return
;===================================================================
Errorp:
    sleep,100
    Click,1700,10 ;点击表格
    Sleep,100
    SendMessage,0x50,0,67699721,,A
    Sleep,50
    clipboard:=
    Sleep,50
    Clipboard:=num
    Sleep,100
    sh.Cells(i, 8).Select
    Sleep,100
    send ,^v
    Sleep,100
    sh.Cells(i, 9).Select
    Sleep,50
    Clipboard:=
    Sleep,100
    return
;===================================================================

/*
    ;加载
    F5::
    Reload
    return
*/ 
    ;停止
                        
    ESc::
    BlockInput, MouseMoveOff
    Pause
    return

    ;加载
    F5::
    BlockInput, MouseMoveOff
    ExitApp
    return

    ;退出
    stop:
    F11::
    BlockInput, MouseMoveOff
    Send {Volume_Up}
    Sleep,50
    SoundSet, 65
    loop , 3
        {
        SoundPlay, C:\test_game\autokey\1214.wav
        Sleep,1800
        }
    ExitApp
    return
