F4::
InputBox ,time,版本,***F4启动***ESC暂停***`r`n***F5刷新***F11退出***`r`n`r`n***请关闭输入法***`r`n输入单量,,200,200
i:=0
loop %time%
{ 
#SingleInstance,force
CoordMode,Pixel
CoordMode,mouse
i:=i+1
;设置
excel:= ComObjActive("Excel.Application")
FileName := "" ;无需指定文件名，;~ filepath:=A_ScriptDir . "\Desktop\cdms.xlsm"
Sh:=excel.Worksheets["Sheet2"]
;添加表格里的参数方便引用 
zhi:=num:=hsc:=pinm:=gongs:=fee:=bctype:=cft:=S1:=S2:=S3:=S4:=""
zhi:=sh.Range("B3").value 
num:= sh.Range("B3").text  
hsc:= sh.Range("E1").text  
pinm:= sh.Range("E2").text 
gongs:= sh.Range("E3").text 
fee:= sh.Range("E7").text
bctype:=sh.Range("E5").text 
cft:=sh.Range("E6").text 
#Include C:\test_game\find.ahk
;gs:="|<公司>42.S000000W000800U000800UAqwQqvUGGG8MGUGGG8EIWGGG8EAQADv6s80000008000000kU"
;quer:="|<确认>23.0UEFtsEUYE111024Ti4Cd48py8ceYFFLseXeFcZIWF1311U"
;yund:="|<运单>48.HyAADyzyM06M86A280TyDy8a7zNa86MaskTy00SY9gNazznz94TyA0m3/S1UDyHz/nzz06G2Q01U06S3rz1U3wSSU"
;tiaocu:="|<点跳出>51.SE8000002+0000004kE000000X2tDbz7DfCN9YnNBgUO98oFDcY3F96W914UO9AYFA8YSF9wW8x4s0080000000100000000800000U"
;mbz:="|<木包装>63.xzzrryzzzzzjzqyzr0M31s0zI0SvvpvjTryyzrTSj1ryzbrk830Ph0ruyzrTOf1Pqwq0yvvJSzSryjzp0OfLvqzyzyTzIMD0rU03a09urvuzuzmzTTSzTjyPjrPs31vzj/fyv3TSzTxbKTrPvvrw0DtwOpT0E7zzzTzhk/vzU"  ;if 查找文字(1028,828,50,50,mbz,"*194",X,Y,OCR,0,0)
;zhuihg:="|<最惠国>35.Ds10zyEFxx04zU42zd11r48TzWG8Ed07QLRSw9+VGYdxx2JtE43zuN2YY0TZ52jzVFFxE1U"  ;if 查找文字(467,867,50,50,zhuihg,"**50",X,Y,OCR,0,0)
;hs:="|<hs>15.T4TAqkZY4sUr44ckZaAaT4Q" ;if 查找文字(704,570,50,50,hs,"**50",X,Y,OCR,0,0)
;suijy:="|<税金>24.S0wSn1ann1ann1ann1ann1ann1ann1anSMwSU" ;if 查找文字(993,769,15,15,suijy,"**50",X,Y,OCR,0,0)
;zengzs:="|<增值税>23.Q0QR415682+AE4IMU8cl0FFW0WWsUsu" ;if 查找文字(992,747,15,15,zengzs,"**50",X,Y,OCR,0,0)
;done:="|<>59.s000s0r000U000E0Y0010000U1c0021VlV03FbVY4YYW05YeYc98740/7JDFGEG80GGeFyMQTw1oTIS";if 查找文字(628,366,100,100,done,"**50",X,Y,OCR,0,0)
;~判断单号是否为空
InputVar := zhi
ro:= StrLen(InputVar) 
if (ro < 9)
    {
    goto stop
    }
else 
    {
    ;~ 输入单号
    Click , 800,10 ;软件位置
    Sleep,200
    SendMessage,0x50,0,67699721,,A
    Sleep,50
    Clipboard:=num
    Sleep,100
    Click , 395,58 ;单号的位置
    Sleep,50
    Click
    Sleep,100
    send ,^v
    Sleep,50
    send ,{enter}
    clipboard:=
    Sleep,200
    ;~ 点击确定
    quer:="|<确认>23.0UEFtsEUYE111024Ti4Cd48py8ceYFFLseXeFcZIWF1311U"
    loop
        {
        if 查找文字(859,674,50,50,quer,"*113",X,Y,OCR,0,0)
        Break
        }
    CoordMode, Mouse
    MouseMove, X, Y
    sleep,100
    MouseClick, left
    sleep,60
    MouseClick, left
    sleep,100
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
    sleep,100
    MouseClick, left
    sleep,500         
    ;~ 判断加载情况
    done:="|<>59.s000s0r000U000E0Y0010000U1c0021VlV03FbVY4YYW05YeYc98740/7JDFGEG80GGeFyMQTw1oTIS"
    loop
        {
        if 查找文字(630,365,100,100,done,"**50",X,Y,OCR,0,0)
        Break
        }
	sleep,300
    ;~如果遇到木包装
    mbz:="|<木包装>63.xzzrryzzzzzjzqyzr0M31s0zI0SvvpvjTryyzrTSj1ryzbrk830Ph0ruyzrTOf1Pqwq0yvvJSzSryjzp0OfLvqzyzyTzIMD0rU03a09urvuzuzmzTTSzTjyPjrPs31vzj/fyv3TSzTxbKTrPvvrw0DtwOpT0E7zzzTzhk/vzU"  
    if 查找文字(1028,828,50,50,mbz,"*194",X,Y,OCR,0,0) 
        {
        CoordMode, Mouse
        MouseMove, x+226, y-68
        sleep,50
        MouseClick, left
        sleep,200
        MouseMove, 360,83
        sleep,200
        MouseClick, left
        sleep,100
            
        Click,1700,10 ;点击表格
        Sleep,100
        clipboard:=
        Sleep,50
        Clipboard:=num
        Sleep,100
        sh.Cells(i, 8).Select
        Sleep,100
        send ,^v
        Sleep,100
        sh.Cells(i, 9).Select
        Sleep,100
        send , 木包装！
        Sleep,50
        send ,{enter}
        Clipboard:=
        Sleep,100
        goto outer
        }
    ;~如果看到最惠国
    zhuihg:="|<最惠国>35.Ds10zyEFxx04zU42zd11r48TzWG8Ed07QLRSw9+VGYdxx2JtE43zuN2YY0TZ52jzVFFxE1U"
    if 查找文字(467,867,50,50,zhuihg,"**50",X,Y,OCR,0,0)
        {
        ;~ hs商品编码
        sleep,50
        Clipboard:=hsc 
        Sleep,100
        Click , 417,697 ;hs的位置417,697 
        Sleep,100
        send ,^v
        Sleep,100
        clipboard:=
        Sleep,50
        send,{enter}
        Sleep,150
        }
    else
    ;~ 如果没有看到最惠国 
        {
        Click,1700,10 ;点击表格
        Sleep,100
        clipboard:=
        Sleep,50
        Clipboard:=num
        Sleep,100
        sh.Cells(i, 8).Select
        Sleep,100
        send ,^v
        Sleep,100
        sh.Cells(i, 9).Select
        Sleep,100
        send , 非最惠国！或者其他不可抗因素
        Sleep,50
        send ,{enter}
        Clipboard:=
        Sleep,100   
        goto outer
        }
    ;~如果税金跳出错误框
    hs:="|<hs>15.T4TAqkZY4sUr44ckZaAaT4Q" 
    if 查找文字(704,570,50,50,hs,"**50",X,Y,OCR,0,0)
        {
        CoordMode, Mouse
        MouseMove, x, y
        sleep,50
        MouseClick, left
        sleep,100
                    
        Click,1700,10 ;点击表格
        Sleep,100
        clipboard:=
        Sleep,50
        Clipboard:=num
        Sleep,100
        sh.Cells(i, 8).Select
        Sleep,100
        send ,^v
        Sleep,100
        sh.Cells(i, 9).Select
        Sleep,100
        send , HS错误！
        Sleep,50
        send ,{enter}
        Clipboard:=
        Sleep,100
        goto outer
        }
    ;~ 如果税金同时为零
    zengzs:="|<增值税>23.Q0QR415682+AE4IMU8cl0FFW0WWsUsu" 
    suijy:="|<税金>24.S0wSn1ann1ann1ann1ann1ann1ann1anSMwSU" 
    if 查找文字(992,747,15,15,zengzs,"**50",X,Y,OCR,0,0) and if 查找文字(993,769,15,15,suijy,"**50",X,Y,OCR,0,0)
        {
        Click,1700,10 ;点击表格
        Sleep,100
        clipboard:=
        Sleep,50
        Clipboard:=num
        Sleep,100
        sh.Cells(i, 8).Select
        Sleep,100
        send ,^v
        Sleep,100
        sh.Cells(i, 9).Select
        Sleep,100
        send , 税金为零！未做
        Sleep,50
        send ,{enter}
        Clipboard:=
        Sleep,100
        goto outer
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
    gs:="|<公司>42.S000000W000800U000800UAqwQqvUGGG8MGUGGG8EIWGGG8EAQADv6s80000008000000kU"
    if 查找文字(791,317,50,50,gs,"*121",X,Y,OCR,0,0)
        {
        CoordMode, Mouse
        MouseMove, X+60, Y
        sleep,50
        MouseClick, left	
        sleep,50 
        Send,{Enter}
        sleep,50
        Clipboard:=gongs 
        Sleep,100
        send ,^v
        Sleep,100
        clipboard:=
        Sleep,100
        }
    ;根据表格中成交方式的的颜色和货值的颜色，进行判断
    ;成交方式和金额颜色判断，这里修改位置
    ;~ PixelGetColor, color1, 1786,405  ;FOB时，显示黄色 1855,474 
    ;~ PixelGetColor, color2, 1850,480  ;货值低于400rmb时，显示蓝色 1850,419
    ;~ 更新判断方式，更精准，不需要限制表格的格式及精确位置
    S4:=获取屏幕坐标处的文本(803,782)
    Sleep,50
    if (cft="FOB" and S4<400)  ;如果是FOB且金额小于400rmb
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
    if (cft="FOB" and S4>=400) ;如果是FOB且金额大于400rmb
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
    if (cft="CIF") ;如果是CIF
        {
        /*    
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
        goto outer
        */
        Sleep,100
        }
        

;这里添加识别税金功能，添加字库
;------------------------------------------------------------------------------------------------------------------------------------------------------
/*
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
Sleep,100
*/
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
    Click , 800,10 ;软件位置
    Sleep,200
    tiaocu:="|<>51.SE8000002+0000004kE000000X2tDbz7DfCN9YnNBgUO98oFDcY3F96W914UO9AYFA8YSF9wW8x4s0080000000100000000800000U"
    if 查找文字(326,77,50,50,tiaocu,"**50",X,Y,OCR,0,0)
        {
        CoordMode, Mouse
        MouseMove, X, Y
        sleep,100
        MouseClick, left
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
			
}
    return

    ;暂停
    ESc::
    Pause
    return
    ;加载
    F5::
    Reload
    return
    ;退出
    stop:
    F11::
    Send {Volume_Up}
    Sleep,50
    SoundSet, 70
    loop , 3
        {
        SoundPlay, C:\test_game\autokey\1214.wav
        Sleep,1800
        }
    ExitApp
    return
