F4:: ;快速启动
#SingleInstance,force
CoordMode,Pixel
CoordMode,mouse
#Include C:\test_game\find.ahk
BlockInput, MouseMove

;--------------------------------------------
;获取get图标
getto:
get:="|<>19.D03AE1g8xo0mODNh1jqkq3APBXkww"
gett:="|<>19.DU3As1gAxy0nPDNhVjqkq3AvBXsww"
got:="|<>18.T02lU2kXZU6mU4GXrmkY2lamD3XU"

if 查找文字(2496,279,150,150,gett,"**50",X,Y,OCR,0,0) or 查找文字(2496,279,150,150,get,"**50",X,Y,OCR,0,0) or 查找文字(2481,278,150,150,got,"**50",X,Y,OCR,0,0)
{
  CoordMode, Mouse
  MouseMove, X, Y
  sleep,100
  click
  sleep,100
}
else
{
  WinActivate, ahk_class IEFrame
  send ^0
  sleep,200
  goto getto
}

;--------------------------------------------
;等待空白框出现
wait:
deng:="|<>13.zzk0M0A060301U0k0M0A060301zzk"

if not 查找文字(1316,307,150,150,deng,"**50",X,Y,OCR,0,0)
{
  sleep ,200
  goto wait
}

;-------------------------
;可视网络情况适当增加等待时间
sleep ,2500

;--------------------------------------------
;查找框是否有打钩,新版发现此处不需要判断,速度更快
;~ AA:
;~ ggg:="|<>13.zzk0M0A1a0n0lgsrMNsAQ64301zzk"
;~ goug:="|<>13.zzk0M0A0q0P0NaQngMwAC62301zzk"

;~ if 查找文字(1316,307,150,150,gou,"**60",X,Y,OCR,0,0)
;~ {
  ;~ sleep ,200
  ;~ goto AA
;~ }

;----------------------------------------------------------------------------------------------------------------
;查找是否有两个竖的空白框,抓图时需特别设置
BB:
kuang:="|<>15.Dzt018090180901809018090180901Dzs00000000000000000000000000000000zzY04U0Y04U0Y04U0Y04U0Y04U0Y04zzY"
kuanger:="|<>14.zzc0+02U0c0+02U0c0+02U0c0+02zzU00000000000000000000000000000000007zx01E0I0501E0I0501E0I0501E0Lzy"

if 查找文字(1315,320,150,150,kuang,"**90",X,Y,OCR,0,0) or 查找文字(1315,320,150,150,kuanger,"**50",X,Y,OCR,0,0)
{
  CoordMode, Mouse
  MouseMove, X, Y-13
  sleep,100
  click
  sleep,100
  send {F6}
  sleep,100

  ;-------------------------------------------------------------------------
  ;持续查找select,直到找到为止
  DD:
  xuan:="|<>34.D0600Na0M01YAtbXeQ6qqPMylPD1URxjo6kw6kkNaPPBhXktbXbU"

  if not 查找文字(1313,166,150,150,xuan,"**50",X,Y,OCR,0,0)
  {
    sleep,50
    goto DD
  }

  if 查找文字(1313,166,150,150,xuan,"**50",X,Y,OCR,0,0)
  {
    CoordMode, Mouse
    MouseMove, X, Y
    sleep,100
    click
    sleep,200
  }

  ;-----------------------------------------------------------------------------
  cim:="|<>21.Dasv4r7kqci0phk6xi0qjkqrP6qvDanQ"

  if 查找文字(1380,241,150,150,cim,"**50",X,Y,OCR,0,0)
  {
    CoordMode, Mouse
    MouseMove, X, Y
    sleep,100
    click
    sleep,200
  }

  ok:="|<>17.D3Cnav39a6KAAyMNgknAnaQwAQ"

  if 查找文字(1331,208,150,150,ok,"**50",X,Y,OCR,0,0)
  {
    CoordMode, Mouse
    MouseMove, X, Y
    sleep,200
    click,2
    sleep,200
  }
  
  sleep , 100
  goto AA
}

;------------------------------------------------------------------------------------
pre:="|<>43.z003000Mk00000ALiHnqraBBhfPPOxbqZZhlUn1qmqCkNgvPPPMASNb7b00000007zzzzzzy"

if 查找文字(2333,615,150,150,pre,"**50",X,Y,OCR,0,0)
{
  CoordMode, Mouse
  MouseMove, X, Y
  sleep , 100
  click
  loop,20
  {

kuang:="|<>15.Dzt018090180901809018090180901Dzs00000000000000000000000000000000zzY04U0Y04U0Y04U0Y04U0Y04U0Y04zzY"
kuanger:="|<>14.zzc0+02U0c0+02U0c0+02U0c0+02zzU00000000000000000000000000000000007zx01E0I0501E0I0501E0I0501E0Lzy"

if 查找文字(1315,320,150,150,kuang,"**90",X,Y,OCR,0,0) or 查找文字(1315,320,150,150,kuanger,"**50",X,Y,OCR,0,0)
	{
  	goto AA
	}
	else
	{
	sleep,150
	}
  }
goto AA
}

else
{
  sleep , 1500
  goto getto
}

return


ESc::
  BlockInput, MouseMoveOff
  ExitApp
return
