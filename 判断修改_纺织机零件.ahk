;=========================================================
;获取位置参数 (需要修改坐标参数,打开鼠标锁定,增加i的次数)
TrayTip,【F1】获取数据 【F4】启动,. ,10,17
return

F1::  
CoordMode, Mouse
ohwab:=gethwnd(969, 200)  ;运单ohwab
odescen:=gethwnd(402, 750)  ;英文品名 odesc
odesccn:=gethwnd(373, 766)  ;中文品名 odesc
ControlGetText, csdh, %ohwab%, ahk_exe CDMSImport.exe
ControlGetText, csypm, %odescen%, ahk_exe CDMSImport.exe
ControlGetText, cszpm, %odesccn%, ahk_exe CDMSImport.exe
MsgBox 单号:%csdh%`r`n英文品名:%csypm%`r`n中文品名:%cszpm%
return

F2::
CoordMode, Mouse
BlockInput, MouseMoveOff
MouseGetPos, x, y
s=%x%, %y%
Clipboard:=s
ToolTip, 抓点成功！可直接粘贴代码！
Sleep, 1000
ToolTip
return
;=========================================================

F4::
#SingleInstance,force
CoordMode,Pixel
CoordMode,mouse
ComObjError(false) ;关闭对象错误提示

;====================================================
;定义要搜索的变量,DA为单号一,DB为单号二,PM为中文品名
AA:
BlockInput, MouseMove
DA:=ENPM:=CNPM:=""
WinActivate ahk_exe CDMSImport.exe
ControlGetText, DA, %ohwab%, ahk_exe CDMSImport.exe
ControlGetText, ENPM, %odescen%, ahk_exe CDMSImport.exe
ControlGetText, CNPM, %odesccn%, ahk_exe CDMSImport.exe
;====================================================

if (DA="")
{
Sleep,100
goto AA
}
Sleep,200

if  (InStr(ENPM,"spare") and InStr(CNPM,"纺织机零件") ) 
{
ControlSetText, %odesccn%,, ahk_exe CDMSImport.exe
Sleep,50
send {PGDN}
Sleep,500
i:=0
goto BB
}
else
{
send {PGDN}
Sleep,500
i:=0
MsgBox 条件不成立,是否下一票
goto BB
}

MsgBox 我是怎么出来的
goto AA 
return

BB:
DB:=""
i:=i+1
ControlGetText, DB, %ohwab%, ahk_exe CDMSImport.exe

if (DA<>DB and DB<>"" or i>50)
{
Sleep,500
goto AA
}
else
{
Sleep,100
goto BB
}
return

ESc::
BlockInput, MouseMoveOff
ExitApp
return

NumLock::
Pause
return

















































;=======================================================================================
;===================================函数部分=============================================
;切换ie标签
	iWeb_Activate(sTitle) 
	{ 

		DllCall("LoadLibrary", "str", "oleacc.dll") 
		DetectHiddenWindows, On 
		ControlGet, hTabBand, hWnd,, TabBandClass1, ahk_class IEFrame
		ControlGet, hTabUI  , hWnd,, DirectUIHWND1, ahk_id %hTabBand% 
		
		If   hTabUI && DllCall("oleacc\AccessibleObjectFromWindow", "Uint", hTabUI, "Uint",-4, "Uint", GUID(IID_IAccessible,"{618736E0-3C3D-11CF-810C-00AA00389B71}"), "UintP", pacc)=0 
		{ 
			Loop, %   pacc.accChildCount 
				If   paccChild:=pacc.accChild[A_Index] 
					If   paccChild.accRole[0] = 0x3C 
					{ 
						paccTab:=paccChild 
						Break 
					} 
		} 
		If   pacc:=paccTab 
		{ 
			Loop, %   pacc.accChildCount 
				If   paccChild:=pacc.accChild[A_Index] 
					If   paccChild.accName[0] = sTitle   
					{ 
						paccChild.accDoDefaultAction[0]
						Break 
					} 
		}  
		WinActivate,% sTitle
	} 
	
 GUID(ByRef GUID, sGUID) ; Converts a string to a binary GUID and returns its address.
{
    VarSetCapacity(GUID, 16, 0)
    return DllCall("ole32\CLSIDFromString", "wstr", sGUID, "ptr", &GUID) >= 0 ? &GUID : ""
}


;=======================================================================================
;获取包含指定url的IE选项卡对象,从而成功操作对应的doucment对象
IEGetFromUrl(url){
	for window in ComObjCreate("Shell.Application").Windows
	{
		if InStr(window.document.url,url) && InStr( window.FullName, "iexplore.exe" )
			return window
	}
}
 
;获取包含指定标题的IE选项卡对象
IEGetFromTabName(IETabName)
{
	For window in ComObjCreate( "Shell.Application" ).Windows
	{
		if ( window.LocationName = IETabName ) && InStr( window.FullName, "iexplore.exe" )
			return window
	}
}


;======================================
gethwnd(x,y)	
{
  BlockInput, MouseMove
  CoordMode, Mouse
  MouseGetPos, newX, newY
  MouseMove, x, y, 0
  MouseGetPos,x,y,id,oHWND
  MouseMove, newx, newy, 0
  BlockInput, MouseMoveOff
  return,oHWND
}


  gettext(x, y) {
  BlockInput, MouseMove
  CoordMode, Mouse
  MouseGetPos, newX, newY
  ;-- 瞬间移动
  MouseMove, x, y, 0
  MouseGetPos,,,, cid, 2
  ControlGetText, s,, ahk_id %cid%
  ;-- 瞬间移动
  MouseMove, newx, newy, 0
  BlockInput, MouseMoveOff
  return, s
}


