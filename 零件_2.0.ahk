;==========================================================
;获取位置参数
TrayTip,鼠标移动到 <中文品名> 的位置按【F1】,. ,10,17
return
F1::
CoordMode, Mouse
MouseGetPos, cx, cy
odesc:=gethwnd(%cx%, %cy%) ;运单号
ControlGetText, desc, %odesc%, ahk_exe CDMSImport.exe
TrayTip,鼠标移动到 <单号> 的位置按【F2】,Desc(CN):%desc% ,10,17
return
F2::
CoordMode, Mouse
MouseGetPos, dx, dy
ohwab:=gethwnd(%dx%, %dy%) ;运单号
ControlGetText, hawb, %ohwab%, ahk_exe CDMSImport.exe
TrayTip,按【F4】 启动,Hawb:%hawb%,10,17
return
;==========================================================

F4::
#SingleInstance,force
CoordMode,Pixel
CoordMode,mouse
ComObjError(false) ;关闭对象错误提示

;====================================================
;定义要搜索的变量
AA:
BlockInput, MouseMove
DA:=DB:=PM:=""
WinActivate ahk_exe CDMSImport.exe
ControlGetText, DA, %ohwab%, ahk_exe CDMSImport.exe
ControlGetText, PM, %odesc%, ahk_exe CDMSImport.exe
;====================================================

if (DA<>"" and PM="")
{
ControlSetText, %odesc%, 零件, ahk_exe CDMSImport.exe
Sleep,50
send {PGDN}
Sleep,100
i:=0

BB:
DB:=""
i:=i+1
ControlGetText, DB, %ohwab%, ahk_exe CDMSImport.exe
if (DA=DB and i<150)
{
Sleep,100
goto BB
}
else
{
goto AA
}
}
else
	{
	Send {Volume_Up}
    Sleep,50
    SoundSet, 60
	loop , 3
        {
        SoundBeep, 2100, 500
        Sleep,200
        }
	} 

Sleep,150
goto AA 
return


ESc::
BlockInput, MouseMoveOff
    Send {Volume_Up}
    Sleep,50
    SoundSet, 0

ExitApp
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


