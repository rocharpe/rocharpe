TrayTip,鼠标移动到单号的位置-->按F2,. ,10,17
return
;=====================
;获取最终的位置
F1::
CoordMode, Mouse
  MouseGetPos, cx, cY

return
F2::
CoordMode, Mouse
  MouseGetPos, dx, dY
  ohwab:=gethwnd(%dx%, %dx%) ;运单号
ControlGetText, thwab, %ohwab%, ahk_exe CDMSImport.exe

return

`::
F11::
ins::
#SingleInstance,force
CoordMode,Pixel
CoordMode,mouse
ComObjError(false) ;关闭对象错误提示



;================================================================================
;定义要搜索的变量,此处定义单号为hawb  WindowsForms10.EDIT.app.0.215472d390  WindowsForms10.EDIT.app.0.202c66624
;ohwab:=gethwnd(967, 196) ;运单号

WinActivate ahk_exe CDMSImport.exe
ControlGetText, hwab, %ohwab%, ahk_exe CDMSImport.exe
;================================================================================

WinActivate ahk_class IEFrame 
iWeb_Activate("NPTS - A DHL Product") 
iWeb_Activate("WebFSQ - ShipmentDetails") 

ie:=IEGetFromUrl("npts2.apis.dhl.com")  ;获取包含指定网页的一个选项卡
ie.document.GetElementsByTagName("a").item(1).Click()

While ie.readyState != 4 || ie.document.readyState != "complete" || wb.busy
sleep,200
ie.document.GetElementsByTagName("input").item(4).value:=hwab  ;通过ie.doucument对当前网页进行操作,kw为搜索框
sleep,50
ie.document.getElementById("searchButton").Click()
;send ,^+S
While ie.readyState != 4 || ie.document.readyState != "complete" || wb.busy
  sleep,200
  send,{enter}
  sleep,300
  send,{enter}
  sleep,300
  send,{enter}

WinActivate ahk_exe CDMSImport.exe
  MouseMove, %cX%, %cY%, 0
  sleep 50
  click
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

    ESc::
    ExitApp
    return
