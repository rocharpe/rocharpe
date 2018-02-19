MsgBox 点击确定后`r`n按F2`r`n然后鼠标移动到单号的位置

F2::
	WinGetTitle, active_title, A
	MouseGetPos, mrX, mrY,, msCtrl
	ControlGetText, hawb, %msCtrl%, ahk_exe SciTE.exe
	MsgBox 请确认单号`r`n%hawb%
	return

#a::
#SingleInstance,force
CoordMode,Pixel
CoordMode,mouse
#Include C:\test_game\find.ahk
#Include C:\test_game\com.ahk
ComObjError(false) ;关闭对象错误提示

;================================================================================
;定义要搜索的变量,此处定义单号为hawb
;~ ControlGetText, hawb, WindowsForms10.EDIT.app.0.202c66624, ahk_exe CDMSImport.exe
;================================================================================

iWeb_Activate(ntps) ;激活初始网址
iWeb_Activate(webspme) ;激活查询结果网址
ie:=IEGetFromUrl("npts2.apis.dhl.com")  ;获取包含指定网页的一个选项卡
ie.document.GetElementsByTagName("a").item(1).Click()

;===========================================================================================================
;等待网页加载的同时,插入查询报警
cuo:="|<>15.T6TAqkZY4sUr44ckZaAaT4Q"
if 查找文字(704,570,150,150,cuo,"**50",X,Y,OCR,0,0)
{
  CoordMode, Mouse
  MouseMove, X, Y
  sleep,50
  click
  sleep,50
}

;=========================
  shawb:=gettext(967,196)
     Sb:=gettext(763,367)
     Sc:=gettext(383,388)
     Sd:=gettext(703,387)
;=========================

if  (InStr(sb,"FUTURE") or InStr(sb,"COVANCE")) 
{
MsgBox FUTURE ELECTRONICS重量看发票`r`n`r`nCOVANCE认运单,不用核重
}
sleep,50
  
if  (InStr(sb,"ge ") or InStr(sb,"jabil") or InStr(sb,"getin")) 
{
MsgBox 注意公司`r`n`r`nGETINGE件数有X核件核重
}
sleep,50

if (InStr(sb,"crg") or InStr(sb,"jetmile") or InStr(sb,"source") or InStr(sb,"sinophile"))
{
MsgBox 可能需要转C类
}
sleep,50

if (InStr(sc,"325") or InStr(sc,"warranty") or InStr(sc,"508"))
{
MsgBox GE325/508/78 JinYing Road
}
sleep,50

if (InStr(sd,"JinYing"))
{
MsgBox 78 JinYing Road
}
sleep,50

if  (InStr(sb,"dresser") or InStr(sb,"dms")) 
{
MsgBox 核重
}
;===========================================================================================================

iWeb_Activate(ntps) ;激活初始网址
iWeb_complete(pwb) ;等待
ie.document.GetElementsByTagName("input").item(4).value:=hawb  ;通过ie.doucument对当前网页进行操作,kw为搜索框
send ,^+S
iWeb_complete(pwb)
  sleep,300
  send,{enter}
  ;~ sleep,300
  ;~ send,{enter}
  ;~ sleep,500
  ;~ send,{enter}


hsc:="|<>23.TzzxU00C000AzzyN004mn09YY0H980aHnlAYY2N964mG29ZawH800aTzzA000Q001jzzyU"
if 查找文字(445,722,150,150,hsc,"**50",X,Y,OCR,0,0)
{
  CoordMode, Mouse
  MouseMove, X-30, Y
  click
  sleep,100
  send,^a
  sleep,100
  send,{del}
  sleep,100
  send,{enter}
  sleep,100

descn:="|<>107.zzzzzzzzzyDzzzzzzz000000000400000002000000000800000004000000000E00000008000040100U0000000LU000Hqt0100000000YU0018YV020000000190002EB24400000002GASC4UO40800000004YYUY90g80E000000099sl0G1ME0U0000000GG0G0YGEU100000000j3bXUbCW220000000100000U080400000002000000000800000004000000000E00000008000000000U0000000Tzzzzzzzzz00000000U"

    if 查找文字(332,763,150,150,descn,"**50",X,Y,OCR,0,0)
    ;------------------------------------------------------发现是空的
    {
	send, 零件
	}

}

return

;=================================
F7::
WinActivate ahk_exe CDMSImport.exe
return

mbutton::
value:="|<>58.zzzzzzzzzy000000000M000000001U000000006000000000Pk03A3U001YU04U20006G00G08008N8ktEkXMk1YYYZ4W4YU6GSEAC8GS0N910V8V901j3XW3zXnW6000000000M000000001U000000006000000000TzzzzzzzzzU"

if 查找文字(858,279,150,150,value,"**50",X,Y,OCR,0,0)
{
  CoordMode, Mouse
  MouseMove, X+50, Y
  click
  sleep,50
  send,^a
  sleep,50
}
return

numlock::
growt:="|<>39.CwQ0ey2GIE5JEUGW0e8A3YE3V0bIW0I84GIE2V0GGW0I81bPU2XVU"

if 查找文字(537,851,150,150,growt,"**50",X,Y,OCR,0,0)
{
  CoordMode, Mouse
  MouseMove, X+80, Y,0
  click
  sleep,50
  send,^a
  sleep,50
}
return

ins::
WinActivate, Microsoft Excel - D类加法工具
excel:= ComObjActive("Excel.Application")
FileName := "" 
excel.Columns("A:A").Select
excel.Selection.ClearContents
;~ kong:= sh.Range("A1:A10").value:=""
;~ excel.Range("A1").Select
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

;=======================================================================================
iWeb_complete(pwb)						;	returns bool for success or failure
{	
	If  !pwb							;	test to see if we have a valid interface pointer
		sleep, 200						;	ExitApp if we dont
	Else
	{
		loop 20							;	sets limit if itenerations to 40 seconds 80*500=40000=40 secs
			If not (rdy:=COM_Invoke(pwb,"readyState") = 4)
				Break				;	return success
			Else	Sleep,100					;	sleep .1 second between cycles
		loop 80							;	sets limit if itenerations to 40 seconds 80*500=40000=40 secs
			If (rdy:=COM_Invoke(pwb,"readyState") = 4)
				Break
			Else	Sleep,100					;	sleep half second between cycles
		Loop	80				
			If	((rdy:=COM_Invoke(pwb,"document.readystate"))="complete")
				Return 	1				;	return success
			Else	Sleep,100
	}
	Return 0						;	lets face it if it got this far it failed
}

;=======================================================================================
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
