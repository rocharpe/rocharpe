/*
修改坐标,通过F2获取坐标点

ohwabb:=gethwnd(904, 196) ;hwab字母位置(用来判断)
ohwab:=gethwnd(967, 196) ;运单号
odecvalue:=gethwnd(957, 280) ;货值
oshipper:=gethwnd(388, 344) ;发件人
olocal:=gethwnd(762, 366) ;收件人
oaddr:=gethwnd(369, 385) ;地址一
oaddrr:=gethwnd(697, 385) ;地址二
ohscodee:=gethwnd(314, 721) ;hscode字母位置(用来判断)
ohscode:=gethwnd(380, 723) ;商品编码
odesc:=gethwnd(381, 763) ;中文品名
ogrowt:=gethwnd(636, 846) ;毛重
ocrno:=gethwnd(356, 433) ;海关编码 1111960286
ocrname:=gethwnd(550, 433) ;中外运敦豪保税仓储（北京）有限公司
ControlGetText, tcrno, %ocrno%, ahk_exe CDMSImport.exe
ControlGetText, tcrname, %ocrname%, ahk_exe CDMSImport.exe

*/

MsgBox,4097,获取CDMS参数, 按 确定 获取CDMS下参数信息`r`n`r`n需在CDMS编辑界面操作
IfMsgBox, ok 
    gosub CDMS 
return 

F1::
CDMS:
CoordMode, Mouse
WinActivate ahk_exe CDMSImport.exe
Sleep,500
ControlClick,Button1,ahk_exe CDMSImport.exe
Sleep,50

gosub AA
gosub BB

;==================================
if (tcrno="1111960286" and tcrname="中外运敦豪保税仓储（北京）有限公司") ;此处判断CDMS是否更新句柄
{
	MsgBox , 请确认`r`nHawb:  %thwab%`r`nShipper:  %tshipper%`r`nLocal_Name:  %tlocal%`r`nCR Name:  %tcrname%
}
else
{
	msgbox , 失败!!!即将退出`r`n1.CDMS是否有提示框存在？请关闭它`r`n`r`n2.请确认是否在MP状态下打开？`r`n`r`n3.还是不行？那就需要更新坐标点
	ExitApp
}
return

F2::
CoordMode, Mouse
MouseGetPos, x, y
s=%x%, %y%
Clipboard:=s
ToolTip, 抓点成功！可直接粘贴代码！
Sleep, 1000
ToolTip
return



#a::
#SingleInstance,force
CoordMode,Pixel
CoordMode,mouse
;~ #Include C:\test_game\find.ahk
;~ #Include C:\test_game\com.ahk
ComObjError(false) ;关闭对象错误提示

;=================================================
;点击错误报警框,同时获取判断的参数HsCode:和Hawb:及其他
ControlClick,Button1,ahk_exe CDMSImport.exe
Sleep,50
WinActivate ahk_exe CDMSImport.exe
Sleep,100
ControlClick,Button1,ahk_exe CDMSImport.exe
Sleep,50
gosub BB
;=================================================

if (tcrno="1111960286" and tcrname="中外运敦豪保税仓储（北京）有限公司") ;此处判断CDMS是否更新布局
{
WinActivate ahk_class IEFrame 
ie:=IEGetFromUrl("npts2.apis.dhl.com")  ;获取包含指定网页的一个选项卡
ie.document.GetElementsByTagName("a").item(1).Click()

;==========发件人
;判断私人件
if (InStr(tshipper,"crg") or InStr(tshipper,"jetmile") or InStr(tshipper,"sinophile"))
{
MsgBox 可能需要转C类(CRG/JETMILE/SINOPHILE)
}
sleep,50

;判断是否是GE特殊公司
if (InStr(tshipper,"ge ") and InStr(taddr,"78")) 
{
MsgBox GE 78 JINYING ROAD？
}
sleep,50

;==========收件人
;判断私人件
if (InStr(tlocal,"crg") or InStr(tlocal,"jetmile") or InStr(tlocal,"source") or InStr(tlocal,"sinophile"))
{
MsgBox 可能需要转C类(CRG/JETMILE/SINOPHILE)
}
sleep,50

;核件核重
if (InStr(tlocal,"getin") )
{
MsgBox 注意公司`r`n`r`nGETINGE件数有X需核件核重
}
sleep,50

;核重
if  (InStr(tlocal,"dresser") or InStr(tlocal,"dms") or InStr(tlocal,"SYNVENTIVE")) 
{
MsgBox 苏州的dresser、dms、SYNVENTIVE需要核重
}
sleep,50

;特殊公司GE/JABIL
if (InStr(tlocal,"ge ")="1" or InStr(tlocal,"jabil")) 
{
MsgBox 注意特殊公司`r`nGE 78 JINYING ROAD`r`nJABIL 600 ROAD
}
sleep,50

;可能是GE公司
if (InStr(taddr,"325 ") or InStr(taddr,"warranty") or InStr(taddr,"508 ") or InStr(taddrr,"JinYing"))
{
MsgBox 可能是 GE 78 JinYing Road
}
sleep,50

;三家重量特殊处理的公司
if  (InStr(tlocal,"FUTURE") or InStr(tlocal,"COVANCE") or InStr(tlocal,"Hisilicon")) 
{
MsgBox FUTURE ELECTRONICS 重量看发票`r`n`r`nCOVANCE 重量认运单,不用核重`r`n`r`nHisilicon 重量认NTPS
}
sleep,50
;===========================================================================================================


WinActivate ahk_class IEFrame 
While ie.readyState != 4 || ie.document.readyState != "complete" || wb.busy
sleep,200
ie.document.GetElementsByTagName("input").item(4).value:=thwab  ;通过ie.doucument对当前网页进行操作,kw为搜索框
send ,^+S
While ie.readyState != 4 || ie.document.readyState != "complete" || wb.busy
  sleep,200
  send,{enter}
  sleep,300
  send,{enter}
  sleep,300
  send,{enter}
  
BlockInput, MouseMove
;======================================================================
;不管有没有hs,先做空处理
WinActivate ahk_exe CDMSImport.exe
ControlSetText, %ohscode%,,ahk_exe CDMSImport.exe
Sleep,50
;如果品名为空则设置为零件
if (tdesc="")
{
ControlSetText, %odesc%,零件,ahk_exe CDMSImport.exe
}
;选中中文品名
ControlClick,%odesc%,ahk_exe CDMSImport.exe
sleep,50
send,^a
}
else
{
	msgbox , 失败!!!`r`n`r`n按 F1 重试`r`n%thscodee% `r`n%thwabb%
	
}
BlockInput, MouseMoveOff
return

;=================================
;激活CDMS
F7::
WinActivate ahk_exe CDMSImport.exe
return

;===============================================
;定位到金额
mbutton::
WinActivate ahk_exe CDMSImport.exe
ControlClick,%odecvalue%,ahk_exe CDMSImport.exe
sleep,50
send,^a
return

;============================================
;定位到毛重
numlock:: 
WinActivate ahk_exe CDMSImport.exe
ControlClick,%ogrowt%,ahk_exe CDMSImport.exe
sleep,50
send,^a
return

;========================================
;激活excel
ins::
WinActivate, Microsoft Excel - D类加法
excel:= ComObjActive("Excel.Application")
FileName := "" 
excel.Columns("A:A").Select
excel.Selection.ClearContents
;~ kong:= sh.Range("A1:A10").value:=""
;excel.Range("A1").Select
return



AA:
;==================================================
;获取坐标对应的控件名
BlockInput, MouseMove

;ohwabb:=gethwnd(904, 196) ;hwab字母位置(用来判断)
ohwab:=gethwnd(967, 196) ;运单号
odecvalue:=gethwnd(957, 280) ;货值
oshipper:=gethwnd(388, 344) ;发件人
olocal:=gethwnd(762, 366) ;收件人
oaddr:=gethwnd(369, 385) ;地址一
oaddrr:=gethwnd(697, 385) ;地址二

ocrno:=gethwnd(356, 433) ;海关编码 1111960286
ocrname:=gethwnd(550, 433) ;中外运敦豪保税仓储（北京）有限公司

;ohscodee:=gethwnd(314, 721) ;hscode字母位置(用来判断)
ohscode:=gethwnd(380, 723) ;商品编码
odesc:=gethwnd(381, 763) ;中文品名
ogrowt:=gethwnd(636, 846) ;毛重


BlockInput, MouseMoveOff
return


BB:
;=============================================================
;获取文本
ControlGetText, tcrno, %ocrno%, ahk_exe CDMSImport.exe
ControlGetText, tcrname, %ocrname%, ahk_exe CDMSImport.exe

ControlGetText, thwabb, %ohwabb%, ahk_exe CDMSImport.exe
ControlGetText, thwab, %ohwab%, ahk_exe CDMSImport.exe
ControlGetText, tdecvalue, %odecvalue%, ahk_exe CDMSImport.exe
ControlGetText, tshipper, %oshipper%, ahk_exe CDMSImport.exe
ControlGetText, tlocal, %olocal%, ahk_exe CDMSImport.exe
ControlGetText, taddr, %oaddr%, ahk_exe CDMSImport.exe
ControlGetText, taddrr, %oaddrr%, ahk_exe CDMSImport.exe
;ControlGetText, thscodee, %ohscodee%, ahk_exe CDMSImport.exe
ControlGetText, thscode, %ohscode%, ahk_exe CDMSImport.exe
ControlGetText, tdesc, %odesc%, ahk_exe CDMSImport.exe
ControlGetText, tgrowt, %ogrowt%, ahk_exe CDMSImport.exe

;~ ControlGetText, thwabb, %ohwabb%, ahk_exe 大漠综合工具.exe
;~ ControlGetText, thwab, %ohwab%, ahk_exe 大漠综合工具.exe
;~ ControlGetText, tdecvalue, %odecvalue%, ahk_exe 大漠综合工具.exe
;~ ControlGetText, tshipper, %oshipper%, ahk_exe 大漠综合工具.exe
;~ ControlGetText, tlocal, %olocal%, ahk_exe 大漠综合工具.exe
;~ ControlGetText, taddr, %oaddr%, ahk_exe 大漠综合工具.exe
;~ ControlGetText, taddrr, %oaddrr%, ahk_exe 大漠综合工具.exe
;~ ControlGetText, thscodee, %ohscodee%, ahk_exe 大漠综合工具.exe
;~ ControlGetText, thscode, %ohscode%, ahk_exe 大漠综合工具.exe
;~ ControlGetText, tdesc, %odesc%, ahk_exe 大漠综合工具.exe
;~ ControlGetText, tgrowt, %ogrowt%, ahk_exe 大漠综合工具.exe

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
