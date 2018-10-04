/*
;由鼠标获取控件ID,改为系统坐标获取ID,感觉更靠谱
;网页增加了新的判断方式

*/

#SingleInstance,force
CoordMode,Pixel
CoordMode,mouse
#Include C:\test_game\find.ahk
ComObjError(false) ;关闭对象错误提示

MsgBox,4097,获取CDMS参数, 按 确定 获取CDMS下参数信息`r`n`r`n需在CDMS编辑界面操作
IfMsgBox, ok 
    gosub CDMS 
return 

F1::
;==============================
CDMS:
CoordMode, Mouse
MouseMove ,967, 196
Sleep,100
send,{enter}
Sleep,100
MouseGetPos, , , cdms, control
gosub AA
gosub BB
;==============================

if (tcrno="1111960286" and tcrname="中外运·敦豪保税仓储（北京）有限公司") ;此处判断CDMS是否更新句柄
{
	MsgBox , 请确认`r`nHawb:  %thwab%`r`nShipper:  %tshipper%`r`nLocal_Name:  %tlocal%`r`nCR Name:  %tcrname%
}
else
{
	msgbox , 失败!!!`r`n1.CDMS是否有提示框存在？请关闭它`r`n`r`n2.最大化CDMS窗口于左屏
}
return

`::
/*
;=================================================
;点击NPTS的页面
nptsweb:="|<>42.1U000006900000AFU0000MgkMX0EF2EMn0EW209G0EXy89GCL8089GHNHzsBGTFG005AEF13k6AMNMwE6ADTA1U00003z00000U"

if 查找文字(1944,41,150,150,nptsweb,"**35",X,Y,OCR,0,0)
{
  CoordMode, Mouse
  MouseMove, X, Y
  sleep,50
  click
}
;=================================================
*/

;=================================================
;点击错误报警框,同时获取判断的参数HsCode:和Hawb:及其他
WinActivate ahk_id %cdms%
Sleep,50
WinActivate ahk_exe CDMSImport.exe
Sleep,50
ControlClick,Button1,ahk_exe CDMSImport.exe
Sleep,50
send,{enter}
Sleep,50

iWeb_Activate("NPTS - A DHL Product") 
iWeb_Activate("WebFSQ - ShipmentDetails") 
Sleep,50

gosub BB
;=================================================

if (tcrno="1111960286" and tcrname="中外运·敦豪保税仓储（北京）有限公司") ;此处判断CDMS是否更新布局
{
WinActivate ahk_class IEFrame 
ie:=IEGetFromUrl("npts.dhl.com")  ;获取包含指定网页的一个选项卡
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

;判断是否是DHLE特殊公司
if (InStr(tshipper,"dhle ") )
{
MsgBox 发件公司:DHLE
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
if  (InStr(tlocal,"VENTIVE")) 
{
MsgBox 苏州的SYNVENTIVE需要核重
}
sleep,50

;特殊公司GE/JABIL
if (InStr(tlocal,"ge ")="1" or InStr(tlocal,"jabil") or InStr(tlocal,"GEHC")) 
{
MsgBox 注意特殊公司`r`nGE 78 JINYING ROAD`r`nJABIL 600 ROAD
}
sleep,50

;可能是GE公司
if (InStr(taddr,"325 ") or InStr(taddr,"warranty") or InStr(taddr,"508 ") or InStr(taddr,"JinYing") or InStr(taddrr,"JinYing"))
{
MsgBox 可能是 GE 78 JinYing Road
}
sleep,50

;三家重量特殊处理的公司
if  (InStr(tlocal,"FUTURE") or InStr(tlocal,"COVANCE") or InStr(tlocal,"Hisilicon")) 
{
MsgBox 以下公司不需要核重`r`n`r`nFUTURE ELECTRONICS 重量看发票`r`n`r`nCOVANCE 重量认运单`r`n`r`nHisilicon 重量认NTPS
}
sleep,50
;===========================================================================================================


WinActivate ahk_class IEFrame 
While ie.readyState != 4 || ie.document.readyState != "complete" || wb.busy
sleep,200
ie.document.GetElementsByTagName("input").item(4).value:=thwab  ;通过ie.doucument对当前网页进行操作,kw为搜索框
sleep,50
ie.document.GetElementsByTagName("searchButton").item(0).Click()
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

pcs:=ie.document.getElementsByTagName("table").item(4).getElementsByClassName("whiteTdNormal").item(4).getElementsByTagName("a").item(0).innerHTML

;件数和CDMS比较
if (ceil(pcs)-ceil(tpcs)<>0)
{
MsgBox 件数不同`r`n`r`n NPTS :  %pcs%`r`nCDMS :  %tpcs%
}


  
BlockInput, MouseMove
;======================================================================
;不管有没有hs,先做空处理
WinActivate ahk_id %cdms%
ControlSetText, ,,ahk_id %ohscode%
Sleep,50
;如果品名为空则设置为零件
if (tdesc="")
{
ControlSetText, ,零件,ahk_id %odesc%
}
;选中中文品名
ControlClick,,ahk_id %odesc%
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
WinActivate ahk_id %cdms%
return

;=================================================
;打勾标九
F4::
if (tdanwei="")
{
ControlClick,,ahk_id %odanwei%
ControlSetText, ,千克,ahk_id %odanwei%
Sleep,100
send,{enter}
Sleep,100
}

ControlClick,,ahk_id %ocsz%
ControlClick,,ahk_id %ocsy%
ControlSetText, ,9,ahk_id %ocsy%
send,{enter}
Sleep,100
ControlClick,,ahk_id %ohwab%
return

;===============================================
;定位到金额
mbutton::
WinActivate ahk_id %cdms%
ControlClick,,ahk_id %odecvalue%
sleep,50
send,^a
return

;============================================
;定位到毛重
numlock:: 
WinActivate ahk_id %cdms%
ControlClick,,ahk_id %ogrowt%
sleep,50
send,^a
return

;========================================
;激活excel
ins::
ComObjError(false) ;关闭对象错误提示
WinActivate, ahk_class XLMAIN
i:=0
DD:
i:=i+1
IfWinActive, Microsoft Excel - D类
	{
	excel:= ComObjActive("Excel.Application")
	FileName := "" 
	excel.Columns("A:A").Select
	excel.Selection.ClearContents
	;~ kong:= sh.Range("A1:A10").value:=""
	;excel.Range("A1").Select
	}
else
	{
	if (i<5)
		{
		send ^{tab}
		Sleep 50
		goto DD
		}
		else
		{
		MsgBox 未找到文件
		}
	}
return

;~ ins::
;~ WinActivate, Microsoft Excel - D类加法
;~ excel:= ComObjActive("Excel.Application")
;~ FileName := "" 
;~ excel.Columns("A:A").Select
;~ excel.Selection.ClearContents
;~ ;kong:= sh.Range("A1:A10").value:=""
;~ ;excel.Range("A1").Select
;~ return

;========================================
;激活Pdf文件及轮流切换Excel和pdf
~CapsLock::
SetTitleMatchMode,2
IfWinActive, ahk_class AcrobatSDIWindow
	WinActivate, ahk_class XLMAIN
else
	WinActivate, ahk_class AcrobatSDIWindow
return

;========================================
;pdf文件操作
#IfWinActive,ahk_class AcrobatSDIWindow
{
 ~Numpad0::^0 	 ;还原
 ~Numpad5::^+-	 ;逆时针
 ~Numpad6::^++ 	 ;顺时针
 ~Numpad4::PGDN  ;下一页
 ~Numpad7::PGUP	 ;上一页
 ~Numpad8::^NumpadAdd ;放大
}
#IfWinActive
 return
 
;========================================
;获取鼠标坐标
F2::
CoordMode, Mouse
MouseGetPos, x, y
s=%x%, %y%
Clipboard:=s
ToolTip, 抓点成功！可直接粘贴代码！
Sleep, 1000
ToolTip
return

AA:
;==================================================
;获取坐标对应的控件名
BlockInput, MouseMove

opcs:=gethwnd(733, 221) ;件数
ohwab:=gethwnd(967, 196) ;运单号
ocsz:=gethwnd(982, 241) ;cs左
ocsy:=gethwnd(1026, 243) ;cs右
odecvalue:=gethwnd(957, 280) ;货值

oshipper:=gethwnd(378, 346) ;发件人
olocal:=gethwnd(748, 390) ;收件人
oaddr:=gethwnd(706, 413) ;地址一
oaddrr:=gethwnd(430, 408) ;地址二

ocrno:=gethwnd(382, 480) ;海关编码 1111960286
ocrname:=gethwnd(585, 481) ;中外运·敦豪保税仓储（北京）有限公司

;Hawb 904, 197 
ohscode:=gethwnd(373, 735) ;商品编码
odesc:=gethwnd(366, 778) ;中文品名
odanwei:=gethwnd(447, 841) ;单位
ogrowt:=gethwnd(646, 862) ;毛重

BlockInput, MouseMoveOff
return


BB:
;=============================================================
;获取文本
ControlGetText, tpcs, , ahk_id %opcs% ;件数
ControlGetText, tcrno, , ahk_id %ocrno% ;海关编码
ControlGetText, tcrname, , ahk_id %ocrname% ;DHL

ControlGetText, thwab, , ahk_id %ohwab% ;运单号
ControlGetText, tdecvalue, , ahk_id %odecvalue% ;货值
ControlGetText, tshipper, , ahk_id %oshipper% ;发件人
ControlGetText, tlocal, , ahk_id %olocal% ;收件人
ControlGetText, taddr, , ahk_id %oaddr% ;地址一
ControlGetText, taddrr, , ahk_id %oaddrr% ;地址二
ControlGetText, thscode, , ahk_id %ohscode% ;商品编码
ControlGetText, tdesc, , ahk_id %odesc% ;中文品名
ControlGetText, tdanwei, , ahk_id %odanwei% ;单位
ControlGetText, tgrowt, , ahk_id %ogrowt% ;毛重

return

CC:
;=============================================================
;获取文本
ControlGetText, tpcs, %opcs%, ahk_id %cdms% ;件数
ControlGetText, tcrno, %ocrno%, ahk_id %cdms% ;海关编码
ControlGetText, tcrname, %ocrname%, ahk_id %cdms% ;DHL

ControlGetText, thwab, %ohwab%, ahk_id %cdms% ;运单号
ControlGetText, tdecvalue, %odecvalue%, ahk_id %cdms% ;货值
ControlGetText, tshipper, %oshipper%, ahk_id %cdms% ;发件人
ControlGetText, tlocal, %olocal%, ahk_id %cdms% ;收件人
ControlGetText, taddr, %oaddr%, ahk_id %cdms% ;地址一
ControlGetText, taddrr, %oaddrr%, ahk_id %cdms% ;地址二
ControlGetText, thscode, %ohscode%, ahk_id %cdms% ;商品编码
ControlGetText, tdesc, %odesc%, ahk_id %cdms% ;中文品名
ControlGetText, tdanwei, %odanwei%, ahk_id %cdms% ;单位
ControlGetText, tgrowt, %ogrowt%, ahk_id %cdms% ;毛重

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
gethwnd(ByRef xl,ByRef yl)
{
return DllCall( "WindowFromPoint", "int", xl, "int", yl )
}


gethwndd(x,y)	
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
