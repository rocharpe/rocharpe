;========================================
;获取鼠标坐标
F2::
pwb := ComObjCreate("InternetExplorer.Application")
pwb.Visible := 1
pwb.Navigate("http://www.infoccsp.com/sso")
return


F4::
InputBox ,s,泰益物流,`r`n***F4    启动***`r`n***ESC  退出***`n`r`n输入单号位置,,200,200
if (s="")
{
s:=2
}
loop
{
#SingleInstance,force
CoordMode,Pixel
CoordMode,mouse
ComObjError(false) ;关闭对象错误提示
;设置
excel:= ComObjActive("Excel.Application")
FileName := "" ;无需指定文件名
Sh:=excel.Worksheets["Sheet1"] 

tt:=sh.Cells(s, 1).text
	if (tt<10000000)
		{
		gosub stop
		}
	else 
{
;======================================
;拆分单号
danhao:=SubStr(tt, 1, 3)"-"SubStr(tt, 4, 7)"` "SubStr(tt,0, 1)
ie:=IEGetFromUrl("www.infoccsp.com")  ;获取包含指定网页的一个选项卡
Sleep 300
ie.document.getElementById("frmright").contentDocument.getElementById("awbFullNo").value:=danhao ;运单号码赋值
Sleep 300
ie.document.getElementById("frmright").contentDocument.getElementById("btnNext").Click() ;点击下一步
sleep,300

;===================================================
;等待页面加载完毕
loop
{
if (ie.document.getElementById("frmright").contentDocument.readyState="complete" )
break
}
Sleep 300

;----------------------------------------------------
;判断发送结果
sendlist:=ie.document.getElementById("frmright").contentDocument.getElementById("sendListTipDiv").innerHTML
;~ MsgBox %sendlist%
if (InStr(sendlist,"发送成功"))
{
	sh.Cells(s, 3):="已发送"
	ie.document.getElementById("frmright").contentDocument.getElementById("btnNext").Click() ;下一票
	Sleep 300
	
	;===================================================
	;等待页面加载完毕
	loop
	{
	if (ie.document.getElementById("frmright").contentDocument.readyState="complete" )
	break
	}
	Sleep 300	
}
else
{
		;------------------------------------------------------------------------------
		;判断是否为修改
		change:=ie.document.getElementById("frmright").contentDocument.GetElementsByTagName("a").item(2).innerHTML 
		if (InStr(change,"修改"))
		{
			ie.document.getElementById("frmright").contentDocument.GetElementsByTagName("a").item(2).Click() ;点击修改
			Sleep 300
			
			;===================================================
			;等待页面加载完毕
			loop
			{
			if (ie.document.getElementById("frmright").contentDocument.readyState="complete" )
			break
			}
			Sleep 300	
			ie.document.getElementById("frmright").contentDocument.getElementById("btnSubmit").Click() ;点击提交
			Sleep 300
			
			;===================================================
			;等待页面加载完毕
			loop
			{
			if (ie.document.getElementById("frmright").contentDocument.readyState="complete" )
			break
			}
			Sleep 300		
	
					;-----------------------------------------------------------------------------
					;弹窗中有发送成功的界面
					tijiao:=ie.document.getElementById("frmright").contentDocument.GetElementsByTagName("li").item(0).innerHTML ;查看提交状态
					if (InStr(tijiao,"发送成功"))
					{
						sh.Cells(s, 3):="发送成功"	
						ie.document.getElementById("frmright").contentDocument.GetElementsByTagName("button").item(1).Click() ;关闭按钮
						Sleep 300
						;===================================================
						;等待页面加载完毕
						loop
						{
						if (ie.document.getElementById("frmright").contentDocument.readyState="complete" )
						break
						}
						Sleep 300	
						ie.document.getElementById("frmright").contentDocument.getElementById("btnNext").Click() ;下一票
						Sleep 300
						;===================================================
						;等待页面加载完毕
						loop
						{
						if (ie.document.getElementById("frmright").contentDocument.readyState="complete" )
						break
						}
						Sleep 300	
					}
					else
					{
						sh.Cells(s, 3):="发送状态未知"	
						ie.document.getElementById("frmright").contentDocument.GetElementsByTagName("button").item(1).Click() ;关闭按钮
						sleep,300			
						ie.document.getElementById("frmright").contentDocument.GetElementsByTagName("button").item(0).Click() ;关闭按钮
						Sleep 300
						;===================================================
						;等待页面加载完毕
						loop
						{
						if (ie.document.getElementById("frmright").contentDocument.readyState="complete" )
						break
						}
						Sleep 300	
						ie.document.getElementById("frmright").contentDocument.getElementById("btnNext").Click() ;下一票
						Sleep 300
						;===================================================
						;等待页面加载完毕
						loop
						{
						if (ie.document.getElementById("frmright").contentDocument.readyState="complete" )
						break
						}
						Sleep 300	
					}
		}
		else ;如果不是修改
		{
			sh.Cells(s, 3):="失败"	
			ie.document.getElementById("frmright").contentDocument.getElementById("btnNext").Click() ;下一票
			Sleep 300
			;===================================================
			;等待页面加载完毕
			loop
			{
			if (ie.document.getElementById("frmright").contentDocument.readyState="complete" )
			break
			}
			Sleep 300		
		}			
			
	}

}

s:=s+1
}

return


;~ webo:=ie.document.getElementById("page")
;~ if(webo=null)
;~ {
;~ ie:=IEGetFromUrl("www.baidu.com")  ;获取包含指定网页的一个选项卡

;~ search:="rocharpe"	

;~ ie.document.getElementById("kw").value:=search   ;通过ie.doucument对当前网页进行操作,kw为搜索框

;~ ie.document.getElementById("su").Click()
;~ }
;~ else
;~ {
;~ ie:=IEGetFromUrl("www.baidu.com")  ;获取包含指定网页的一个选项卡

;~ search:=""	

;~ ie.document.getElementById("kw").value:=search   ;通过ie.doucument对当前网页进行操作,kw为搜索框

;~ ie.document.getElementById("result_logo").Click()
;~ }




























































































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
	

	stop:
    ESc::
    ExitApp
    return
