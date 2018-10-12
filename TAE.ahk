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
;~ ie:=IEGetFromUrl("www.infoccsp.com")  ;获取包含指定网页的一个选项卡

;======================================
;获取当前页的文档
ie:=WBGet() 
Sleep 300
ie.document.getElementById("frmright").contentDocument.getElementById("awbFullNo").value:=danhao ;运单号码赋值
Sleep 300
ie.document.getElementById("frmright").contentDocument.getElementById("btnNext").Click() ;点击下一步
sleep,300
;~ next:=ie.document.getElementById("frmright").contentDocument.getElementById("awbFullNo").value
;~ MsgBox %next%
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
		;~ MsgBox %change%
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







;========================================
;获取鼠标坐标
F2::
ie := ComObjCreate("InternetExplorer.Application")
ie.Visible := 1
ie.Navigate("http://www.infoccsp.com/sso/ui-agent_index.do")
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

WBGet(WinTitle="ahk_class IEFrame", Svr#=1) {               ;// based on ComObjQuery docs
   static msg := DllCall("RegisterWindowMessage", "str", "WM_HTML_GETOBJECT")
        , IID := "{0002DF05-0000-0000-C000-000000000046}"   ;// IID_IWebBrowserApp
;//     , IID := "{332C4427-26CB-11D0-B483-00C04FD90119}"   ;// IID_IHTMLWindow2
   SendMessage msg, 0, 0, Internet Explorer_Server%Svr#%, %WinTitle%
   if (ErrorLevel != "FAIL") {
      lResult:=ErrorLevel, VarSetCapacity(GUID,16,0)
      if DllCall("ole32\CLSIDFromString", "wstr","{332C4425-26CB-11D0-B483-00C04FD90119}", "ptr",&GUID) >= 0 {
         DllCall("oleacc\ObjectFromLresult", "ptr",lResult, "ptr",&GUID, "ptr",0, "ptr*",pdoc)
         return ComObj(9,ComObjQuery(pdoc,IID,IID),1), ObjRelease(pdoc)
      }
   }
}



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



	stop:
    ESc::
    ExitApp
    return
