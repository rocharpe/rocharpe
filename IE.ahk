;激活IE浏览器，在原来的窗口重复搜索复制的内容
WinActivate ahk_class IEFrame ;激活IE浏览器
Send !d ;定位网址窗口，输入网址alt+D
;============================================

#NoEnv
#Persistent
#SingleInstance, force

;=================
;关闭对象错误提示
ComObjError(false) 

;=========================
;激活ie指定窗口,此处激活百度
iWeb_Activate("百度")

;===============================
;获取包含https://www.baidu.com网页的一个选项卡
ie:=IEGetFromUrl("www.baidu.com") 

;==========================================================================
;等待网页加载完毕
While ie.readyState != 4 || ie.document.readyState != "complete" || wb.busy
    Sleep, 100
    
;================
;定义要搜索的变量
search:="rocharpe"

;=============================================
;通过ie.doucument对当前网页进行操作,kw为搜索框
ie.document.getElementById("kw").value:=search  

;==============================================================================
;点击搜索,其中的item(0)代表第一次出现
ie.document.GetElementsById("su").GetElementsByTagName("submit").item(0).Click() 
send , {enter}

;========================================================================
;点击百度首页新闻，即id=u1，标签为a第一次出现的链接
ie.document.getElementById("u1").GetElementsByTagName("a").item(0).Click() 

;====================================================================================
;获取链接下的内容
biaoqian:=ie.document.getElementById("u1").GetElementsByTagName("a").item(3).innerHTML
texta:=ie.document.getElementById("su").value
sen:=ie.document.getElementById("lh").GetElementsByTagName("a").item(0).innerHTML







;==================================================================================
;===================================函数部分=======================================
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


;============================================================================================
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

