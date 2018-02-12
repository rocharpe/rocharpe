;激活IE浏览器，在原来的窗口重复搜索复制的内容
WinActivate ahk_class IEFrame ;激活IE浏览器
Send !d ;定位网址窗口，输入网址alt+D
;============================================

#NoEnv
#Persistent
#SingleInstance, force
ComObjError(false) ;关闭对象错误提示
ie:=IEGetFromUrl("www.baidu.com")  ;获取包含https://www.baidu.com网页的一个选项卡

;定义要搜索的变量
search:="rocharpe"	
ie.document.getElementById("kw").value:=search   ;通过ie.doucument对当前网页进行操作,kw为搜索框
;点击搜索,其中的item(0)代表第一次出现
ie.document.GetElementsById("su").GetElementsByTagName("submit").item(0).Click() 
send , {enter}

;点击百度首页新闻，即id=u1，标签为a第一次出现的链接
sleep , 3800
ie.document.getElementById("u1").GetElementsByTagName("a").item(0).Click() 
;获取链接下的内容
biaoqian:=ie.document.getElementById("u1").GetElementsByTagName("a").item(3).innerHTML
texta:=ie.document.getElementById("su").value
sen:=ie.document.getElementById("lh").GetElementsByTagName("a").item(0).innerHTML
MsgBox %biaoqian% %texta% %sen%


;===========================================
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
