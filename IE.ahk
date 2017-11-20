;激活IE浏览器，在原来的窗口重复搜索复制的内容
WinActivate ahk_class IEFrame
Send !d ;原来的窗口输入网址alt+D
Sleep 50
Clipboard =  https://www.baidu.com/s?wd=%Clipboard%
Sleep 50
send  ^v{Enter}
