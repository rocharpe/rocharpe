
`::

zuo:=gethwnd(212, 12)
you:=gethwnd(1023, 11)

zuo1:=gethwnd(405, 316)
zuo2:=gethwnd(387, 352)
you1:=gethwnd(1027, 317)
you2:=gethwnd(1009, 344)
return
F4::
loop,5
{
;~ 虚拟按键码,0x22代表pgdn
;~ https://docs.microsoft.com/en-us/windows/desktop/inputdev/virtual-key-codes
;~ SendMessage, 0x100, 0x22, 0x014B0001,, ahk_id %zuo% ;WM_KEYDOWN := 0x100
;~ SendMessage, 0x101, 0x22, 0xC14B0001,, ahk_id %zuo% ;WM_KEYUP := 0x101

;~ SendMessage, 0x100, 0x22, 0x014B0001,, ahk_id %you% ;WM_KEYDOWN := 0x100
;~ SendMessage, 0x101, 0x22, 0xC14B0001,, ahk_id %you% ;WM_KEYUP := 0x101

Sleep,100
WinActivate, ahk_id %you%
Sleep,100
ControlSend , , {PGDN},ahk_id %you% ;左屏下一页

Sleep,100
WinActivate ,ahk_id %zuo% ;左屏下一页
Sleep,100
ControlSend , , {PGDN},ahk_id %zuo2% ;左屏下一页

Sleep,100
}

return




;======================================
gethwnd(ByRef xl,ByRef yl)
{
return DllCall( "WindowFromPoint", "int", xl, "int", yl )
}
