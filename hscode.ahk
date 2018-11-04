#SingleInstance,force
CoordMode,Pixel
CoordMode,mouse
ComObjError(false) ;关闭对象错误提示

`::

ControlGetText, tcnpm, , ahk_id %ocnpm% ;中文品名

csvfile=fy.csv
co1:=[]
co2:=[]

loop, read,%csvfile%
	{
		LineNumber = %A_Index%
		loop, parse, A_LoopReadLine, CSV
			co%A_Index%[LineNumber]:=A_LoopField
	}
	

for k,v in co1
if (tcnpm=v)
  {
  i:=k
  for key,hs in co2
  if ( key = i)
    {
    ControlSetText, ,%hs%,ahk_id %ohscode% ;trans是冒号后面的值
    }
  }

return

;==============================================================
F1::

MouseGetPos, newX, newY
ocnpm:=gethwnd(newX, newY)

return

F2::

MouseGetPos, newX, newY
ohscode:=gethwnd(newX, newY)

return



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

return
