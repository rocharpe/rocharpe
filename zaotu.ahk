#Include C:\test_game\find.ahk

+j::
gs:="|<>42.S000000W000800U000800UAqwQqvUGGG8MGUGGG8EIWGGG8EAQADv6s80000008000000kU"
zgs:="|<>46.01k1i0000102E000040B000QME0oNsOGF02mJGM740/7JDUYE0YZIVlzk7FxFs"
quer:="|<>23.0UEFtsEUYE111024Ti4Cd48py8ceYFFLseXeFcZIWF1311U"
yund:="|<>48.HyAADyzyM06M86A280TyDy8a7zNa86MaskTy00SY9gNazznz94TyA0m3/S1UDyHz/nzz06G2Q01U06S3rz1U3wSSU"
change:="|<>51.Ak0M0D061bnv01g0kNa3M7nrq2rUPwlUBDVb2n6A1aOFXvMggAnGMMP6ZVaOCn2kpMAnGQPA6alwOyOnkhutXMCIn6NEMPT3ABa66CU" 
tiaocu:="|<>51.SE8000002+0000004kE000000X2tDbz7DfCN9YnNBgUO98oFDcY3F96W914UO9AYFA8YSF9wW8x4s0080000000100000000800000U"

;找确定
if 查找文字(859,674,50,50,quer,"*113",X,Y,OCR,0,0)
{
  CoordMode, Mouse
  MouseMove, X, Y
  sleep,100
  MouseClick, left
  sleep,50
  MouseClick, left
  sleep,100
  MouseMove, 350,83 ;此处修改坐标
  sleep,300
  
;寻找加载框
Loop
  {
  PixelGetColor,tuy,631,525
  PixelGetColor,tue,535,532
  PixelGetColor,tus,890,536 
  if (tuy = "0xFFFFFF" and tue = "0xFFFFFF" and tus = "0xFFFFFF" )
  break
  } 
sleep,50

;~ 判断加载情况
done:="|<>59.s000s0r000U000E0Y0010000U1c0021VlV03FbVY4YYW05YeYc98740/7JDFGEG80GGeFyMQTw1oTIS"
loop
  {
  Click,350,83 ;此处同上
  Sleep,50
  Click,489,83
  Sleep,150
  if 查找文字(630,365,100,100,done,"**50",X,Y,OCR,0,0)
  Break
  }
sleep,100
}

;输入公司名 996,365;664,365
else
{
  zhuihg:="|<最惠国>35.Ds10zyEFxx04zU42zd11r48TzWG8Ed07QLRSw9+VGYdxx2JtE43zuN2YY0TZ52jzVFFxE1U"
  if 查找文字(467,867,50,50,zhuihg,"**50",X,Y,OCR,0,0)
    {
    Click , 417,739 ;品名的位置 661,739
    Sleep,50
    Send,{Enter}
    sleep,50
    Send,无规格
    sleep,100
    MouseClickDrag,L,996,375,650,375 ;公司的位置
    Sleep,50
    }

;提示
  else
    {
    zhuihg:="|<最惠国>35.Ds10zyEFxx04zU42zd11r48TzWG8Ed07QLRSw9+VGYdxx2JtE43zuN2YY0TZ52jzVFFxE1U"
    if not 查找文字(467,867,50,50,zhuihg,"**50",X,Y,OCR,0,0)
      {
      MsgBox 非最惠国
      }
      
      ;输入运单
      else
      {
      if 查找文字(274,57,50,50,yund,"**50",X,Y,OCR,0,0)
        {
        CoordMode, Mouse
        MouseMove, X+125, Y
        sleep,50
        MouseClick, left
        sleep,50
        MouseClick, left
        sleep,50
        }
      }
    }
  }
return


;保存，退出
+l::
change:="|<>51.Ak0M0D061bnv01g0kNa3M7nrq2rUPwlUBDVb2n6A1aOFXvMggAnGMMP6ZVaOCn2kpMAnGQPA6alwOyOnkhutXMCIn6NEMPT3ABa66CU" 
send, {F9}
sleep,100
loop
  {
  send, {F9}
  Sleep,50
  if 查找文字(535,106,500,50,change,"**50",X,Y,OCR,0,0)
  Break
  }
send, {F8}
sleep,100

tiaocu:="|<>51.SE8000002+0000004kE000000X2tDbz7DfCN9YnNBgUO98oFDcY3F96W914UO9AYFA8YSF9wW8x4s0080000000100000000800000U"
if 查找文字(326,77,50,50,tiaocu,"**50",X,Y,OCR,0,0)
{
  sleep,500
  CoordMode, Mouse
  MouseMove, X, Y
  sleep,100
  MouseClick, left
  sleep,200
}
MouseMove, 399, 60
sleep,200
MouseClick, left
sleep,20
MouseClick, left
sleep,50
return

;暂停
ESc::
Pause
return

#z::Reload
return

#t::Exitapp
return
