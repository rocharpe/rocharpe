;调用英文输入法
SendMessage,0x50,0,67699721,,A
;弹出提示框，并输入需要的
InputBox ,time,标题,***F4启动***ESC暂停***`r`n***F5刷新***F11退出***`r`n`r`n***请关闭输入法***`r`n输入,,200,200
;绝对值坐标需要的参数，好像每个参数需要换行才生效
#SingleInstance,force
CoordMode,mouse
