'vb创建获取当前输入法句柄，先添加一个文本框，然后双击后输入以下代码

 '首先，编写读取每种输入法键盘布局句柄的程序
    '声明Windows API函数
    Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal flags As Long) As Long
    Private Declare Function GetKeyboardLayout Lib "USER32.DLL" (ByVal dwLayout As Long) As Long
    '获取当前输入法的句柄
    Private Sub Form_Load()
    Dim dwLayout As Long
    dwLayout = GetKeyboardLayout(0)
    Text1.Text = dwLayout
    End Sub
