'横向比对公式，先选择需要显示的一列，再在条件格式里选择公式
=IF(AND(A1=G1,B1=H1,C1=I1,D1=J1),0,1)
  
'显示重复项
'1、选中Sheet1的A:A（假设第一行是标题）
'2、“条件格式”，在打开的对话窗口中做如下操作：
'第一格：拉下来选"公式"
'第二格：输入公式 
'（其中sheet2是数据库，$A:$A为数据库所在位置，A1是开始选中列的起时位置）
  =countif(sheet2!$A:$A,A1)=1

  '表格多列统计到一列宏代码，大概的意思，选中AD1，然后在里面输入公式，完事后选中AD1，复制，再选中AD1到AD11000，黏贴
  '公式的含义，R代表行，R3就是第三行开始，C代表列，后面的[]代表当前单元格的偏移量，即相对于AD1的偏移量是向左-29格
  '后面的22代表需要统计22列的数据
  '=OFFSET(A$3,TRUNC((ROW()-1)/22),MOD(ROW()-1,22)) 原形
  Sub 按钮4_Click()
    Range("AD1").Select
    ActiveCell.FormulaR1C1 = _
        "=OFFSET(R3C[-29],TRUNC((ROW()-1)/22),MOD(ROW()-1,22))"
    Range("AD1").Select
    Selection.Copy
    Range("AD1:AD11000").Select
    ActiveSheet.Paste
  End Sub

        
  ● =IF(ISNA(VLOOKUP(A3,$D:$E,2,0)),"",VLOOKUP(A3,$D:$E,2,0))            ISNA可以去掉错误，使表格看起来舒服，公式意思是如果说VLOOKUP没有找到内容，就返回""空，否则，返回找到的内容
  ● =VLOOKUP(A2,$D:$E,2,0) A2指查找谁，$D:$E指在哪里查找（查找区域必须在这第一列），打上美元符号指固定查找位置，不然会联动，2指的是查找区域第二列，0是精确查找
  ● =VLOOKUP(A2&"*",$D:$E,2,0)   VLOOKUP通配符查找 
  ● =COLUMN() 在哪里就是第几列，=COLUMN(C1)返回值为3
  ● =MATCH(B8,$1:$1,0) B8指查找谁，$1:$1指在哪里查找且固定位置，绝对引用，0指精确匹配，返回的是B8在$1:$1中排的第几位数值
  ● =LEFT(A1,LENB(A1)-LEN(A1))  中文数字分离
  ● Alt＋ 求和
  ● =COUNTIF(A$1:A1,A1)  重复值
  ● =COUNTIF(C1:C1500,"1") C列1数和
  ● =IF(A1=1,1," ") 提取A列中的1
  ● Alt+Enter在表格内换行
  ● Ctrl+Shift+上/下，选择该列所有数据
  ● Ctrl+上/下，跳至表格最上下方
  ● Ctrl+C/V，不仅仅复制表格内容（格式和公式）
  ● Ctrl+D/R，复制上行数据/左列数据
  ● F4
  ● Ctrl+Shift+～ 常规
  ● Ctrl+Shift+！ 数值
  ● 显示隐藏列，选择1后按Ctrl+Shift+下，右击鼠标选最后一项U
  ● =LOOKUP(1,0/($A$14:$A$22=F17)*(B16:B24=G17),$C$14:$C$22)  寻找1，在数组里找0，lookup不会查找错误值，0/0会返回错误值，所以符合条件的就会是0，最后一个是在返回的值        
