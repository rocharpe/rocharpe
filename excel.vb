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
  Sub 按钮4_Click()
    Range("AD1").Select
    ActiveCell.FormulaR1C1 = _
        "=OFFSET(R3C[-29],TRUNC((ROW()-1)/22),MOD(ROW()-1,22))"
    Range("AD1").Select
    Selection.Copy
    Range("AD1:AD11000").Select
    ActiveSheet.Paste
  End Sub
