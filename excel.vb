'横向比对公式，先选择需要显示的一列，再在条件格式里选择公式
=IF(AND(A1=G1,B1=H1,C1=I1,D1=J1),0,1)
  
'显示重复项
1、选中Sheet1的A:A（假设第一行是标题）
2、“条件格式”，在打开的对话窗口中做如下操作：
第一格：拉下来选"公式"
第二格：输入公式 =countif(sheet2!$A:$A,A1)=1
（其中sheet2是数据库，$A:$A为数据库所在位置，A1是开始选中列的起时位置）
