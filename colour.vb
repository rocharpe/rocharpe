http://club.excelhome.net/thread-513353-1-1.html
'统计颜色 F9刷新工作表
'选中当前工作表的A1单元格，然后定义一个好记的名称，如qjq，然后输入
=GET.CELL(63,Sheet1!A1)
'在sheet2的A1单元格里输入
=qjq
'然后进行左右上下拉伸，面积和工作表一样大小或更大
