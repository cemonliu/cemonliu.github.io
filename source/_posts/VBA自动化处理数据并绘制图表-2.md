---
title: VBA自动化处理数据并绘制图表-2
date: 2021-06-17 14:23:32
tags: VBA
categories: 编程学习
---
# VBA自动化处理数据并绘制图表-2

## VBA的一些介绍

前后用了一周左右的时间来完成这个功能，VBA相比其他语言来说，在EXCEL编辑上优势还是很大的。虽然python有很多轮子，但是VBA的录制宏功能不可谓之不强。当我想去实现一个功能时，可能我记不住属性，记不住方法。但是我手工操作一遍，**后台的录制宏会自动写好代码**。这时候我们要做的就是删除庸余的代码，添加循环即可。



## 筛选数据和复制粘贴到新的EXCEL文件中

### 存储筛选条件

在VBA存储数据时，有数组和字典类型。我个人比较喜欢字典类型的存储。它是属于mapping关系。一个字典可以通过add方法添加item和key值。它的优势在于：

1. 字典读取速度快
2. 取值方便dic(item)就可以快速获取值

所以我在这里也使用了字典的来存取筛选条件。将需要筛选的值放到data sheet里面，使用字典进行存取。

```vb
'创建一个字典
Set dic = CreateObject("scripting.dictionary")
Sheets("data").Activate
    '循环读取值并存入字典中
    For i = 1 To Cells(Rows.Count, "A").End(xlUp).row
        dic.Add i, Cells(i, 1)
    Next
```

### excel创建新的sheet与excel 文件

#### 创建sheet

我们可以使用sheets的add方法添加一个worksheet在Excel文件里

```vb
'添加一个sheet 在sheets(data)前面
Sheets.Add before:=Sheets("data")
'重命名sheet 方便后面读取
ActiveSheet.Name = "Topdata"
```

#### 创建workbook

使用workbooks的add方法添加一个excel文件，但是在重命名时却报错。后面就使用workbook的saveas方法来定义文件的名称

```vb
'新增workbook
Workbooks.Add
'另存为test.xlsx
ActiveWorkbook.SaveAs Filename:="D:\test.xlsx"
ActiveWorkbook.Close
```

### 筛选

我们可以使用循环来设定excel的filter。这里假设第一列为筛选区域，晒选的结果贴到sheets(2)里面去

```vb
for i=1 to dic.count
    Sheets("Raw data").Rows("1:1").Select
    Selection.AutoFilter
    Selection.AutoFilter Field:=1, Criteria1:=dic(i)
    selection.copy sheets(2).range("A1")
next
```

