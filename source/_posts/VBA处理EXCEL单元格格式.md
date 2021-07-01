---
title: VBA自动化处理数据4-单元格格式设置
date: 2021-07-01 08:23:32
tags: VBA
categories: 编程学习
---
# VBA处理EXCEL单元格格式

使用VBA可以快速处理一些格式设置。这里主要分三步

1. 循环EXCEL区域
2. 判定符合条件的单元格
3. 设定单元格格式

## 设置单元格格式

主要是针对填充单元格的颜色。使用的方法是Range().Interior.Color或者ColorIndex

```vb
sub setcellcolor()
    '设置区域A1:A10的颜色为绿色
    Range("A1:A10").Interior.Color = vbGreen
    '设置区域A1:A10的颜色为2即白色
    Range("A1:A10").Interior.ColorIndex = 2
end sub
```

> color也可以用RGB样式来引用。
>
> colorIndex是Excel内置的颜色排序，使用数字来对应不同的颜色。有56个颜色。

![Excel图片颜色](https://cdn.jsdelivr.net/gh/cemonliu/blogpic@main/360截图170010197210278.jpg)



## 设置字体样式

常用的字体设置有是否加粗，字体类型，字体大小，颜色等。

``` vb
sub setfont()
    '设置字体颜色为白色'
    Range("A1:A10").Font.Color = vbWhite
    '设置字体加粗'
    Range("A1:A10").Font.Bold = True
	'设置字体类型为Arial'
    Range("A1:A10").Font.name = "Arial"
    '设置字体大小为15'
    Range("A1:A10").Font.Size = 15
end sub

```



## 设置边框样式

边框使用Range的borders属性。

```vb
sub testborder()
    '设置斜上线'
    Range("D1").Borders(xlDiagonalUp).LineStyle = xlDash
    '设置斜下线实线'
    Range("D2").Borders(xlDiagonalDown).LineStyle = xlContinuous
    '设置底部线点线'
    Range("D3").Borders(xlEdgeBottom).LineStyle = xlDash
    '设置左部线实线'
    Range("D4").Borders(xlEdgeLeft).LineStyle = xlContinuous
    '设置右部线点线'
    Range("D5").Borders(xlEdgeRight).LineStyle = xlDash
    '设置顶部线双实线'
    Range("D6").Borders(xlEdgeTop).LineStyle = xlDouble
end sub

```

![样式图片](https://cdn.jsdelivr.net/gh/cemonliu/blogpic@main/image-20210701103602112.png)

## 最后放上思维导图

![思维导图](C:\Users\cemon_liu\Desktop\image-20210701103706267.png)