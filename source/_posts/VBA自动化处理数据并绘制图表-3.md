---
title: VBA自动化处理数据并绘制图表-3
date: 2021-06-29 08:23:32
tags: VBA
categories: 编程学习
---
# VBA关于日期时间的处理

## 提取日期里的年月日

对于EXCEL里的日期，有常用函数year, month, day 来获取日期的年月日。在VBA里面也支持同样的函数。假如单元格A1里面的日期为2021/06/28。那我们可以用下面的方法获取年月日

```vb
'获取年
year=Year(Range("A1"))
'获取月
month=Month(Range("A1"))
'获取日
day=Day(Range("A1"))
```

## 日期之间的运算

### 两个日期之间的差

这里包含几个种类，比如两个日期之间的天数、两个日期之间的月份差、两个日期之间的年份差。对于这个的运算可以使用VBA DateDiff()这个函数。

>DateDiff(interval, date1, date2 [,firstdayofweek[, firstweekofyear]])
>
>*Interval* - 一个必需的参数。 它可以采用以下值。
>
>- *d* - 一年中的一天
>- *m* - 一年中的月份
>- *y* - 一年中的年份
>- *yyyy* - 年份
>- *w* - 工作日
>- *ww* - 星期
>- *q* - 季度
>- *h* - 小时
>- *m* - 分钟
>- *s* - 秒钟

假如在A1和B1分列存储两个日期。那么获取他们的两个日期之间间隔可以用下列方式。

```vb
'天数的差异
DateDiff("d",range("A1"),Range("B1"))
'月份的差异
DateDiff("m",range("A1"),Range("B1"))
'年份的差异
DateDiff("y",range("A1"),Range("B1"))
```

这个函数的好处在于，我们不用分别取年月日然后在进行换算得到差值。

### 日期月份的增加

假如我希望得到一个序列，分别是**2021/1/1,2021/2/1,2021/3/1...**. 这里最简单的做法是用EXCEL数据拖动的方式来实现。但是如果条件改变了，要获得的每个月最后一天的日期，即**2021/1/31,2021/2/28,2021/3/31...** 这里又如何实现呢？

在EXCEL里面有两个函数可以使用。分别是EDate()与EOMONTH()。这两个函数的解释如下。

1. Edate() 是与指定日期 (start_date) 相隔（之前或之后）指示的月份数。 使用函数 EDATE 可以计算与发行日处于一月中同一天的到期日的日期。

``` vb
EDATE(start_date, months)
'如果计算2021/1/31后面的一个月的日期,这里会得到2021/2/28
EDATE("2021/1/31",1)
```

> Start_date必需。 一个代表开始日期的日期。 应使用 DATE 函数输入日期，或者将日期作为其他公式或函数的结果输入。例如，使用函数 DATE(2008,5,23) 输入 2008 年 5 月 23 日。 如果日期以文本形式输入，则会出现问题。
> Months 必需。 start_date 之前或之后的月份数。 months 为正值将生成未来日期；为负值将生成过去日期。

2. EOMONTH() 可以计算正好在特定月份中最后一天到期的到期日。

``` vb
EOMONTH(start_date, months)
'2021/1/1后面一个月的最后一天，及2021/2/28
EOMONTH("2021/1/1", 1)
```

到这里，前面的序列问题就很简单了。我们可以设定一个循环来获取。

```vb

Sub GETDATE()
'设定循环
    For i = 0 To 11
	'起始月份
    dd = "2021/1/1"
    'VBA调用EXCEL函数EoMonth    
    newdate = WorksheetFunction.EoMonth(dd, i)
    
    '若起始日期为1月最后一天2021/1/31，那么使用Edate()也可以
    'newdate = WorksheetFunction.Edate(dd, i)
    '输出日期 cdate()是将数字转换为日期格式
    Debug.Print (CDate(newdate))
    Next    
End Sub

```

