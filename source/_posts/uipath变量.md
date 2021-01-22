---
title: uipath变量
date: 2021-01-21 14:56:08
tags: uipath
categories: 编程学习
---

## 变量

### 1.变量属性

- Name
- Type
- Default Value
- Scope

### 2. 创建变量的方法

- **From the Variables panel** – Open the Variables panel, select the ‘Create new Variable’
- **From the Designer panel** – Drag an activity with a variable field visible (i.e. ‘Assign’) and press **Ctrl+K**. 
- **From the Properties panel** – In the Properties panel of the activity, place the cursor in the field in which the variable is needed (i.e. Output) and press **Ctrl+K.**

![](https://cdn.jsdelivr.net/gh/cemonliu/blogpic@main/20210121151613.png)

### 3. 常用数据类型

- Numeric

  - Int32
  - Long
  - Double

- Boolean

- Date and Time

  - DateTime 常用存储方式为(mm/dd/yyyy hh:mm:ss)
  - TimeSpan常用存储方式为(mm/dd/yyyy hh:mm:ss)

- String  简单的文本类型

- Collection

  - Array
  - List
  - Dictionary

- GenericValue 任意数值类型 如果是抓取元素里的值一般采用此类型

  > 初始化一个数组并显示出来

![](https://cdn.jsdelivr.net/gh/cemonliu/blogpic@main/20210121153916.png)