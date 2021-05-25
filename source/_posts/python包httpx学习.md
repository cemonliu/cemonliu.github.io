---
title: python包httpx学习
date: 2021-05-25 14:23:32
tags: python
categories: 编程学习
---

## Python学习之路-httpx

### httpx包

之前在学习网络爬虫时，一直使用的是requests的包去抓取资料。最近在github看了一些API，发现已经使用httpx也越来越多。简单查询了相关资料，发现httpx能够实现同步和异步两部分功能。而requests却是同步功能的实现。

### 简单调用

最简单的实例即将之前requests包调用时，将requests换为httpx即可直接使用。

```python
import httpx
import requests
from pprint import pprint
#res=requests.get('https://www.baidu.com')
res=httpx.get('https://www.baidu.com')
pprint(res.text)
```



### 高级用法

因为异步的还没怎么用到。后续具体使用后再做参考。这里一个比较有改变的地方是，可以把**get**等方法作为参数使用。

使用httpx.Client建立类似requests的session功能。然后可以进行build_request。

```python
import httpx
from pprint import pprint
headers = {"X-Api-Key": "...", "X-Client-ID": "ABC123"}

with httpx.Client(headers=headers) as client:
    request = client.build_request("POST", "https://httpbin.org/post")
    pprint(request.headers["X-Client-ID"])  # "ABC123"
    response = client.send(request)
    pprint(response.text)
```





