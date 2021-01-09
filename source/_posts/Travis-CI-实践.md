---
title: Travis 实践
date: 2021-01-09 14:22:54
categories: 编程学习
---

## 注册 Travis CI

登录 [Travis CI官网](https://travis-ci.com/) 可以使用github来登录.然后授权即可。然后选择监控github的仓库

## 生成 Github key 在 CI 注册  

登录github,在下面路径链接[生成key](https://github.com/settings/tokens)。注意复制保存，不然关闭后就无法查看，只能重新生成。

![image-20210109143119616](https://cdn.jsdelivr.net/gh/cemonliu/blogpic@main/image-20210109143119616.png)

将生成的 token 在 travis 里面对应的仓库里面去设定 key 值

![image-20210109143705570](https://cdn.jsdelivr.net/gh/cemonliu/blogpic@main/image-20210109143705570.png)  





## blog 目录下新建文件 .travis.yml. 具体设置如下

**主要修改的是分支名称**  

``` yaml
# 编译语言、环境
dist: xenial
os: linux
language: node_js

# Node.js 版本
node_js:
  - 12

# 只有 hexo 分支检出更改才触发 CI
branches:
  only:
    - backup

before_install:
  - export TZ='Asia/Shanghai' # 配置时区为东八区 UTC+8
  - npm install hexo-cli      # 安装 hexo

install:
  - npm install # 安装依赖

script: # 执行脚本，清除缓存，生成静态文件
  - hexo clean
  - hexo generate

deploy:
  strategy: git
  provider: pages
  skip_cleanup: true      # 跳过清理
  token: $GH_TOKEN        # GitHub Token 变量
  keep_history: true      # 保持推送记录，以增量提交的方式
  local_dir: public       # 需要推送到 GitHub 的静态文件目录
  target_branch: master   # 推送的目标分支 local_dir -> master 分支
  on:
    branch: backup # 工作分支
```

  <br/>

## 设置完成后，将文件 git到仓库，travis会自动运行

![image-20210109143547434](https://cdn.jsdelivr.net/gh/cemonliu/blogpic@main/image-20210109143547434.png)  

