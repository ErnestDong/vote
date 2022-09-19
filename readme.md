# 不可维护的烂代码

代码里面 python js 横飞，magic number 纵走，一出问题就 crash，正常运行则会打印一些日志直至结束。

## 运行前提

1. python。推荐 3.10，因为3.10中引入了新的类型注解写法，如为 python3.6-3.9，则需删掉代码中的 `: list[str]` 这种才能运行；如果为 3.6以下，则可能由于 `pathlib` 不兼容需要升级。
   1. python 安装时，切记要加入路径
   2. python 安装包 https://www.python.org/downloads/，安装适合自己系统的
2. selenium 驱动与 Google Chrome
   1. Firefox等浏览器也不是不行，但是已经硬编码为 Chrome 了
   2. selenium 驱动在这里：https://chromedriver.chromium.org/downloads
   3. selenium 驱动和 Chrome 版本要一致

## 运行方式

下载代码，并转到当前目录

``` shell
cd xxxx # 文件资源管理器/Finder显示的下载后代码的目录地址
```

把II投票.xlsx放到当前目录下，然后安装依赖

```shell
pip3 install -r requirements.txt
```

用记事本/vscode等修改 config.py：修改为 某某组、分析师(按顺序排，星多的优先)、代理(不会报错，但建议多换 IP，删掉末尾的 break 可以一直跑下去)，代理用尽/账号用尽/出bug会停下来

在 powershell(Windows)/termial(macOS)中运行 main.py，注意似乎注册有检测是否在台前，但我没找出来，所以建议就让 Chrome 在台前别动。代码编写者可以保证不含有危害本地计算机的代码，祝运行者好运！

``` shell
python3 main.py
```

运行结束可以根据打印出来的日志信息删除掉 Excel 里面的内容，以免重复跑多次。

## 联系方式

wxid-dongcy
