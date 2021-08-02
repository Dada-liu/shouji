# pyinstaller模块

打包发布.py文件的模块



1、带程序logo打包：pyinstaller -F -w -i favicon.ico JobHelperend.py



2、为exe文件添加版本信息：

> python set_version.py file_version_info.txt JobHelperend.exe



3、**怎样压缩pyinstaller打包的程序大小**

在.py程序中使用import语句时尽量不使用

​                `from 模块 import * `

而使用 

​               ` from 模块名 import 具体使用到的函数`



4、用pyinstaller获取.exe文件的版本信息，得到file_version_info.txt版本信息文件，但是此文件的格式不知道，后面有时间把这个文件的组成弄清，写一个文档；

在E:\python\Lib\site-packages\PyInstaller\utils\cliutils目录下有pyinstaller自带的工具