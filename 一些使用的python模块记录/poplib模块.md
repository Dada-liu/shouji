# poplib模块



1. poplib.POP3()
2. server.set_debuglevel(1)
3. server.getwelcome()
4. server.user()
5. server.pass_()
6. server.stat()
7. server.top()
8. server.list()
9. server.retr(index)
10. server.quit() 关闭服务器
11. server.dele(index)   删除索引为index的邮件

# emile模块





# MIME

*MIME*(Multipurpose Internet Mail Extensions)多用途互联网邮件扩展类型。

#### MIME邮件的格式

参考资料：https://blog.csdn.net/robin_417556552/article/details/70853831

![image-20201217090406040](F:\大学\学习\第二学位本科学习\笔记\images\image-20201217090406040.png)



#### multipart类型的层次关系

![image-20201217094626548](F:\大学\学习\第二学位本科学习\笔记\images\image-20201217094626548.png)



#### MIME中的Content-Type

```
text/plain  text文件类型的默认设置
text/html
image/jpeg
image/png
audio/mpeg
audio/ogg
audio/*
video/mp4
application/*
application/json
application/javascript
application/ecmascript
application/octet-stream  二进制文件的默认设置
multipart/form-data  可用于HTML表单从浏览器发送信息给服务器。
multipart/byteranges 用于把部分的响应报文发送回浏览器。
text/html
text/css
text/javascript   你可能发现某些内容在 text/javascript 媒体类型末尾有一个 charset 参数，指定用于表                   示代码内容的字符集。这不是合法的，而且在大多数场景下会导致脚本不被载入。
```

邮件的各个部分叫做MIME段，每段前也缀以一个特别的头。

邮件的组成：头部分

​					   体部分



MIME信息头和MIME段头、



解析出来的

