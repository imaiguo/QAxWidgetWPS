
# WPSHelper

Windows环境通过COM方式操控WPS [12.1.0.22529] 实现文档的一般自动化操作

## 配置开发运行环境

```bash
> cmd
> "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat"
> set Path=D:\devtools\Qt.6.9.1\bin;D:\devtools\openssl.3.0.8\bin;%Path%
> set OPENSSL_ROOT_DIR=D:/devtools/openssl.3.0.8/
>
```

## WPS SDK功能测试

- 文档自动打开保存
- 图片插入
- 字符插入
- 表格插入
- 页眉页脚插
- 设置字体名称、大小、颜色
- 字符查找
- 标题格式设置
- 获取文档所有字符
- 获取文档所有段落内容
- 获取段落中图片并保存到剪切板
- 另存为PDF格式

## 注册ocx控件

```bash
> regsvr32 ksaddndr.dll
> regsvr32 /u ksaddndr.dll
>
```

## 参考

- Active Qt https://doc.qt.io/qt-6/activeqt-index.html
- QAxBase Class https://doc.qt.io/qt-6/zh/qaxbase.html
- QAxWidget Class https://doc.qt.io/qt-6/qaxwidget.html

- WPS二次开发 https://gitee.com/zouyf/wps/tree/master
- C++ WPS二次开发资料汇总 https://gitcode.com/open-source-toolkit/54081 https://gitcode.com/open-source-toolkit/54081
- Qt/C++操作word文档 https://blog.csdn.net/weixin_42214392/article/details/141925713
- Qt窗口内嵌Word，PPT，Excel https://blog.csdn.net/a379039233/article/details/122838893
- 使用VS2008通过C++在WPS文档中添加内容的实践指南 https://blog.csdn.net/weixin_30798867/article/details/147543302
- PostMessage 向Windows窗口发送Alt组合键 https://blog.csdn.net/ribavnu/article/details/51437052
- Virtual-Key 代码 https://learn.microsoft.com/zh-cn/windows/win32/inputdev/virtual-key-codes
- WdBuiltinStyle枚举 https://qn.cache.wpscdn.cn/encs/doc/office_v11/index.htm?page=Java%20%E5%BA%94%E7%94%A8%E9%9B%86%E6%88%90%20WPS%20%E6%8C%87%E5%8D%97/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdBuiltinStyle%20%E6%9E%9A%E4%B8%BE.htm
