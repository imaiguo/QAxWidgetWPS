#include <QApplication>
#include <QDebug>
#include <comutil.h>
#include <QWidget>

#include <QAxWidget>
#include <QAxObject>

#pragma comment(lib,"comsupp.lib")

int main(int argc, char *argv[]) {
    QApplication a(argc, argv);

    QWidget w;
 	CoInitialize(NULL);
    // 创建Word应用程序对象
    QAxWidget* word = new QAxWidget("KWPS.Application", &w);
    // word->winId();
    // word->setParent(&w);
    word->setProperty("Visible", true); // 隐藏Word应用程序界面
 
    // 打开Word文档
    QAxObject* documents = word->querySubObject("Documents");
    QAxObject* document = documents->querySubObject("Open(const QString&)", "C:/Test/数据恢复测试.docx");
 
    if (!document) {
        qDebug() << "Failed to open the Word document!";
        return -1;
    }
 
    // 获取文档内容并修改
    QAxObject* selection = word->querySubObject("Selection");
    selection->dynamicCall("TypeText(const QString&)", "This is added text by Qt.");
 
    // 保存文档
    document->dynamicCall("SaveAs(const QString&)", "C:/Test/数据恢复测试.alter.docx");
 
    w.show();
    // 关闭文档和Word应用程序
    // document->dynamicCall("Close()");
    // word->dynamicCall("Quit()");
 
    // delete word;
 
    return a.exec();
}