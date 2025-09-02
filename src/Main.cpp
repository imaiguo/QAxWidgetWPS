#include <QApplication>

#include <comutil.h>

#include <QAxWidget>
#include <QAxObject>
#include <QVBoxLayout>

#include "MainWindow.h"

#pragma comment(lib,"comsupp.lib")

int main(int argc, char *argv[]) {
    CoInitialize(NULL);

    QApplication a(argc, argv);

    MainWindow window;
    window.setWindowTitle("实验报告助手");
    if(!window.initUI()){
        return -1;
    }

    window.show();

    // QWidget window;
    // window.setFixedSize(1200, 900);

    // QVBoxLayout *layout = new QVBoxLayout(&window);

    // // 创建Word应用程序对象
    // QAxWidget* axWidget = new QAxWidget("KWPS.Application", &window);

    // // axWidget->setProperty("Visible", true); // 隐藏Word应用程序界面

    // layout->addWidget(axWidget);
    // axWidget->setProperty("Visible", true);
    // // 打开Word文档
    // QAxObject* documents = axWidget->querySubObject("Documents");
    // QAxObject* document = documents->querySubObject("Open(const QString&)", "C:/Test/数据恢复测试.docx");
 
    // if (!document) {
    //     qDebug() << "Failed to open the Word document!";
    //     return -1;
    // }
 
    // // 获取文档内容并修改
    // QAxObject* selection = axWidget->querySubObject("Selection");
    // selection->dynamicCall("TypeText(const QString&)", "This is added text by Qt.");
 
    // // 保存文档
    // document->dynamicCall("SaveAs(const QString&)", "C:/Test/数据恢复测试.alter.docx");
 
    // window.show();
    // // 关闭文档和Word应用程序
    // document->dynamicCall("Close()");
    // word->dynamicCall("Quit()");
 
    // delete word;
 
    return a.exec();
}