#include <QApplication>

#include <comutil.h>

#include <QAxWidget>
#include <QAxObject>
#include <QVBoxLayout>

#include "MainWindow.h"

#pragma comment(lib,"comsupp.lib")

int main(int argc, char *argv[]) {
    // CoInitialize(NULL);
    // CoInitializeEx(NULL, COINIT_MULTITHREADED);

    QApplication a(argc, argv);

    MainWindow window;
    window.setWindowTitle("实验报告助手");
    if(!window.initUI()){
        return -1;
    }

    window.show();

    return a.exec();
}
