#include <QApplication>

#include <comutil.h>

#include <QAxWidget>
#include <QAxObject>
#include <QVBoxLayout>

#include "MainWindow.h"

#pragma comment(lib,"comsupp.lib")

// #import "C:\Users\ephraim\AppData\Local\kingsoft\WPS Office\12.1.0.22529\office6\ksaddndr.dll" no_namespace, raw_interfaces_only

// #import "C:\Users\ephraim\AppData\Local\kingsoft\WPS Office\12.1.0.22529\office6\wpscore.dll"

int main(int argc, char *argv[]) {
    QApplication a(argc, argv);

    a.setStyleSheet("QPushButton { border-width: 2px; border-color: black; background: beige; }");

    MainWindow window;
    window.setWindowTitle("WPS助手");
    if(!window.initUI()){
        return -1;
    }

    window.show();

    return a.exec();
}
