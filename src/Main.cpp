#include <QApplication>

#include <comutil.h>

#include <QAxWidget>
#include <QAxObject>
#include <QVBoxLayout>

#include "MainWindow.h"

#pragma comment(lib,"comsupp.lib")

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
