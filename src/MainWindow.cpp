#include "MainWindow.h"

#include <QAxObject>
#include <QMessageBox>

#include <Windows.h>

static const int Padding = 2;
static const int HeadPadding = 45;

MainWindow::MainWindow(QWidget *parent):QWidget(parent){

}

bool MainWindow::initUI(){
    setMinimumSize(600, 400);
    setWindowState(Qt::WindowMaximized);
    m_axWdiget = new QAxWidget("KWPS.Application", nullptr, Qt::WindowMaximizeButtonHint);
    m_axWdiget->setProperty("Visible", true);
    m_axWdiget->setProperty("DisplayAlerts", false); 

    // // 查找WPS主窗口
    m_hWps = FindWindow(nullptr, L"WPS Office");
    if(m_axWdiget->isNull() || m_hWps == 0){
        QMessageBox::critical(this, "错误", "系统未找到WPS,请安装WPS后再次尝试!", QMessageBox::Ok);
        return false;
    }
    // SetParent(m_hWps, (HWND)winId());

    QWindow* window = QWindow::fromWinId((WId)m_hWps);
    ShowWindow(m_hWps, SW_MAXIMIZE);


    m_widgetContainer = QWidget::createWindowContainer(window, this);
    return true;
}

MainWindow::~MainWindow(){
    if(m_axWdiget){
        // document->dynamicCall("Close()");
        m_axWdiget->dynamicCall("Quit()");
        // delete word;
    }
}

void MainWindow::moveEvent(QMoveEvent *ev){
    if(m_widgetContainer){
        ShowWindow(m_hWps, SW_MAXIMIZE);
        SetWps();
    }

    QWidget::moveEvent(ev);
}

void MainWindow::showEvent(QShowEvent *event){
    if(m_widgetContainer){
        ShowWindow(m_hWps, SW_MAXIMIZE);
        SetWps();
    }
}

void MainWindow::resizeEvent(QResizeEvent *ev){
    SetWps();
    QWidget::resizeEvent(ev);
}

void MainWindow::SetWps(){
    m_widgetContainer->setFixedSize(size().width() - Padding*2, size().height() + HeadPadding - Padding);
    m_widgetContainer->move(Padding, -HeadPadding);
}
