#include "MainWindow.h"

#include <QAxObject>

#include <Windows.h>

MainWindow::MainWindow(QWidget *parent):QWidget(parent){

}

void MainWindow::initUI(){
    setWindowState(Qt::WindowMaximized);
    m_axWdiget = new QAxWidget("KWPS.Application", nullptr, Qt::WindowMaximizeButtonHint);
    m_axWdiget->setProperty("Visible", true);
    m_axWdiget->setProperty("DisplayAlerts", false); 

    // // 查找WPS主窗口
    m_hWps = FindWindow(nullptr, L"WPS Office");
    QWindow* window = QWindow::fromWinId((WId)m_hWps);
    ShowWindow(m_hWps, SW_MAXIMIZE);

    if (m_hWps) {
        m_widgetContainer = QWidget::createWindowContainer(window, this);
    }
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
        m_widgetContainer->move(50, 50);
    }

    QWidget::moveEvent(ev);
}

void MainWindow::showEvent(QShowEvent *event){
    if(m_widgetContainer){
        ShowWindow(m_hWps, SW_MAXIMIZE);
        m_widgetContainer->move(50, 50);
    }
}

void MainWindow::resizeEvent(QResizeEvent *ev){
    QWidget::resizeEvent(ev);
}
