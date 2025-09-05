#include "MainWindow.h"

#include <QAxObject>
#include <QMessageBox>
#include <QCoreApplication>

#include <Windows.h>

static const int Padding = 2;
// static const int HeadPadding = 45;
static const int HeadPadding = 1;

static const int ChildWidth = 100;

MainWindow::MainWindow(QWidget *parent):QWidget(parent){

}

bool MainWindow::initUI(){
    setMinimumSize(600, 400);
    setWindowState(Qt::WindowMaximized);
    m_axWdiget = new QAxWidget("KWPS.Application", nullptr, Qt::WindowMaximizeButtonHint);

    // 查找WPS主窗口
    m_hWps = FindWindow(nullptr, L"WPS Office");
    if(m_axWdiget->isNull() || m_hWps == 0){
        QMessageBox::critical(this, "错误", "系统未找到WPS,请安装WPS后再次尝试!", QMessageBox::Ok);
        return false;
    }
    // SetParent(m_hWps, (HWND)winId());

    QWindow* window = QWindow::fromWinId((WId)m_hWps);
    m_axWdiget->setProperty("Visible", true);
    m_axWdiget->setProperty("DisplayAlerts", false); 

    ShowWindow(m_hWps, SW_MAXIMIZE);

    m_widgetContainer = QWidget::createWindowContainer(window, this);
    m_Funtion = new FunctionWidget(this);
    m_Funtion->initUI();
    m_Funtion->setFixedSize(ChildWidth, size().height() - Padding*2);
    m_Funtion->show();

    addConnection();
    m_Documents = m_axWdiget->querySubObject("Documents");

    return true;
}

MainWindow::~MainWindow(){
    if(m_axWdiget){
        m_axWdiget->dynamicCall("Quit()");
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
    m_Funtion->setFixedSize(ChildWidth, size().height() - Padding*2);
    m_Funtion->move(size().width() - ChildWidth - Padding, Padding);
    QWidget::resizeEvent(ev);
}

void MainWindow::closeEvent(QCloseEvent *event){
    if(m_axWdiget){
        ShowWindow(m_hWps, SW_HIDE);
        QAxObject* documents = m_axWdiget->querySubObject("Documents");
        if(!documents->isNull())
            documents->querySubObject("Close(const QString&)", "wpsDoNotSaveChanges");
        m_axWdiget->dynamicCall("Quit()");
    }

    QWidget::closeEvent(event);
}

void MainWindow::SetWps(){
    m_widgetContainer->setFixedSize(size().width() - Padding*2 - ChildWidth, size().height() + HeadPadding - Padding);
    m_widgetContainer->move(Padding, -HeadPadding);
}

void MainWindow::addConnection(){
    connect(m_Funtion, &FunctionWidget::New, this, &MainWindow::onNew);
    connect(m_Funtion, &FunctionWidget::Open, this, &MainWindow::onOpen);
    connect(m_Funtion, &FunctionWidget::ShowCatalog, this, &MainWindow::onShowCatalog);
    connect(m_Funtion, &FunctionWidget::StringReplace, this, &MainWindow::onStringReplace);
    connect(m_Funtion, &FunctionWidget::AddPicture, this, &MainWindow::onAddPicture);
}

void MainWindow::onNew(){
    m_Documents->querySubObject("Close(const QString&)", "wpsDoNotSaveChanges");
    m_Document = m_Documents->querySubObject("Add");
    qDebug() << "m_Documents->isNull:[" << m_Documents->isNull() << "].";
    qDebug() << "m_Document->isNull:[" << m_Document->isNull() << "].";
}

void MainWindow::onOpen(){
    m_Documents->querySubObject("Close(const QString&)", "wpsDoNotSaveChanges");
    QString tmpPath = QCoreApplication::applicationDirPath();
    qDebug() << "applicationDirPath=" << tmpPath;
    tmpPath.append("/doc/测试文档.docx");
    m_Document = m_Documents->querySubObject("Open(const QString&)", tmpPath);
}

void MainWindow::onShowCatalog(){
    QMessageBox::information(this, "提示", "onShowCatalog", QMessageBox::Ok);
    // With Selection.Find
    //     .Text = "Hello"
    //     .Replacement.Text = "Goodbye"
    //     .Execute Replace:=wpsReplaceAll

}

void MainWindow::onStringReplace(){
    QMessageBox::information(this, "提示", "onStringReplace", QMessageBox::Ok);
    // qDebug() << "onStringReplace called.";
    // QAxObject* sections =  m_Document->querySubObject("Sections");
    // qDebug() << sections->isNull();
    // QAxObject* section = sections->querySubObject("Count");
    // long c = section->asVariant().toLongLong();
    // qDebug () << section->isNull();
    // section->dynamicCall("Find(const QString&, const QString&, const QString&)", "测试", "Goodbye", "wpsReplaceAll");
}

void MainWindow::onAddPicture(){
    QMessageBox::information(this, "提示", "onAddPicture", QMessageBox::Ok);
}
