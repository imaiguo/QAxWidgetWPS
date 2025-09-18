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

    onNew();

    // 1. 获取当前文档选择范围/插入点
    m_Selection = m_axWdiget->querySubObject("Selection");

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
    connect(m_Funtion, &FunctionWidget::FunctionInvoke, this, &MainWindow::onFunctionInvoke);
}

void MainWindow::onNew(){
    qDebug() << "MainWindow::onNew";
    m_Documents->querySubObject("Close(const QString&)", "wpsDoNotSaveChanges");
    m_Document = m_Documents->querySubObject("Add");
    // qDebug() << "m_Documents->isNull:[" << m_Documents->isNull() << "].";
    // qDebug() << "m_Document->isNull:[" << m_Document->isNull() << "].";
}

void MainWindow::onOpen(){
    qDebug() << "MainWindow::onOpen";
    m_Documents->querySubObject("Close(const QString&)", "wpsDoNotSaveChanges");
    QString tmpPath = QCoreApplication::applicationDirPath();
    qDebug() << "applicationDirPath=" << tmpPath;
    tmpPath.append("/doc/测试文档.docx");
    m_Document = m_Documents->querySubObject("Open(const QString&)", tmpPath);
}

void MainWindow::onShowCatalog(){
    qDebug() << "MainWindow::onShowCatalog";
    QMessageBox::information(this, "提示", "onShowCatalog", QMessageBox::Ok);
    // With Selection.Find
    //     .Text = "Hello"
    //     .Replacement.Text = "Goodbye"
    //     .Execute Replace:=wpsReplaceAll

}

void MainWindow::onStringReplace(){
    qDebug() << "MainWindow::onStringReplace";
    QMessageBox::information(this, "提示", "onStringReplace", QMessageBox::Ok);
}

void MainWindow::onAddPicture(){
    qDebug() << "MainWindow::onAddPicture";
    QMessageBox::information(this, "提示", "onAddPicture", QMessageBox::Ok);
}

void MainWindow::onFunctionInvoke(){
    qDebug() << "MainWindow::onFunctionInvoke";
    
    // 1. 设置窗口状态
    // 0 wpsWindowStateNormal 正常
    // 1 wpsWindowStateMaximize 最大化 
    // 2 wpsWindowStateMinimize 最小化
    // m_axWdiget->setProperty("WindowState", 2);      // 最小化

    // 2. 在选定内容或插入点插入指定的文本
    m_Selection->dynamicCall("TypeText(const QString&)", "花间一壶酒，独酌无相亲。举杯邀明月，对影成三人。");
 
    // 3 插入段落 TypeParagraph 方法与 Enter（回车键）的功能相同。
    m_Selection->dynamicCall("TypeParagraph");
    QAxObject* para =  m_Selection->querySubObject("Paragraphs");
    if(para != nullptr){
        qDebug() << "Get Paragraphs.";
        qDebug() << para;
    }
    

    // 4.  文档中插入图片
    QString tmpPath = QCoreApplication::applicationDirPath();
    tmpPath.append("/image/setpasswd.png");
    qDebug() << "Image DirPath=" << tmpPath;
    QAxObject* nlineShapes =  m_Selection->querySubObject("InlineShapes");
    qDebug() << nlineShapes;
    // dynamicCall 和 querySubObject方法均可调用AddPicture
    const QVariant v1(false), v2(true), v3(0);
    // QVariant shape = nlineShapes->dynamicCall("AddPicture(const QString&, const QVariant &, const QVariant &)", tmpPath, v1, v2, v3);
    QAxObject* shape = nlineShapes->querySubObject("AddPicture(const QString&, const QVariant &, const QVariant &)", tmpPath, v1, v2, v3);
    qDebug() << shape;
    // 对象的 IDispatch 实现所暴露的属性可通过 Qt Object Model（Qt 对象模型）提供的属性系统进行读写（两个子类都是QObjects，因此可以使用QObject::setProperty() 和QObject::property() ）。
    // 不支持带有多个参数的属性。
    if( shape != nullptr){
        shape->setProperty("Width", 100);
        shape->setProperty("Height", 100);
        shape->dynamicCall("Select()");
    }
}

