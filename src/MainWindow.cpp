#include "MainWindow.h"

#include <QAxObject>
#include <QMessageBox>
#include <QCoreApplication>
#include <QKeyEvent>

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
    if(m_widgetContainer){
        m_widgetContainer->setFixedSize(size().width() - Padding*2 - ChildWidth, size().height() + HeadPadding - Padding);
        m_widgetContainer->move(Padding, -HeadPadding);
    }
}

void MainWindow::addConnection(){
    connect(m_Funtion, &FunctionWidget::New, this, &MainWindow::onNew);
    connect(m_Funtion, &FunctionWidget::Open, this, &MainWindow::onOpen);
    connect(m_Funtion, &FunctionWidget::ShowCatalog, this, &MainWindow::onShowCatalog);
    connect(m_Funtion, &FunctionWidget::StringReplace, this, &MainWindow::onTypeText);
    connect(m_Funtion, &FunctionWidget::AddPicture, this, &MainWindow::onAddPicture);
    connect(m_Funtion, &FunctionWidget::FunctionInvoke, this, &MainWindow::onFunctionInvoke);
}

void MainWindow::onNew(){
    qDebug() << "MainWindow::onNew";
    m_Documents->querySubObject("Close(const QString&)", "wpsDoNotSaveChanges");
    m_Document = m_Documents->querySubObject("Add");
}

void MainWindow::onOpen(){
    qDebug() << "MainWindow::onOpen";
    m_Documents->querySubObject("Close(const QString&)", "wpsDoNotSaveChanges");
    QString tmpPath = QCoreApplication::applicationDirPath();
    qDebug() << "applicationDirPath=" << tmpPath;
    tmpPath.append("/doc/测试文档.docx");
    m_Document = m_Documents->querySubObject("Open(const QString&)", tmpPath);
}

void ListChild(HWND parent, QString split){
    HWND child = FindWindowEx(parent, NULL, NULL, NULL);
    do{
        if( child != nullptr){
            qDebug() << split << child;
            QString splittmp = split + "-";
            ListChild(child, splittmp);
        }
        HWND tmp = FindWindowEx(parent, child, NULL, NULL);
        child = tmp;
    }while(child != nullptr);
}

void MainWindow::onShowCatalog(){
    qDebug() << "MainWindow::onShowCatalog";

    // 1. 遍历所有子窗口, 找出窗口名字为DocView的窗口
    qDebug() << m_hWps;
    ListChild(m_hWps, "-");



    // QWindow* widget = QWindow::fromWinId((WId)m_hWps)->findChild<QWindow*>("DocView");
    // qDebug() << widget;
    // qDebug() << widget->winId();

    // QMessageBox::information(this, "提示", "onShowCatalog", QMessageBox::Ok);
    // With Selection.Find
    //     .Text = "Hello"
    //     .Replacement.Text = "Goodbye"
    //     .Execute Replace:=wpsReplaceAll

    // 左侧展示：alt+w,d,e
    // 右侧展示：alt+w,d,r
    // 隐藏：alt+w,d,v

    // HWND id = (HWND)m_axWdiget->winId();
    // VK_A - VK_Z are the same as ASCII 'A' - 'Z' (0x41 - 0x5A)
    // qDebug() << m_hWps;

    // PostMessage(m_hWps, WM_SYSKEYDOWN, VK_MENU, 0);
    // // PostMessage(m_hWps,WM_SYSKEYDOWN,0x41,0);
    // Sleep(50);
    // // PostMessage(m_hWps,WM_SYSKEYUP,0x41,0);
    // PostMessage(m_hWps, WM_SYSKEYUP, VK_MENU, 0);

    // setforegroundwindow(hWnd)
    // ::SetForegroundWindow(m_hWps);
    // ::PostMessage(m_hWps, WM_KEYDOWN, MAKEWPARAM(VK_LMENU, 0), NULL);
    // ::PostMessage(m_hWps, WM_KEYUP, MAKEWPARAM(VK_LMENU, 0), NULL);

    // ::PostMessage(m_hWps, WM_KEYDOWN, MAKEWPARAM(VK_MENU, 0), NULL);
    // Sleep(50);
    // ::PostMessage(m_hWps, WM_KEYUP, MAKEWPARAM(VK_MENU, 0), NULL);


    // SendMessage(m_hWps, 0x0104, 0x00000012, 0x20380001);
    // SendMessage(m_hWps, 0x0105, 0x00000012, 0xC0380001);//(0x00000012 == VK_MENU(ALT键))



    // PostMessage(m_hWps, WM_KEYDOWN, VK_MENU, 0);
    // // PostMessage(m_hWps,WM_SYSKEYDOWN,0x41,0);
    // Sleep(50);
    // // PostMessage(m_hWps,WM_SYSKEYUP,0x41,0);
    // PostMessage(m_hWps, WM_KEYUP, VK_MENU, 0);


    // 当值为1时表示ALT键被按下！这不正是我需要的吗？于是把29位设置为1,函数调用变成
    // 经过测试，发现这个就是Alt+A的效果
    // PostMessage(hWnd,WM_SYSKEYDOWN,0x41,1<<29);


    // SendMessage(m_hWps, WM_SYSKEYDOWN, 0x44, 0);
    // Sleep(50);
    // SendMessage(m_hWps, WM_SYSKEYUP, 0x44, 0);
    // SendMessage(m_hWps, WM_SYSKEYDOWN, 0x45, 0);
    // Sleep(50);
    // SendMessage(m_hWps, WM_SYSKEYUP, 0x45, 0);

    // PostMessage(m_hWps,WM_SYSKEYDOWN,0x41,1<<29);

    // QKeyEvent keyPress(QEvent::KeyPress, Qt::Key_W, Qt::AltModifier);
    // QKeyEvent keyRelease(QEvent::KeyRelease, Qt::Key_W, Qt::AltModifier);

    // QWidget *w = QWidget::find((WId)m_hWps);
    // qDebug() <<w;
    // QCoreApplication::sendEvent(w, &keyPress);
    // QCoreApplication::sendEvent(w, &keyRelease);
}

void MainWindow::onTypeText(){
    qDebug() << "MainWindow::onStringReplace";
    m_Selection->dynamicCall("TypeParagraph");
    m_Selection->querySubObject("Paragraphs")->setProperty("Alignment", 1);
    QString txt = "花间一壶酒，独酌无相亲。举杯邀明月，对影成三人。";
    m_Selection->dynamicCall("TypeText(const QString&)", txt);
}

void MainWindow::onAddPicture(){
    qDebug() << "MainWindow::onAddPicture";
    m_Selection->dynamicCall("TypeParagraph");
    m_Selection->querySubObject("Paragraphs")->setProperty("Alignment", 1);

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
        int w = shape->property("Width").toInt();
        int h = shape->property("Height").toInt();
        qDebug() << w;
        qDebug() << h;
        shape->setProperty("Width", 332);
        shape->setProperty("Height", 263);
    }
}

void MainWindow::onFunctionInvoke(){
    qDebug() << "MainWindow::onFunctionInvoke";
    
    // 1. 设置窗口状态
    // wpsWindowStateNormal 正常 0
    // wpsWindowStateMaximize 最大化 1 
    // wpsWindowStateMinimize 最小化 2
    // m_axWdiget->setProperty("WindowState", 2);      // 最小化

    // 2. 在选定内容或插入点插入指定的文本
    m_Selection->querySubObject("Paragraphs")->setProperty("Alignment", 1);
    QString txt = "花间一壶酒，独酌无相亲。举杯邀明月，对影成三人。";
    m_Selection->dynamicCall("TypeText(const QString&)", txt);

    // 3 插入段落 TypeParagraph 方法与 Enter（回车键）的功能相同。
    m_Selection->dynamicCall("TypeParagraph");
    m_Selection->querySubObject("Paragraphs")->setProperty("Alignment", 1);
   
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
        int w = shape->property("Width").toInt();
        int h = shape->property("Height").toInt();
        qDebug() << w;
        qDebug() << h;
        shape->setProperty("Width", 400);
        shape->setProperty("Height", 400);
        shape->dynamicCall("Select()");
    }
}

