
#include "MainWindow.h"
#include "Tools.h"

#include <QAxObject>
#include <QMessageBox>
#include <QCoreApplication>
#include <QKeyEvent>
#include <QFileDialog>

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

    m_axWdiget->setProperty("Visible", true);
    m_axWdiget->setProperty("DisplayAlerts", false); 

    ShowWindow(m_hWps, SW_MAXIMIZE);

    m_Funtion = new FunctionWidget(this);
    m_Funtion->initUI();
    m_Funtion->setFixedSize(ChildWidth, size().height() - Padding*2);
    m_Funtion->show();

    addConnection();
    m_Documents = m_axWdiget->querySubObject("Documents");

    onNew();

    // 自动显示 目录导航
    // ShowWpsNavigation(m_hWps);
    // Sleep(200);

    // 获取当前文档选择范围/插入点

    // QWindow* window = QWindow::fromWinId((WId)m_hWps);
    m_widgetContainer = QWidget::createWindowContainer(QWindow::fromWinId((WId)m_hWps), this);

    // // 全屏显示
    // QAxObject* window = m_axWdiget->querySubObject("ActiveWindow");
    // qDebug() << "window->" << window;
    // QAxObject* pane = window->querySubObject("ActivePane");
    // qDebug() << "pane->" << pane;
    // QAxObject* view = pane->querySubObject("View");     
    // qDebug() << "view->" << view;
    // view->setProperty("FullScreen", true);

    // setWpsFloatButtonHide();

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
            documents->dynamicCall("Close(int)", 0);
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
    connect(m_Funtion, &FunctionWidget::AddTable, this, &MainWindow::onAddTable);
    connect(m_Funtion, &FunctionWidget::StringInsert, this, &MainWindow::onTypeText);
    connect(m_Funtion, &FunctionWidget::StringReplace, this, &MainWindow::onStringReplace);
    connect(m_Funtion, &FunctionWidget::AddPicture, this, &MainWindow::onAddPicture);
    connect(m_Funtion, &FunctionWidget::FunctionInvoke, this, &MainWindow::onFunctionInvoke);
    connect(m_Funtion, &FunctionWidget::AddHeadLine1, this, &MainWindow::onAddHeadLine1);
    connect(m_Funtion, &FunctionWidget::AddHeadLine2, this, &MainWindow::onAddHeadLine2);
    connect(m_Funtion, &FunctionWidget::AddHeadFoot, this, &MainWindow::onAddHeadFoot);
}

void MainWindow::onNew(){
    qDebug() << "MainWindow::onNew";
    // 0关闭时不保存 1关闭时提示保存 2关闭时保存
    m_Documents->dynamicCall("Close(int)", 0);
    m_Document = m_Documents->querySubObject("Add");
    m_Selection = m_axWdiget->querySubObject("Selection");
}

void MainWindow::onOpen(){
    qDebug() << "MainWindow::onOpen";
    m_Documents->dynamicCall("Close(int)", 0);
    QString tmpPath = QCoreApplication::applicationDirPath();
    qDebug() << "applicationDirPath=" << tmpPath;
    tmpPath.append("/doc/测试文档.docx");
    m_Document = m_Documents->querySubObject("Open(const QString&)", tmpPath);
    m_Selection = m_axWdiget->querySubObject("Selection");
}

void MainWindow::onAddTable(){
    QAxObject* doc = m_axWdiget->querySubObject("ActiveDocument");
    qDebug() << doc;
    QAxObject* tables = doc->querySubObject("Tables");
    qDebug() << tables;
    m_Selection->dynamicCall("TypeParagraph");
    // Range 代表文档中的一个连续范围
    QAxObject* rangetable = m_Selection->querySubObject("Range");
    // 10行3列 defaultBehavior(1),AutoFitBehavior(2);//后一个2换成1,表格就变密集了
    // DefaultTableBehavior Variant 类型，可选。设置一个值，来指定 Wps Word 是否要根据单元格中的内容自动调整表格单元格的大小（“自动调整”功能）。
    // AutoFitBehavior Variant 类型，可选。该属性用于设置 Word 调整表格大小的“自动调整”规则。
    int row = 10, colum = 3;
    QStringList headers = {"序号", "姓名", "性别"};
    QStringList names = {"张三", "李四", "王五", "赵某", "钱某", "孙某", "李某", "周某", "吴某", "郑某"};
    QStringList sexs = {"男", "女"};
    if( rangetable != nullptr){
        QVariant v1(1), v2(2);
        tables->dynamicCall("Add(const QVariant &, int, int, const QVariant &, const QVariant &)", rangetable->asVariant(), row+1, colum, v1, v2);
    
        // 添加表头
        for( QString tmp: headers){
            // Range 代表文档中的一个连续范围
            QAxObject* range = m_Selection->querySubObject("Range");
            range->setProperty("Bold", true);

            // QAxObject* font = m_Selection->querySubObject("Font");
            // font->setProperty("Bold", true);

            m_Selection->dynamicCall("TypeText(const QString&)", tmp);
            m_Selection->querySubObject("Paragraphs")->setProperty("Alignment", 1);

            // 1表示单元格 1表示一个单位 0表示wpsMove 
            m_Selection->dynamicCall("MoveRight(int, int, int)", 1, 1, 0);
        }

        // 添加表主题内容
        for(int i = 0; i < row; i++){
            // 4表示表格行 1表示一个单位 0表示wpsMove 
            m_Selection->dynamicCall("MoveDown(int, int, int)", 4, 1, 0);
            for( int j=0; j < colum; j++){
                QString value = "";
                if( 0 == j)
                    value = std::to_string(i+1).c_str();
                if( 1 == j)
                    value = names.at(i); 
                if( 2 == j){
                    value = sexs.at(j%2);
                }
  
                m_Selection->dynamicCall("TypeText(const QString&)", value);
                m_Selection->dynamicCall("MoveRight(int, int, int)", 1, 1, 0);
            }
        }
        m_Selection->dynamicCall("MoveDown(int, int, int)", 4, 1, 0);
        m_Selection->dynamicCall("MoveDown(int, int, int)", 5, 1, 0);
    }
}

void MainWindow::onTypeText(){
    QAxObject* font = m_Selection->querySubObject("Font");
    qDebug() << font;
    font->setProperty("Size", 10);
    font->setProperty("Name", "幼圆");
    // font->setProperty("Name", "黑体");
    // BGR
    font->setProperty("Color", 0xFF0000);
    // 添加下划线
    font->setProperty("Underline", true);

    m_Selection->dynamicCall("TypeParagraph");
    // 获取选择区域 Range不适用段落 代表文档中的一个连续范围
    // QAxObject* range = m_Selection->querySubObject("Range");
    // range->setProperty("Bold", true);
    QStringList poem = {
        "李白《月下独酌·其一》",
        "花间一壶酒，独酌无相亲。",
        "举杯邀明月，对影成三人。",
        "月既不解饮，影徒随我身。",
        "暂伴月将影，行乐须及春。",
        "我歌月徘徊，我舞影零乱。",
        "醒时同交欢，醉后各分散。",
        "永结无情游，相期邈云汉。"
    };

    // QString txt = "从明天起，做一个幸福的人。喂马，劈柴，周游世界。从明天起，关心粮食和蔬菜。我有一所房子，面朝大海，春暖花开。";
    for(QString tmp: poem){
        m_Selection->dynamicCall("TypeText(const QString&)", tmp);
    }

    // ParagraphFormat
    QAxObject* paragraphformat = m_Selection->querySubObject("ParagraphFormat");
    // 0 靠左对齐 1剧中对齐 2 靠右对齐
    paragraphformat->setProperty("Alignment", 0);
    // 段落缩进2个字符
    paragraphformat->dynamicCall("IndentFirstLineCharWidth(int)", 2);

    // 字体还原
    font->setProperty("Color", 0x000000);
    font->setProperty("Underline", false);
}

void MainWindow::onAddPicture(){
    qDebug() << "MainWindow::onAddPicture";
    m_Selection->dynamicCall("TypeParagraph");
    // 居中对齐
    m_Selection->querySubObject("Paragraphs")->setProperty("Alignment", 1);

    QString tmpPath = QCoreApplication::applicationDirPath();
    tmpPath.append("/image/Gfp-wisconsin-madison-the-nature-boardwalk.jpg");
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

void MainWindow::onAddHeadFoot(){
    // 设置页眉页脚
    QAxObject* window = m_axWdiget->querySubObject("ActiveWindow");
    qDebug() << "window->" << window;
    QAxObject* pane = window->querySubObject("ActivePane");
    qDebug() << "pane->" << pane;
    QAxObject* view = pane->querySubObject("View");     
    qDebug() << "view->" << view;

    if(view == nullptr)
        return;

    // 进入设置页眉视图
    view->setProperty("SeekView", 9);
    // 居中对齐
    m_Selection->querySubObject("Paragraphs")->setProperty("Alignment", 1);
    QString txt = "研发中心编制";
    m_Selection->dynamicCall("TypeText(const QString&)", txt);
    // 进入正文编辑
    view->setProperty("SeekView", 0);

    // 进入设置页脚视图
    view->setProperty("SeekView", 10);
    // 居中对齐
    m_Selection->querySubObject("Paragraphs")->setProperty("Alignment", 1);
    m_Selection->dynamicCall("TypeText(const QString&)", "-");
    QAxObject* fields = m_Document->querySubObject("Fields");
    QAxObject* range = m_Selection->querySubObject("Range");
    fields->dynamicCall("Add(const QVariant &, int, const QString&, bool)", range->asVariant(), 33, "PAGE ", FALSE);
    m_Selection->dynamicCall("TypeText(const QString&)", "-");
    // 进入正文编辑
    view->setProperty("SeekView", 0);
}

void MainWindow::onStringReplace(){
    
    QVariant name = m_Document->property("FullName");
    qDebug() << name.toString();
    QAxObject* find = m_Selection->querySubObject("Find");
    if(find == nullptr)
        return;

    // // 进入设置页眉页脚视图
    // QAxObject* window = m_axWdiget->querySubObject("ActiveWindow");
    // qDebug() << "window->" << window;
    // QAxObject* pane = window->querySubObject("ActivePane");
    // qDebug() << "pane->" << pane;
    // QAxObject* view = pane->querySubObject("View");     
    // qDebug() << "view->" << view;
    // // 进入设置页眉视图
    // view->setProperty("SeekView", 9);
    // // 进入设置页脚视图
    // view->setProperty("SeekView", 10);

    // 方法1 循环查找
    find->setProperty("Text", "待替换文字");
    find->setProperty("MatchCase", false);
    find->setProperty("MatchWholeWord", false);
    // 如果查找操作向前搜索，则本属性为 True。如果向后搜索，则本属性为 False
    find->setProperty("Forward", true);

    QVariant result = find->dynamicCall("Execute()");
    if(result.toBool()){
        m_Selection->dynamicCall("TypeText(const QString&)", "大家好");
    }

    // // 进入正文编辑
    // view->setProperty("SeekView", 0);

    // // 方法2 14个参数
    // // Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward, Format, ReplaceWith, Replace, MatchKashida, MatchDiacritics, MatchAlefHamza, MatchControl)
    // QList<QVariant> vars;
    // QAxObject* replace = find->querySubObject("Replacement");
    // replace->setProperty("Text", "大家好");
    // vars.append(QVariant("待替换文字"));//FindText
    // vars.append(QVariant(false));//MatchCase
    // vars.append(QVariant(false));//MatchWholeWord
    // vars.append(QVariant(false));//MatchWildcards
    // vars.append(QVariant(false));//MatchSoundsLike
    // vars.append(QVariant(false));//MatchAllWordForms
    // vars.append(QVariant(true));//Forward
    // vars.append(QVariant(false));//Format
    // // vars.append(QVariant("大家好"));//ReplaceWith
    // vars.append(replace->asVariant());//ReplaceWith
    
    // vars.append(QVariant(3));//Replace 0不替换
    // vars.append(QVariant(false));//MatchKashida
    // vars.append(QVariant(false));//MatchDiacritics
    // vars.append(QVariant(false));//MatchAlefHamza
    // vars.append(QVariant(false));//MatchControl
    // QVariant result = find->dynamicCall("Execute(const QString&, bool, bool, bool, bool, bool, bool, bool, const QVariant &, int, bool, bool, bool, bool)", vars);
    // qDebug() << result.toBool();
}

void MainWindow::onAddHeadLine1(){
    // -1正文 -2Head1 -3Head2 ... -10Head9
    QAxObject* style = m_Selection->querySubObject("Style");
    style->setProperty("NextParagraphStyle", -2);
    m_Selection->dynamicCall("TypeParagraph");
    m_Selection->dynamicCall("TypeText(const QString&)", "壹级标题");
}

void MainWindow::onAddHeadLine2(){
    QAxObject* style = m_Selection->querySubObject("Style");
    style->setProperty("NextParagraphStyle", -3);
    m_Selection->dynamicCall("TypeParagraph");
    m_Selection->dynamicCall("TypeText(const QString&)", "贰级标题");
}

void MainWindow::onFunctionInvoke(){
    qDebug() << "MainWindow::onFunctionInvoke";
    
    // 1. 设置窗口状态
    // wpsWindowStateNormal 正常 0
    // wpsWindowStateMaximize 最大化 1 
    // wpsWindowStateMinimize 最小化 2
    // m_axWdiget->setProperty("WindowState", 2);      // 最小化

    // // 2. 在选定内容或插入点插入指定的文本
    // m_Selection->querySubObject("Paragraphs")->setProperty("Alignment", 1);
    // QString txt = "花间一壶酒，独酌无相亲。举杯邀明月，对影成三人。";
    // m_Selection->dynamicCall("TypeText(const QString&)", txt);

    // // 3 插入段落 TypeParagraph 方法与 Enter（回车键）的功能相同。
    // m_Selection->dynamicCall("TypeParagraph");
    // m_Selection->querySubObject("Paragraphs")->setProperty("Alignment", 1);
   
    // // 4.  文档中插入图片
    // QString tmpPath = QCoreApplication::applicationDirPath();
    // tmpPath.append("/image/setpasswd.png");
    // qDebug() << "Image DirPath=" << tmpPath;
    // QAxObject* nlineShapes =  m_Selection->querySubObject("InlineShapes");
    // qDebug() << nlineShapes;
    // // dynamicCall 和 querySubObject方法均可调用AddPicture
    // const QVariant v1(false), v2(true), v3(0);
    // // QVariant shape = nlineShapes->dynamicCall("AddPicture(const QString&, const QVariant &, const QVariant &)", tmpPath, v1, v2, v3);
    // QAxObject* shape = nlineShapes->querySubObject("AddPicture(const QString&, const QVariant &, const QVariant &)", tmpPath, v1, v2, v3);
    // qDebug() << shape;
    // // 对象的 IDispatch 实现所暴露的属性可通过 Qt Object Model（Qt 对象模型）提供的属性系统进行读写（两个子类都是QObjects，因此可以使用QObject::setProperty() 和QObject::property() ）。
    // // 不支持带有多个参数的属性。
    // if( shape != nullptr){
    //     int w = shape->property("Width").toInt();
    //     int h = shape->property("Height").toInt();
    //     qDebug() << w;
    //     qDebug() << h;
    //     shape->setProperty("Width", 400);
    //     shape->setProperty("Height", 400);
    //     shape->dynamicCall("Select()");
    // }

    // // 5. 设置段落字体
    // QAxObject* font = m_Selection->querySubObject("Font");
    // qDebug() << font;
    // font->setProperty("Size", 10);
    // font->setProperty("Name", "幼圆");
    // // font->setProperty("Name", "黑体");
    // // BGR
    // font->setProperty("Color", 0xFF0000);
    // // 添加下划线
    // font->setProperty("Underline", true);

    // m_Selection->dynamicCall("TypeParagraph");
    // // 获取选择区域 Range不适用段落 代表文档中的一个连续范围
    // // QAxObject* range = m_Selection->querySubObject("Range");
    // // range->setProperty("Bold", true);

    // QString txt = "从明天起，做一个幸福的人。喂马，劈柴，周游世界。从明天起，关心粮食和蔬菜。我有一所房子，面朝大海，春暖花开。";
    // m_Selection->dynamicCall("TypeText(const QString&)", txt);

    // // ParagraphFormat
    // QAxObject* paragraphformat = m_Selection->querySubObject("ParagraphFormat");
    // // 靠左对齐
    // paragraphformat->setProperty("Alignment", 0);
    // // 段落缩进2个字符
    // paragraphformat->dynamicCall("IndentFirstLineCharWidth(int)", 2);

    // // 字体还原
    // font->setProperty("Color", 0x000000);
    // font->setProperty("Underline", false);

    // // 6. 设置页眉页脚
    // QAxObject* window = m_axWdiget->querySubObject("ActiveWindow");
    // qDebug() << "window->" << window;
    // QAxObject* pane = window->querySubObject("ActivePane");
    // qDebug() << "pane->" << pane;
    // QAxObject* view = pane->querySubObject("View");     
    // qDebug() << "view->" << view;

    // // 进入设置页眉视图
    // view->setProperty("SeekView", 9);
    // // 居中对齐
    // m_Selection->querySubObject("Paragraphs")->setProperty("Alignment", 1);
    // QString txt = "研发中心编制";
    // m_Selection->dynamicCall("TypeText(const QString&)", txt);
    // // 进入正文编辑
    // view->setProperty("SeekView", 0);

    // // 进入设置页脚视图
    // view->setProperty("SeekView", 10);
    // // 居中对齐
    // m_Selection->querySubObject("Paragraphs")->setProperty("Alignment", 1);
    // m_Selection->dynamicCall("TypeText(const QString&)", "-");
    // QAxObject* fields = m_Document->querySubObject("Fields");
    // QAxObject* range = m_Selection->querySubObject("Range");
    // fields->dynamicCall("Add(const QVariant &, int, const QString&, bool)", range->asVariant(), 33, "PAGE ", FALSE);
    // m_Selection->dynamicCall("TypeText(const QString&)", "-");
    // // 进入正文编辑
    // view->setProperty("SeekView", 0);

    // 7. 文档另存为
    // QString selectDir = QFileDialog::getExistingDirectory();
    // if(selectDir.length() > 0){
    //     QString dateTimeString = QDateTime::currentDateTime().toString("yyyyMMdd.hhmmss");
    //     selectDir += "/实验笔记." + dateTimeString + ".docx";
    //     qDebug() << "Dir Saved Path:" << selectDir;
    //     QAxObject* doc = m_axWdiget->querySubObject("ActiveDocument");
    //     doc->dynamicCall("SaveAs(const QString&)", selectDir);
    // }

    // 8. 获取文档中所有文本
    // QAxObject* doc = m_axWdiget->querySubObject("ActiveDocument");
    // QAxObject* range = doc->querySubObject("Content");
    // QVariant text = range->property("Text");
    // qDebug() << text.toString();

    // 9. 获取文档中所有的段落
    // QAxObject* doc = m_axWdiget->querySubObject("ActiveDocument");
    // QAxObject* paras = doc->querySubObject("Paragraphs");
    // QVariant count =  paras->dynamicCall("Count()");
    // qDebug() << "总共: [" << count.toInt() << "]段。";

    // 9.1 获取段落文本
    // for(int i = 1; i< count.toInt() + 1; i++){
    //     // Item数组首个元素从1开始, 不是0
    //     QAxObject* para = paras->querySubObject("Item(int)", i);
    //     QAxObject* range = para->querySubObject("Range");
    //     QVariant text = range->property("Text");
    //     qDebug() << text.toString();
    //     qDebug() << "------------------------------------------------------------";
    // }

    // // 9.2 获取段落中的所有图片 并拷贝到剪切板
    // // WdInlineShapeType 枚举
    // // wdInlineShapePicture	3	图片。
    // for(int i = 1; i< count.toInt() + 1; i++){
    //     QAxObject* para = paras->querySubObject("Item(int)", i);
    //     QAxObject* range = para->querySubObject("Range");
    //     range->dynamicCall("Select()");
    //     QAxObject* shapes =  m_Selection->querySubObject("InlineShapes");
    //     QVariant c = shapes->property("Count");
    //     qDebug() << "Shapes count->[" << c.toInt() << "]";
    //     for( int j = 0; j < c.toInt(); j++){
    //         QAxObject* shape = shapes->querySubObject("Item(int)", j+1);
    //         QVariant type = shape->property("Type");
    //         // 3 图片
    //         qDebug() << "Shape type->[" << type.toInt() << "]";
    //         shape->dynamicCall("Select()");
    //         m_Selection->dynamicCall("Copy()");
    //     }
    //     qDebug() << "------------------------------------------------------------";
    // }

    // 10. 保存为pdf
    // WdSaveFormat 枚举
    // wdFormatPDF	17	PDF 格式。
    // wdFormatTemplate97	1	WPS 97 模板格式。
    // wdFormatXMLDocument	12	XML 文档格式。
    // wdFormatXMLDocumentMacroEnabled	13	启用了宏的 XML 文档格式。
    // wdFormatXMLTemplate	14	XML 模板格式。
    // wdFormatXMLTemplateMacroEnabled	15	启用了宏的 XML 模板格式。
    // wdFormatXPS	18	XPS 格式。
    QString selectDir = QFileDialog::getExistingDirectory();
    if(selectDir.length() > 0){
        QString dateTimeString = QDateTime::currentDateTime().toString("yyyyMMdd.hhmmss");
        selectDir += "/实验笔记." + dateTimeString + ".pdf";
        qDebug() << "Dir Saved Path:" << selectDir;
        QAxObject* doc = m_axWdiget->querySubObject("ActiveDocument");
        // QVariant v1(selectDir);
        QVariant name =  doc->property("FullName");
        qDebug() << name.toString();

        // 10.1 保存为DPF
        doc->dynamicCall("SaveAs(const QString&, int)", selectDir, 17);

        // // 10.2保存PDF带密码只读格式
        // QVariantList args;
        // args.append(QVariant(selectDir));//FileName
        // args.append(QVariant(17));//FileFormat
        // args.append(QVariant(false));//LockComments
        // args.append(QVariant("password"));//Password
        // args.append(QVariant(false));//AddToRecentFiles
        // args.append(QVariant("wpassword"));//WritePassword
        // args.append(QVariant(true));//ReadOnlyRecommended
        // doc->dynamicCall("SaveAs(const QString&, int, bool, const QString&, bool, const QString&, bool)", args);
    }
}
