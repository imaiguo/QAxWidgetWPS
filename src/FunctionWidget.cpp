
#include "FunctionWidget.h"

#include <QMessageBox>

FunctionWidget::FunctionWidget(QWidget * parent):QWidget(parent){

}

FunctionWidget::~FunctionWidget(){

}

bool FunctionWidget::initUI(){
    int height = 10;
    m_New = new QPushButton(this);
    m_New ->setText("新建");
    m_New->move(10, height);

    height += 30;
    m_Open = new QPushButton(this);
    m_Open ->setText("打开");
    m_Open->move(10, height);
 
    height += 30;
    m_ShowCatalog = new QPushButton(this);
    m_ShowCatalog ->setText("显示目录");
    m_ShowCatalog->move(10, height);

    height += 30;
    m_StringReplace = new QPushButton(this);
    m_StringReplace ->setText("字符替换");
    m_StringReplace->move(10, height);

    height += 30;
    m_AddPicture = new QPushButton(this);
    m_AddPicture ->setText("插入图片");
    m_AddPicture->move(10, height);

    height += 30;
    m_FunctionInvoke = new QPushButton(this);
    m_FunctionInvoke ->setText("功能调用");
    m_FunctionInvoke->move(10, height);

    addConnection();

    return true;
}

void FunctionWidget::addConnection(){
    connect(m_New, &QPushButton::clicked, this, [&]()-> void { emit New();});
    connect(m_Open, &QPushButton::clicked, this, [&]()-> void { emit Open();});
    connect(m_ShowCatalog, &QPushButton::clicked, this, [&]()-> void { emit ShowCatalog();});
    connect(m_StringReplace, &QPushButton::clicked, this, [&]()-> void { emit StringReplace();});
    connect(m_AddPicture, &QPushButton::clicked, this, [&]()-> void { emit AddPicture();});
    connect(m_FunctionInvoke, &QPushButton::clicked, this, [&]()-> void { emit FunctionInvoke();});
}
