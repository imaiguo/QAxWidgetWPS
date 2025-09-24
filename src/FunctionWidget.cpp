
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
    m_StringInsert = new QPushButton(this);
    m_StringInsert ->setText("插入字符");
    m_StringInsert->move(10, height);

    height += 30;
    m_AddPicture = new QPushButton(this);
    m_AddPicture ->setText("插入图片");
    m_AddPicture->move(10, height);

    height += 30;
    m_AddTable = new QPushButton(this);
    m_AddTable ->setText("插入表格");
    m_AddTable->move(10, height);

    height += 30;
    m_AddHeadFoot = new QPushButton(this);
    m_AddHeadFoot ->setText("插入页眉页脚");
    m_AddHeadFoot->move(10, height);

    height += 30;
    m_StringReplace = new QPushButton(this);
    m_StringReplace ->setText("查找替换");
    m_StringReplace->move(10, height);

    height += 30;
    m_FunctionInvoke = new QPushButton(this);
    m_FunctionInvoke ->setText("另存为");
    m_FunctionInvoke->move(10, height);

    addConnection();

    return true;
}

void FunctionWidget::addConnection(){
    connect(m_New, &QPushButton::clicked, this, [&]()-> void { emit New();});
    connect(m_Open, &QPushButton::clicked, this, [&]()-> void { emit Open();});
    connect(m_AddTable, &QPushButton::clicked, this, [&]()-> void { emit AddTable();});
    connect(m_AddHeadFoot, &QPushButton::clicked, this, [&]()-> void { emit AddHeadFoot();});
    connect(m_StringInsert, &QPushButton::clicked, this, [&]()-> void { emit StringInsert();});
    connect(m_StringReplace, &QPushButton::clicked, this, [&]()-> void { emit StringReplace();});
    connect(m_AddPicture, &QPushButton::clicked, this, [&]()-> void { emit AddPicture();});
    connect(m_FunctionInvoke, &QPushButton::clicked, this, [&]()-> void { emit FunctionInvoke();});
}
