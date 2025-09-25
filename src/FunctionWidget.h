
#pragma once

#include <QWidget>

#include <QPushButton>

class FunctionWidget: public QWidget{
    Q_OBJECT
public:
    FunctionWidget(QWidget * parent = nullptr);
    ~FunctionWidget();

    bool initUI();

signals:
    void New();
    void Open();

    void AddHeadFoot();
    void AddTable();
    void StringReplace();
    void StringInsert();
    void AddPicture();
    void AddHeadLine1();
    void AddHeadLine2();
    void FunctionInvoke();

private:
    void addConnection();

private:
    QPushButton * m_New;
    QPushButton * m_Open;
    QPushButton * m_AddTable;
    QPushButton * m_StringInsert;
    QPushButton * m_StringReplace;
    QPushButton * m_AddPicture;
    QPushButton * m_AddHeadFoot;
    QPushButton * m_AddHeadLine1;
    QPushButton * m_AddHeadLine2;
    QPushButton * m_FunctionInvoke;
};

