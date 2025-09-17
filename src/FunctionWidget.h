
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
    void ShowCatalog();
    void StringReplace();
    void AddPicture();
    void FunctionInvoke();

private:
    void addConnection();

private:
    QPushButton * m_New;
    QPushButton * m_Open;
    QPushButton * m_ShowCatalog;
    QPushButton * m_StringReplace;
    QPushButton * m_AddPicture;
    QPushButton * m_FunctionInvoke;
};

