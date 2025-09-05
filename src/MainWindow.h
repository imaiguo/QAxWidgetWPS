
#pragma once

#include <QWindow>
#include <QAxWidget>

#include "FunctionWidget.h"

class MainWindow : public QWidget{
    Q_OBJECT
public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();
    bool initUI();
    void moveEvent(QMoveEvent *event) override;
    void resizeEvent(QResizeEvent *event) override;
    void showEvent(QShowEvent *event) override;
    void closeEvent(QCloseEvent *event) override;

private:
    void SetWps();
    void addConnection();

private:
    QAxWidget* m_axWdiget = nullptr;
    QAxObject* m_Documents = nullptr;
    QAxObject* m_Document = nullptr;
    QWidget* m_widgetContainer = nullptr;
    FunctionWidget *m_Funtion = nullptr;
    HWND m_hWps;

public slots:
    void onNew();
    void onOpen();
    void onShowCatalog();
    void onStringReplace();
    void onAddPicture();
};
