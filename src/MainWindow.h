
#pragma once

#include <QWindow>
#include <QAxWidget>

class MainWindow : public QWidget{
    Q_OBJECT
public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();
    bool initUI();
    void moveEvent(QMoveEvent *event) override;
    void resizeEvent(QResizeEvent *event) override;
    void showEvent(QShowEvent *event) override;

private:
    void SetWps();

private:
    QAxWidget* m_axWdiget = nullptr;
    QWidget* m_widgetContainer = nullptr;
    HWND m_hWps;
};
