
#include <QString>
#include <QDebug>

#include "Tools.h"

void ListChild(HWND parent, QString split){
    HWND child = FindWindowEx(parent, NULL, NULL, NULL);
    do{
        if( child != nullptr){
            wchar_t buf[512];
            memset(buf, 0, sizeof(buf));
            GetWindowText(child, buf, 512);
            std::wstring caption = std::wstring(buf);
            qDebug() << "Title->[" << std::wstring(buf) << "] " << split << child;
            if(caption.compare(L"DocView") == 0){
                qDebug() << "Get DocView handle->[" << child << "].";
            }

            QString splittmp = split + "-";
            ListChild(child, splittmp);
        }
        HWND tmp = FindWindowEx(parent, child, NULL, NULL);
        child = tmp;
    }while(child != nullptr);
}

HWND GetTargetChild(HWND parent, std::wstring caption){
    HWND child = FindWindowEx(parent, NULL, NULL, NULL);
    do{
        if( child != nullptr){
            wchar_t buf[512];
            memset(buf, 0, sizeof(buf));
            GetWindowText(child, buf, 512);
            std::wstring tmp = std::wstring(buf);
            // qDebug() << "Title->[" << std::wstring(buf) << "] " << child;
            if(tmp.compare(caption) == 0){
                qDebug() << "Get DocView handle->[" << child << "].";
                return child;
            }

            HWND hd = GetTargetChild(child, caption);
            if(hd != nullptr)
                return hd;
        }
        HWND tmp = FindWindowEx(parent, child, NULL, NULL);
        child = tmp;
    }while(child != nullptr);
    return nullptr;
}

void ShowWpsNavigation(HWND parent){
    // 遍历所有子窗口, 找出窗口名字为DocView的窗口

    HWND contentHwnd = GetTargetChild(parent, L"DocView");
    qDebug() << "Yes -> " << contentHwnd;

    // 左侧展示：alt+w,d,e
    // 右侧展示：alt+w,d,r
    // 隐藏：alt+w,d,v
    SetForegroundWindow(contentHwnd);
    PostMessage(contentHwnd, WM_SYSKEYDOWN, VK_MENU, 0);
    // Sleep(50);
    PostMessage(contentHwnd, WM_SYSKEYUP, VK_MENU, 0);
    PostMessage(contentHwnd, WM_KEYDOWN, 0x57, 0);
    PostMessage(contentHwnd, WM_KEYDOWN, 0x44, 0);
    PostMessage(contentHwnd, WM_KEYDOWN, 0x45, 0);
}

bool setWpsFloatButtonHide(){
    HWND hd = FindWindowEx(NULL, NULL, L"Qt5QWindowToolSaveBits", NULL);
    if(hd != 0){
        ShowWindow(hd, SW_MINIMIZE);
        qDebug() << "set Qt5QWindowToolSaveBits Hide.";
        return true;
    }

    return false;
}
