
#pragma once

#include "Windows.h"

void ListChild(HWND parent, QString split);

HWND GetTargetChild(HWND parent, std::wstring caption);

void ShowWpsNavigation(HWND parent);

bool setWpsFloatButtonHide();
