#pragma once
#include <windows.h>

bool RegisterPrintToolPageClass(HINSTANCE hInstance);
HWND CreatePrintToolPage(HWND parent, HINSTANCE hInstance, const RECT& rc);
void PrintToolPage_SetVisible(HWND hwnd, bool visible);
void PrintToolPage_Resize(HWND hwnd, const RECT& rc);
