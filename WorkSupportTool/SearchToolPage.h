#pragma once
#include <windows.h>

bool RegisterSearchToolPageClass(HINSTANCE hInstance);
HWND CreateSearchToolPage(HWND parent, HINSTANCE hInstance, const RECT& rc);
void SearchToolPage_SetVisible(HWND hwnd, bool visible);
void SearchToolPage_Resize(HWND hwnd, const RECT& rc);
