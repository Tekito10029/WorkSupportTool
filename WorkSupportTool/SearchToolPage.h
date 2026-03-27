#pragma once
#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#include <windows.h>
#include <string>
#include <vector>

bool RegisterSearchToolPageClass(HINSTANCE hInstance);
HWND CreateSearchToolPage(HWND parent, HINSTANCE hInstance, const RECT& rc);
void SearchToolPage_SetVisible(HWND hwnd, bool visible);
void SearchToolPage_Resize(HWND hwnd, const RECT& rc);
std::vector<std::wstring> SearchToolPage_GetResultPaths();
