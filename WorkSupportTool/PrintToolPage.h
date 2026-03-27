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

bool RegisterPrintToolPageClass(HINSTANCE hInstance);
HWND CreatePrintToolPage(HWND parent, HINSTANCE hInstance, const RECT& rc);
void PrintToolPage_SetVisible(HWND hwnd, bool visible);
void PrintToolPage_Resize(HWND hwnd, const RECT& rc);
void PrintToolPage_SetFiles(const std::vector<std::wstring>& files);
