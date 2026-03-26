#pragma once
#include <windows.h>
#include <string>
#include <vector>

bool RegisterPrintToolPageClass(HINSTANCE hInstance);
HWND CreatePrintToolPage(HWND parent, HINSTANCE hInstance, const RECT& rc);
void PrintToolPage_SetVisible(HWND hwnd, bool visible);
void PrintToolPage_Resize(HWND hwnd, const RECT& rc);

// 印刷対象ファイル一覧を置き換える
void PrintToolPage_SetFiles(const std::vector<std::wstring>& files);
