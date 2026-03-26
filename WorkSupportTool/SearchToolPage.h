#pragma once
#include <windows.h>
#include <string>
#include <vector>

bool RegisterSearchToolPageClass(HINSTANCE hInstance);
HWND CreateSearchToolPage(HWND parent, HINSTANCE hInstance, const RECT& rc);
void SearchToolPage_SetVisible(HWND hwnd, bool visible);
void SearchToolPage_Resize(HWND hwnd, const RECT& rc);

// 検索結果のファイルパス一覧を取得
std::vector<std::wstring> SearchToolPage_GetResultPaths();
