#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#include <windows.h>
#include <commctrl.h>
#include <objbase.h>
#include "SearchToolPage.h"
#include "PrintToolPage.h"

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "ole32.lib")

namespace {

HINSTANCE g_hInst = nullptr;
HWND g_hwndMain = nullptr;
HWND g_tabMain = nullptr;
HWND g_hwndSearchPage = nullptr;
HWND g_hwndPrintPage = nullptr;
int g_currentTab = 0;
HFONT g_hFontTab = nullptr;

RECT GetPageRect(HWND hwnd) {
    RECT rc{};
    GetClientRect(hwnd, &rc);

    const int padding = 10;
    const int tabH = 38;

    RECT page{
        padding,
        padding + tabH + 8,
        rc.right - padding,
        rc.bottom - padding
    };
    return page;
}

void LayoutMain(HWND hwnd) {
    RECT rc{};
    GetClientRect(hwnd, &rc);

    const int padding = 10;
    const int tabH = 36;

    MoveWindow(g_tabMain, padding, padding, max(300, rc.right - padding * 2), tabH, TRUE);

    RECT page = GetPageRect(hwnd);
    SearchToolPage_Resize(g_hwndSearchPage, page);
    PrintToolPage_Resize(g_hwndPrintPage, page);
}

void ApplyMainTab() {
    const bool isSearch = (g_currentTab == 0);
    if (!isSearch) {
        PrintToolPage_SetFiles(SearchToolPage_GetResultPaths());
    }
    SearchToolPage_SetVisible(g_hwndSearchPage, isSearch);
    PrintToolPage_SetVisible(g_hwndPrintPage, !isSearch);
}

LRESULT CALLBACK MainWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE:
    {
        g_hwndMain = hwnd;

        if (!g_hFontTab) {
            g_hFontTab = CreateFontW(
                -16, 0, 0, 0, FW_MEDIUM, FALSE, FALSE, FALSE,
                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                DEFAULT_PITCH | FF_DONTCARE, L"Meiryo UI");
        }
        HFONT hFont = g_hFontTab ? g_hFontTab : (HFONT)GetStockObject(DEFAULT_GUI_FONT);

        g_tabMain = CreateWindowExW(0, WC_TABCONTROLW, L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS,
            0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        SendMessageW(g_tabMain, WM_SETFONT, (WPARAM)hFont, TRUE);

        TCITEMW ti{};
        ti.mask = TCIF_TEXT;
        ti.pszText = const_cast<LPWSTR>(L"検索ツール");
        TabCtrl_InsertItem(g_tabMain, 0, &ti);
        ti.pszText = const_cast<LPWSTR>(L"印刷ツール");
        TabCtrl_InsertItem(g_tabMain, 1, &ti);
        TabCtrl_SetCurSel(g_tabMain, 0);
        //TabCtrl_SetItemSize(g_tabMain, 0, MAKELPARAM(110, 28));
        //TabCtrl_SetItemSize(g_tabMain, 0, MAKELPARAM(110, 28));

        RECT page = GetPageRect(hwnd);
        g_hwndSearchPage = CreateSearchToolPage(hwnd, g_hInst, page);
        g_hwndPrintPage = CreatePrintToolPage(hwnd, g_hInst, page);

        g_currentTab = 0;
        ApplyMainTab();
        LayoutMain(hwnd);
        return 0;
    }

    case WM_GETMINMAXINFO:
    {
        MINMAXINFO* p = reinterpret_cast<MINMAXINFO*>(lParam);
        if (p) {
            p->ptMinTrackSize.x = 1360; // 最小幅
            p->ptMinTrackSize.y = 980;  // 最小高さ
        }
        return 0;
    }

    case WM_SIZE:
        LayoutMain(hwnd);
        return 0;

    case WM_NOTIFY:
    {
        auto* hdr = reinterpret_cast<LPNMHDR>(lParam);
        if (hdr && hdr->hwndFrom == g_tabMain && hdr->code == TCN_SELCHANGE) {
            g_currentTab = TabCtrl_GetCurSel(g_tabMain);
            ApplyMainTab();
            LayoutMain(hwnd);
            return 0;
        }
        break;
    }

    case WM_DESTROY:
        if (g_hFontTab) { DeleteObject(g_hFontTab); g_hFontTab = nullptr; }
        PostQuitMessage(0);
        return 0;
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

} // namespace

int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR, int nCmdShow) {
    g_hInst = hInstance;

    const HRESULT hrInit = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    (void)hrInit;

    INITCOMMONCONTROLSEX icc{};
    icc.dwSize = sizeof(icc);
    icc.dwICC = ICC_WIN95_CLASSES | ICC_LISTVIEW_CLASSES | ICC_BAR_CLASSES | ICC_PROGRESS_CLASS | ICC_DATE_CLASSES | ICC_TAB_CLASSES;
    InitCommonControlsEx(&icc);

    RegisterSearchToolPageClass(hInstance);
    RegisterPrintToolPageClass(hInstance);

    WNDCLASSW wc{};
    wc.lpfnWndProc = MainWndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"ExcelFinderTabbedMainWin";
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    RegisterClassW(&wc);

    HWND hwnd = CreateWindowW(
        L"ExcelFinderTabbedMainWin",
        L"Excel検索 / 指定シート印刷",
        (WS_OVERLAPPEDWINDOW & ~WS_THICKFRAME & ~WS_MAXIMIZEBOX),
        120, 40, 1360, 980,
        nullptr, nullptr, hInstance, nullptr
    );

    if (!hwnd) {
        CoUninitialize();
        return 0;
    }

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    MSG msg{};
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    CoUninitialize();
    return 0;
}
