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

constexpr UINT WM_APP_REFRESH_ACTIVE_PAGE = WM_APP + 100;

RECT GetPageRect() {
    RECT rc{};
    if (g_tabMain && IsWindow(g_tabMain)) {
        GetClientRect(g_tabMain, &rc);
        TabCtrl_AdjustRect(g_tabMain, FALSE, &rc);
        return rc;
    }

    if (g_hwndMain && IsWindow(g_hwndMain)) {
        GetClientRect(g_hwndMain, &rc);
        const int padding = 8;
        RECT page{
            padding,
            padding + 32,
            rc.right - padding,
            rc.bottom - padding
        };
        return page;
    }

    return rc;
}

void RefreshWindowTree(HWND hwnd) {
    if (!hwnd || !IsWindow(hwnd)) return;
    RedrawWindow(hwnd, nullptr, nullptr,
        RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN | RDW_UPDATENOW);
    UpdateWindow(hwnd);
}

void RefreshActivePageNow() {
    RECT page = GetPageRect();

    if (g_hwndSearchPage && IsWindow(g_hwndSearchPage)) {
        SearchToolPage_Resize(g_hwndSearchPage, page);
    }
    if (g_hwndPrintPage && IsWindow(g_hwndPrintPage)) {
        PrintToolPage_Resize(g_hwndPrintPage, page);
    }

    const bool showSearch = (g_currentTab == 0);

    if (g_hwndSearchPage && IsWindow(g_hwndSearchPage)) {
        ShowWindow(g_hwndSearchPage, showSearch ? SW_SHOW : SW_HIDE);
    }
    if (g_hwndPrintPage && IsWindow(g_hwndPrintPage)) {
        ShowWindow(g_hwndPrintPage, showSearch ? SW_HIDE : SW_SHOW);
    }

    HWND activePage = showSearch ? g_hwndSearchPage : g_hwndPrintPage;
    if (activePage && IsWindow(activePage)) {
        SetWindowPos(activePage, HWND_TOP,
                     page.left, page.top,
                     page.right - page.left,
                     page.bottom - page.top,
                     SWP_NOACTIVATE | SWP_SHOWWINDOW);
        RefreshWindowTree(activePage);
    }

    if (g_tabMain && IsWindow(g_tabMain)) {
        RefreshWindowTree(g_tabMain);
    }
    if (g_hwndMain && IsWindow(g_hwndMain)) {
        RefreshWindowTree(g_hwndMain);
    }
}

void LayoutMain(HWND hwnd) {
    RECT rc{};
    GetClientRect(hwnd, &rc);

    const int padding = 8;
    MoveWindow(g_tabMain, padding, padding,
               max(300, rc.right - padding * 2),
               max(120, rc.bottom - padding * 2),
               TRUE);

    PostMessageW(hwnd, WM_APP_REFRESH_ACTIVE_PAGE, 0, 0);
}

void ApplyMainTab() {
    if (g_tabMain && IsWindow(g_tabMain)) {
        TabCtrl_SetCurSel(g_tabMain, g_currentTab);
    }
    PostMessageW(g_hwndMain, WM_APP_REFRESH_ACTIVE_PAGE, 0, 0);
}

LRESULT CALLBACK MainWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE:
    {
        g_hwndMain = hwnd;

        HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

        g_tabMain = CreateWindowExW(0, WC_TABCONTROLW, L"",
            WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN,
            0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        SendMessageW(g_tabMain, WM_SETFONT, (WPARAM)hFont, TRUE);

        TCITEMW ti{};
        ti.mask = TCIF_TEXT;
        ti.pszText = const_cast<LPWSTR>(L"検索ツール");
        TabCtrl_InsertItem(g_tabMain, 0, &ti);
        ti.pszText = const_cast<LPWSTR>(L"印刷ツール");
        TabCtrl_InsertItem(g_tabMain, 1, &ti);
        TabCtrl_SetCurSel(g_tabMain, 0);

        RECT page = GetPageRect();
        g_hwndSearchPage = CreateSearchToolPage(g_tabMain, g_hInst, page);
        g_hwndPrintPage  = CreatePrintToolPage(g_tabMain, g_hInst, page);

        if (!g_hwndSearchPage || !g_hwndPrintPage) {
            MessageBoxW(hwnd, L"子ページの作成に失敗しました。", L"エラー", MB_OK | MB_ICONERROR);
        }

        g_currentTab = 0;
        LayoutMain(hwnd);
        ApplyMainTab();
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
            return 0;
        }
        break;
    }

    case WM_APP_REFRESH_ACTIVE_PAGE:
        RefreshActivePageNow();
        return 0;

    case WM_DESTROY:
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
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 1360, 980,
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
