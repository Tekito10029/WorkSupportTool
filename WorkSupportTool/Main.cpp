#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#include <windows.h>
#include <commctrl.h>
#include <objbase.h>
#include <uxtheme.h>
#include <iterator>
#include "SearchToolPage.h"
#include "PrintToolPage.h"

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "uxtheme.lib")

namespace {

HINSTANCE g_hInst = nullptr;
HWND g_hwndMain = nullptr;
HWND g_tabMain = nullptr;
HWND g_hwndSearchPage = nullptr;
HWND g_hwndPrintPage = nullptr;
int g_currentTab = 0;
HFONT g_hFontTab = nullptr;
HBRUSH g_hBrushMainBg = nullptr;
constexpr COLORREF kColorAppBg = RGB(245, 247, 250);
constexpr COLORREF kColorCardBg = RGB(255, 255, 255);
constexpr COLORREF kColorBorder = RGB(221, 227, 234);
constexpr COLORREF kColorText = RGB(31, 41, 55);
constexpr COLORREF kColorMutedText = RGB(107, 114, 128);
constexpr COLORREF kColorPrimary = RGB(37, 99, 235);

void ApplyModernControlTheme(HWND hwnd) {
    // 共通コントロールへ Explorer テーマを適用し、標準部品でも現代的な見た目に寄せる
    if (hwnd) {
        SetWindowTheme(hwnd, L"Explorer", nullptr);
    }
}

void FillRoundRect(HDC hdc, const RECT& rc, COLORREF fill, COLORREF border, int radius) {
    // タブ描画で使う角丸背景を塗りつぶし色と枠線色を分けて描画する
    HBRUSH brush = CreateSolidBrush(fill);
    HPEN pen = CreatePen(PS_SOLID, 1, border);
    HGDIOBJ oldBrush = SelectObject(hdc, brush);
    HGDIOBJ oldPen = SelectObject(hdc, pen);
    RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, radius, radius);
    SelectObject(hdc, oldPen);
    SelectObject(hdc, oldBrush);
    DeleteObject(pen);
    DeleteObject(brush);
}

bool DrawModernTab(const DRAWITEMSTRUCT* dis) {
    // オーナードローのタブを自前描画し、選択中タブには下線アクセントを付ける
    if (!dis || dis->CtlType != ODT_TAB || dis->hwndItem != g_tabMain) return false;

    wchar_t text[64]{};
    TCITEMW item{};
    item.mask = TCIF_TEXT;
    item.pszText = text;
    item.cchTextMax = (int)std::size(text);
    TabCtrl_GetItem(g_tabMain, (int)dis->itemID, &item);

    const bool selected = ((int)dis->itemID == TabCtrl_GetCurSel(g_tabMain));
    RECT rc = dis->rcItem;
    InflateRect(&rc, -3, -2);

    COLORREF fill = selected ? kColorCardBg : RGB(238, 242, 247);
    COLORREF border = selected ? kColorBorder : RGB(229, 234, 240);
    FillRoundRect(dis->hDC, rc, fill, border, 10);

    if (selected) {
        HBRUSH accent = CreateSolidBrush(kColorPrimary);
        RECT underline{ rc.left + 14, rc.bottom - 4, rc.right - 14, rc.bottom - 2 };
        FillRect(dis->hDC, &underline, accent);
        DeleteObject(accent);
    }

    HFONT font = reinterpret_cast<HFONT>(SendMessageW(g_tabMain, WM_GETFONT, 0, 0));
    HGDIOBJ oldFont = font ? SelectObject(dis->hDC, font) : nullptr;
    SetBkMode(dis->hDC, TRANSPARENT);
    SetTextColor(dis->hDC, selected ? kColorText : kColorMutedText);
    DrawTextW(dis->hDC, text, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
    if (oldFont) SelectObject(dis->hDC, oldFont);
    return true;
}

RECT GetPageRect(HWND hwnd) {
    RECT rc{};
    GetClientRect(hwnd, &rc);

    const int padding = 14;
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
    // メインウィンドウのサイズ変更に合わせ、タブと各ページを同じ表示領域へ再配置する
    RECT rc{};
    GetClientRect(hwnd, &rc);

    const int padding = 14;
    const int tabH = 34;

    MoveWindow(g_tabMain, padding, padding, max(300, rc.right - padding * 2), tabH, TRUE);

    RECT page = GetPageRect(hwnd);
    SearchToolPage_Resize(g_hwndSearchPage, page);
    PrintToolPage_Resize(g_hwndPrintPage, page);
}

void ApplyMainTab() {
    // 印刷タブへ移動する直前に検索結果のパス一覧を渡し、印刷対象を同期する
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
        if (!g_hBrushMainBg) {
            g_hBrushMainBg = CreateSolidBrush(kColorAppBg);
        }

        if (!g_hFontTab) {
            g_hFontTab = CreateFontW(
                -18, 0, 0, 0, FW_MEDIUM, FALSE, FALSE, FALSE,
                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                DEFAULT_PITCH | FF_DONTCARE, L"Yu Gothic UI");
        }
        HFONT hFont = g_hFontTab ? g_hFontTab : (HFONT)GetStockObject(DEFAULT_GUI_FONT);

        g_tabMain = CreateWindowExW(0, WC_TABCONTROLW, L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | TCS_FOCUSNEVER | TCS_OWNERDRAWFIXED,
            0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        ApplyModernControlTheme(g_tabMain);
        SendMessageW(g_tabMain, WM_SETFONT, (WPARAM)hFont, TRUE);
        SendMessageW(g_tabMain, TCM_SETITEMSIZE, 0, MAKELPARAM(150, 32));

        TCITEMW ti{};
        ti.mask = TCIF_TEXT;
        ti.pszText = const_cast<LPWSTR>(L"検索ツール");
        TabCtrl_InsertItem(g_tabMain, 0, &ti);
        ti.pszText = const_cast<LPWSTR>(L"印刷ツール");
        TabCtrl_InsertItem(g_tabMain, 1, &ti);
        TabCtrl_SetCurSel(g_tabMain, 0);

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

    case WM_ERASEBKGND:
    {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        RECT rc{};
        GetClientRect(hwnd, &rc);
        FillRect(hdc, &rc, g_hBrushMainBg ? g_hBrushMainBg : reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1));
        return 1;
    }

    case WM_SIZE:
        LayoutMain(hwnd);
        return 0;

    case WM_DRAWITEM:
        if (DrawModernTab(reinterpret_cast<DRAWITEMSTRUCT*>(lParam))) return TRUE;
        break;

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
        if (g_hBrushMainBg) { DeleteObject(g_hBrushMainBg); g_hBrushMainBg = nullptr; }
        PostQuitMessage(0);
        return 0;
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

} // 名前空間終了

int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR, int nCmdShow) {
    g_hInst = hInstance;

    // Excel/印刷設定ダイアログなど COM を使う処理に備え、UIスレッドをSTAで初期化する
    const HRESULT hrInit = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    const bool comInitialized = SUCCEEDED(hrInit);

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
    wc.hbrBackground = CreateSolidBrush(kColorAppBg);
    RegisterClassW(&wc);

    HWND hwnd = CreateWindowW(
        L"ExcelFinderTabbedMainWin",
        L"Excel検索 / 指定シート印刷",
        (WS_OVERLAPPEDWINDOW & ~WS_THICKFRAME & ~WS_MAXIMIZEBOX),
        120, 40, 1360, 980,
        nullptr, nullptr, hInstance, nullptr
    );

    if (!hwnd) {
        if (comInitialized) CoUninitialize();
        return 0;
    }

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    MSG msg{};
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    if (comInitialized) CoUninitialize();
    return 0;
}
