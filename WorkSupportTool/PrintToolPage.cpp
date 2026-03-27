#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#include "PrintToolPage.h"

#include <windows.h>
#include <commctrl.h>
#include <commdlg.h>
#include <objbase.h>
#include <oleauto.h>

#include <algorithm>
#include <sstream>
#include <string>
#include <vector>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")

namespace {

enum : int {
    IDC_STATIC_FILES = 5001,
    IDC_EDIT_FILES,
    IDC_BTN_ADD_FILES,
    IDC_BTN_CLEAR_FILES,

    IDC_STATIC_SHEETS,
    IDC_EDIT_SHEETS,
    IDC_STATIC_SHEETS_HINT,

    IDC_STATIC_COPIES,
    IDC_EDIT_COPIES,
    IDC_CHK_PREVIEW,

    IDC_BTN_PRINT,
    IDC_EDIT_LOG
};

HINSTANCE g_hInst = nullptr;
HWND g_hwndPage = nullptr;

HWND g_staticFiles = nullptr;
HWND g_editFiles = nullptr;
HWND g_btnAddFiles = nullptr;
HWND g_btnClearFiles = nullptr;

HWND g_staticSheets = nullptr;
HWND g_editSheets = nullptr;
HWND g_staticSheetsHint = nullptr;

HWND g_staticCopies = nullptr;
HWND g_editCopies = nullptr;
HWND g_chkPreview = nullptr;

HWND g_btnPrint = nullptr;
HWND g_editLog = nullptr;

std::vector<std::wstring> g_files;

std::wstring J(const wchar_t* s) { return std::wstring(s); }

void AddLog(const std::wstring& line) {
    if (!g_editLog) return;

    int len = GetWindowTextLengthW(g_editLog);
    SendMessageW(g_editLog, EM_SETSEL, (WPARAM)len, (LPARAM)len);

    std::wstring text = line;
    text += L"\r\n";
    SendMessageW(g_editLog, EM_REPLACESEL, FALSE, (LPARAM)text.c_str());
}

void RefreshFileList() {
    if (!g_editFiles) return;

    std::wstringstream ss;
    for (size_t i = 0; i < g_files.size(); ++i) {
        ss << g_files[i];
        if (i + 1 < g_files.size()) ss << L"\r\n";
    }
    SetWindowTextW(g_editFiles, ss.str().c_str());
}

std::wstring Trim(const std::wstring& s) {
    const wchar_t* ws = L" \t\r\n";
    const size_t a = s.find_first_not_of(ws);
    if (a == std::wstring::npos) return L"";
    const size_t b = s.find_last_not_of(ws);
    return s.substr(a, b - a + 1);
}

std::vector<std::wstring> SplitSheetNames(const std::wstring& s) {
    std::vector<std::wstring> out;
    std::wstring cur;

    auto flush = [&]() {
        std::wstring t = Trim(cur);
        if (!t.empty()) out.push_back(t);
        cur.clear();
    };

    for (wchar_t c : s) {
        if (c == L'\r') continue;
        if (c == L'\n' || c == L',' || c == L';' || c == 0x3001) {
            flush();
            continue;
        }
        cur.push_back(c);
    }
    flush();

    std::sort(out.begin(), out.end());
    out.erase(std::unique(out.begin(), out.end()), out.end());
    return out;
}

int GetCopies() {
    wchar_t buf[32]{};
    GetWindowTextW(g_editCopies, buf, 31);
    int copies = _wtoi(buf);
    if (copies <= 0) copies = 1;
    if (copies > 99) copies = 99;
    return copies;
}

bool GetCheck(HWND h) {
    return SendMessageW(h, BM_GETCHECK, 0, 0) == BST_CHECKED;
}

bool PickFiles(std::vector<std::wstring>& outFiles) {
    wchar_t buffer[65536]{};

    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = g_hwndPage;
    ofn.lpstrFilter =
        L"Excel Files (*.xlsx;*.xlsm;*.xls;*.xlsb)\0*.xlsx;*.xlsm;*.xls;*.xlsb\0"
        L"All Files (*.*)\0*.*\0\0";
    ofn.lpstrFile = buffer;
    ofn.nMaxFile = static_cast<DWORD>(std::size(buffer));
    ofn.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT | OFN_HIDEREADONLY;

    if (!GetOpenFileNameW(&ofn)) return false;

    const wchar_t* p = buffer;
    std::wstring first = p;
    p += first.size() + 1;

    if (*p == L'\0') {
        outFiles.push_back(first);
        return true;
    }

    std::wstring dir = first;
    while (*p) {
        std::wstring name = p;
        outFiles.push_back(dir + L"\\" + name);
        p += name.size() + 1;
    }
    return true;
}

HRESULT AutoWrap(int autoType, VARIANT* pvResult, IDispatch* pDisp, LPCOLESTR ptName, int cArgs, ...) {
    if (!pDisp) return E_INVALIDARG;

    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    DISPID dispid = 0;
    LPOLESTR nameBuf = const_cast<LPOLESTR>(ptName);
    HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, &nameBuf, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return hr;

    std::vector<VARIANT> args(static_cast<size_t>(cArgs));
    va_list marker;
    va_start(marker, cArgs);
    for (int i = 0; i < cArgs; ++i) {
        VARIANT* pArg = va_arg(marker, VARIANT*);
        args[static_cast<size_t>(cArgs - 1 - i)] = *pArg;
    }
    va_end(marker);

    dp.cArgs = cArgs;
    dp.rgvarg = cArgs ? args.data() : nullptr;

    DISPID dispidNamed = DISPID_PROPERTYPUT;
    if (autoType & DISPATCH_PROPERTYPUT) {
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
    }

    EXCEPINFO ex{};
    UINT argErr = (UINT)-1;
    return pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, autoType, &dp, pvResult, &ex, &argErr);
}

bool GetDispatchProperty(IDispatch* disp, LPCOLESTR name, IDispatch** out) {
    if (!disp || !out) return false;
    *out = nullptr;

    VARIANT v;
    VariantInit(&v);
    HRESULT hr = AutoWrap(DISPATCH_PROPERTYGET, &v, disp, name, 0);
    if (FAILED(hr) || v.vt != VT_DISPATCH || !v.pdispVal) {
        VariantClear(&v);
        return false;
    }
    *out = v.pdispVal;
    return true;
}

bool CallMethod(IDispatch* disp, LPCOLESTR name) {
    VARIANT v;
    VariantInit(&v);
    HRESULT hr = AutoWrap(DISPATCH_METHOD, &v, disp, name, 0);
    VariantClear(&v);
    return SUCCEEDED(hr);
}

bool CallMethod1(IDispatch* disp, LPCOLESTR name, VARIANT* a1) {
    VARIANT v;
    VariantInit(&v);
    HRESULT hr = AutoWrap(DISPATCH_METHOD, &v, disp, name, 1, a1);
    VariantClear(&v);
    return SUCCEEDED(hr);
}

bool CallMethodN(IDispatch* disp, LPCOLESTR name, int n, ...) {
    VARIANT v;
    VariantInit(&v);

    std::vector<VARIANT> copied;
    copied.reserve((size_t)n);

    va_list marker;
    va_start(marker, n);
    for (int i = 0; i < n; ++i) {
        VARIANT* pArg = va_arg(marker, VARIANT*);
        copied.push_back(*pArg);
    }
    va_end(marker);

    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    DISPID dispid = 0;
    LPOLESTR mutableName = const_cast<LPOLESTR>(name);
    HRESULT hr = disp->GetIDsOfNames(IID_NULL, &mutableName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    std::reverse(copied.begin(), copied.end());
    dp.cArgs = n;
    dp.rgvarg = n ? copied.data() : nullptr;

    EXCEPINFO ex{};
    UINT argErr = (UINT)-1;
    hr = disp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &v, &ex, &argErr);
    VariantClear(&v);
    return SUCCEEDED(hr);
}

IDispatch* GetWorksheetByName(IDispatch* worksheets, const std::wstring& sheetName) {
    if (!worksheets) return nullptr;

    VARIANT arg;
    VariantInit(&arg);
    arg.vt = VT_BSTR;
    arg.bstrVal = SysAllocString(sheetName.c_str());

    VARIANT result;
    VariantInit(&result);
    HRESULT hr = AutoWrap(DISPATCH_PROPERTYGET, &result, worksheets, L"Item", 1, &arg);

    VariantClear(&arg);

    if (FAILED(hr) || result.vt != VT_DISPATCH || !result.pdispVal) {
        VariantClear(&result);
        return nullptr;
    }
    return result.pdispVal;
}

bool SetBoolProperty(IDispatch* disp, LPCOLESTR name, bool value) {
    VARIANT arg;
    VariantInit(&arg);
    arg.vt = VT_BOOL;
    arg.boolVal = value ? VARIANT_TRUE : VARIANT_FALSE;

    VARIANT result;
    VariantInit(&result);
    HRESULT hr = AutoWrap(DISPATCH_PROPERTYPUT, &result, disp, name, 1, &arg);
    VariantClear(&result);
    return SUCCEEDED(hr);
}

bool OpenWorkbook(IDispatch* workbooks, const std::wstring& path, IDispatch** outBook) {
    if (!workbooks || !outBook) return false;
    *outBook = nullptr;

    VARIANT a1;
    VariantInit(&a1);
    a1.vt = VT_BSTR;
    a1.bstrVal = SysAllocString(path.c_str());

    VARIANT result;
    VariantInit(&result);
    HRESULT hr = AutoWrap(DISPATCH_METHOD, &result, workbooks, L"Open", 1, &a1);

    VariantClear(&a1);

    if (FAILED(hr) || result.vt != VT_DISPATCH || !result.pdispVal) {
        VariantClear(&result);
        return false;
    }

    *outBook = result.pdispVal;
    return true;
}

bool PrintWorksheet(IDispatch* ws, int copies, bool preview) {
    VARIANT vCopies, vPreview;
    VariantInit(&vCopies);
    VariantInit(&vPreview);

    vCopies.vt = VT_I4;
    vCopies.lVal = copies;

    vPreview.vt = VT_BOOL;
    vPreview.boolVal = preview ? VARIANT_TRUE : VARIANT_FALSE;

    return CallMethodN(ws, L"PrintOut", 2, &vCopies, &vPreview);
}

bool PrintOneBook(const std::wstring& filePath,
                  const std::vector<std::wstring>& sheetNames,
                  int copies,
                  bool preview,
                  std::wstring& outMessage) {
    outMessage.clear();

    CLSID clsid{};
    HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
    if (FAILED(hr)) {
        outMessage = L"Excel.Application could not be created.";
        return false;
    }

    IDispatch* app = nullptr;
    hr = CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&app);
    if (FAILED(hr) || !app) {
        outMessage = L"Failed to start Excel.";
        return false;
    }

    SetBoolProperty(app, L"Visible", false);
    SetBoolProperty(app, L"DisplayAlerts", false);

    IDispatch* workbooks = nullptr;
    if (!GetDispatchProperty(app, L"Workbooks", &workbooks)) {
        outMessage = L"Failed to get Workbooks.";
        app->Release();
        return false;
    }

    IDispatch* book = nullptr;
    if (!OpenWorkbook(workbooks, filePath, &book)) {
        outMessage = L"Failed to open workbook.";
        workbooks->Release();
        app->Release();
        return false;
    }

    IDispatch* worksheets = nullptr;
    bool okWorksheets = GetDispatchProperty(book, L"Worksheets", &worksheets);

    std::vector<std::wstring> printed;
    std::vector<std::wstring> missing;

    if (okWorksheets && worksheets) {
        for (const auto& target : sheetNames) {
            IDispatch* ws = GetWorksheetByName(worksheets, target);
            if (!ws) {
                missing.push_back(target);
                continue;
            }

            if (PrintWorksheet(ws, copies, preview)) printed.push_back(target);
            else missing.push_back(target + L"(print failed)");

            ws->Release();
        }
    }

    {
        std::wstringstream ss;
        if (!printed.empty()) {
            ss << L"Printed: " << filePath << L" / ";
            for (size_t i = 0; i < printed.size(); ++i) {
                if (i) ss << L", ";
                ss << printed[i];
            }
        } else {
            ss << L"No target sheets printed: " << filePath;
        }

        if (!missing.empty()) {
            ss << L" / Missing or failed: ";
            for (size_t i = 0; i < missing.size(); ++i) {
                if (i) ss << L", ";
                ss << missing[i];
            }
        }
        outMessage = ss.str();
    }

    if (worksheets) worksheets->Release();

    VARIANT saveChanges;
    VariantInit(&saveChanges);
    saveChanges.vt = VT_BOOL;
    saveChanges.boolVal = VARIANT_FALSE;
    CallMethod1(book, L"Close", &saveChanges);

    book->Release();
    workbooks->Release();

    CallMethod(app, L"Quit");
    app->Release();

    return !printed.empty();
}

void DoPrint() {
    if (g_files.empty()) {
        MessageBoxW(g_hwndPage, L"Please add Excel workbooks first.", L"Print Tool", MB_ICONWARNING);
        return;
    }

    wchar_t sheetBuf[4096]{};
    GetWindowTextW(g_editSheets, sheetBuf, 4095);
    std::vector<std::wstring> sheetNames = SplitSheetNames(sheetBuf);
    if (sheetNames.empty()) {
        MessageBoxW(g_hwndPage, L"Please enter at least one sheet name.", L"Print Tool", MB_ICONWARNING);
        return;
    }

    const int copies = GetCopies();
    const bool preview = GetCheck(g_chkPreview);

    AddLog(L"Start printing...");

    for (const auto& file : g_files) {
        std::wstring msg;
        PrintOneBook(file, sheetNames, copies, preview, msg);
        AddLog(msg);
        MSG m{};
        while (PeekMessageW(&m, nullptr, 0, 0, PM_REMOVE)) {
            TranslateMessage(&m);
            DispatchMessageW(&m);
        }
    }

    AddLog(L"Done.");
}

void LayoutPage(HWND hwnd) {
    RECT rc{};
    GetClientRect(hwnd, &rc);

    const int margin = 12;
    const int gap = 8;
    const int btnW = 120;
    const int labelW = 92;
    const int copiesW = 60;

    int x = margin;
    int y = margin;
    int w = rc.right - rc.left - margin * 2;

    MoveWindow(g_staticFiles, x, y + 4, labelW, 20, TRUE);
    MoveWindow(g_btnAddFiles, x + w - btnW * 2 - gap, y, btnW, 26, TRUE);
    MoveWindow(g_btnClearFiles, x + w - btnW, y, btnW, 26, TRUE);
    y += 30;

    MoveWindow(g_editFiles, x, y, w, 120, TRUE);
    y += 120 + gap;

    MoveWindow(g_staticSheets, x, y + 4, labelW, 20, TRUE);
    MoveWindow(g_editSheets, x + labelW, y, w - labelW, 52, TRUE);
    y += 56;
    MoveWindow(g_staticSheetsHint, x + labelW, y, w - labelW, 18, TRUE);
    y += 24;

    MoveWindow(g_staticCopies, x, y + 4, labelW, 20, TRUE);
    MoveWindow(g_editCopies, x + labelW, y, copiesW, 24, TRUE);
    MoveWindow(g_chkPreview, x + labelW + copiesW + 16, y, 120, 24, TRUE);
    MoveWindow(g_btnPrint, x + w - btnW, y - 1, btnW, 28, TRUE);
    y += 34;

    MoveWindow(g_editLog, x, y, w, rc.bottom - y - margin, TRUE);
}

LRESULT CALLBACK PrintToolPageWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE:
        g_hwndPage = hwnd;

        g_staticFiles = CreateWindowW(L"STATIC", L"Target workbooks",
            WS_CHILD | WS_VISIBLE,
            0, 0, 0, 0, hwnd, (HMENU)IDC_STATIC_FILES, g_hInst, nullptr);

        g_editFiles = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_LEFT | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY,
            0, 0, 0, 0, hwnd, (HMENU)IDC_EDIT_FILES, g_hInst, nullptr);

        g_btnAddFiles = CreateWindowW(L"BUTTON", L"Add files...",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_ADD_FILES, g_hInst, nullptr);

        g_btnClearFiles = CreateWindowW(L"BUTTON", L"Clear list",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_CLEAR_FILES, g_hInst, nullptr);

        g_staticSheets = CreateWindowW(L"STATIC", L"Sheet names",
            WS_CHILD | WS_VISIBLE,
            0, 0, 0, 0, hwnd, (HMENU)IDC_STATIC_SHEETS, g_hInst, nullptr);

        g_editSheets = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | ES_LEFT | ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL,
            0, 0, 0, 0, hwnd, (HMENU)IDC_EDIT_SHEETS, g_hInst, nullptr);

        g_staticSheetsHint = CreateWindowW(L"STATIC",
            L"Multiple names allowed: newline / , / ; separators",
            WS_CHILD | WS_VISIBLE,
            0, 0, 0, 0, hwnd, (HMENU)IDC_STATIC_SHEETS_HINT, g_hInst, nullptr);

        g_staticCopies = CreateWindowW(L"STATIC", L"Copies",
            WS_CHILD | WS_VISIBLE,
            0, 0, 0, 0, hwnd, (HMENU)IDC_STATIC_COPIES, g_hInst, nullptr);

        g_editCopies = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"1",
            WS_CHILD | WS_VISIBLE | ES_LEFT | ES_AUTOHSCROLL,
            0, 0, 0, 0, hwnd, (HMENU)IDC_EDIT_COPIES, g_hInst, nullptr);

        g_chkPreview = CreateWindowW(L"BUTTON", L"Preview",
            WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
            0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_PREVIEW, g_hInst, nullptr);

        g_btnPrint = CreateWindowW(L"BUTTON", L"Print",
            WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
            0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_PRINT, g_hInst, nullptr);

        g_editLog = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT",
            L"Only sheets that exist in the workbook are printed.\r\nMissing sheets are skipped.\r\n",
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_LEFT | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY,
            0, 0, 0, 0, hwnd, (HMENU)IDC_EDIT_LOG, g_hInst, nullptr);

        LayoutPage(hwnd);
        return 0;

    case WM_SIZE:
        LayoutPage(hwnd);
        return 0;

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDC_BTN_ADD_FILES: {
            std::vector<std::wstring> picked;
            if (PickFiles(picked)) {
                for (const auto& f : picked) {
                    if (std::find(g_files.begin(), g_files.end(), f) == g_files.end()) {
                        g_files.push_back(f);
                    }
                }
                RefreshFileList();
                AddLog(L"Files added.");
            }
            return 0;
        }

        case IDC_BTN_CLEAR_FILES:
            g_files.clear();
            RefreshFileList();
            AddLog(L"File list cleared.");
            return 0;

        case IDC_BTN_PRINT:
            DoPrint();
            return 0;
        }
        break;
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

} // namespace



void PrintToolPage_SetFiles(const std::vector<std::wstring>& files) {
    g_files = files;
    std::sort(g_files.begin(), g_files.end());
    g_files.erase(std::unique(g_files.begin(), g_files.end()), g_files.end());
    RefreshFileList();
    AddLog(L"検索結果を印刷対象一覧へ反映しました。");
}

bool RegisterPrintToolPageClass(HINSTANCE hInstance) {
    g_hInst = hInstance;

    WNDCLASSW wc{};
    wc.lpfnWndProc = PrintToolPageWndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"PrintToolPageWindow";
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);

    return RegisterClassW(&wc) != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS;
}

HWND CreatePrintToolPage(HWND parent, HINSTANCE hInstance, const RECT& rc) {
    return CreateWindowExW(
        0,
        L"PrintToolPageWindow",
        L"",
        WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN,
        rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top,
        parent,
        nullptr,
        hInstance,
        nullptr
    );
}


void PrintToolPage_SetVisible(HWND hwnd, bool visible) {
    if (!hwnd) return;
    ShowWindow(hwnd, visible ? SW_SHOW : SW_HIDE);
}

void PrintToolPage_Resize(HWND hwnd, const RECT& rc) {
    if (!hwnd) return;
    MoveWindow(hwnd, rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top, TRUE);
}
