#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#include "PrintToolPage.h"

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <commdlg.h>
#include <objbase.h>
#include <oleauto.h>
#include <winspool.h>

#include <algorithm>
#include <sstream>
#include <string>
#include <vector>

#include <ShlObj.h>
#include <filesystem>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "winspool.lib")

namespace {

    enum : int {
        IDC_STATIC_FILES = 5001,
        IDC_LIST_FILES,
        IDC_BTN_ADD_FILES,
        IDC_BTN_REMOVE_FILE,
        IDC_BTN_CLEAR_FILES,
        IDC_STATIC_REMOVE_TARGET,

        IDC_STATIC_SHEETS,
        IDC_EDIT_SHEETS,
        IDC_STATIC_SHEETS_HINT,
        IDC_BTN_SAVE_SHEETSET,

        IDC_STATIC_COPIES,
        IDC_EDIT_COPIES,
        IDC_CHK_PREVIEW,

        IDC_STATIC_PRINTER,
        IDC_CMB_PRINTER,
        IDC_BTN_PRINTER_PROP,

        IDC_STATIC_PAPER,
        IDC_CMB_PAPER,

        IDC_BTN_PRINT,
        IDC_EDIT_LOG
    };

    struct PrinterInfo {
        std::wstring name;
        std::wstring port;
    };

    HINSTANCE g_hInst = nullptr;
    HWND g_hwndPage = nullptr;

    HWND g_staticFiles = nullptr;
    HWND g_listFiles = nullptr;
    HWND g_btnAddFiles = nullptr;
    HWND g_btnRemoveFile = nullptr;
    HWND g_btnClearFiles = nullptr;
    HWND g_staticRemoveTarget = nullptr;

    HWND g_staticSheets = nullptr;
    HWND g_editSheets = nullptr;
    HWND g_staticSheetsHint = nullptr;
    HWND g_btnSaveSheetSet = nullptr;

    HWND g_staticCopies = nullptr;
    HWND g_editCopies = nullptr;
    HWND g_chkPreview = nullptr;

    HWND g_staticPrinter = nullptr;
    HWND g_cmbPrinter = nullptr;
    HWND g_btnPrinterProp = nullptr;

    HWND g_staticPaper = nullptr;
    HWND g_cmbPaper = nullptr;

    HWND g_btnPrint = nullptr;
    HWND g_editLog = nullptr;

    HWND g_log = nullptr;

    HFONT g_hFontUi = nullptr;
    HFONT g_hFontUiBold = nullptr;

    std::vector<std::wstring> g_files;
    std::vector<PrinterInfo> g_printers;
    std::wstring g_selectedPrinter;
    HGLOBAL g_hDevMode = nullptr;
    HGLOBAL g_hDevNames = nullptr;

    static std::wstring GetDefaultPrinterName()
    {
        DWORD needed = 0;
        GetDefaultPrinterW(nullptr, &needed);
        if (needed == 0) return L"";

        std::wstring name(needed, L'\0');
        if (!GetDefaultPrinterW(name.data(), &needed)) return L"";

        while (!name.empty() && name.back() == L'\0') name.pop_back();
        return name;
    }

    static bool SetDefaultPrinterName(const std::wstring& name)
    {
        if (name.empty()) return false;
        return SetDefaultPrinterW(name.c_str()) != FALSE;
    }

    std::wstring GetExeDir() {
        wchar_t buf[MAX_PATH * 4]{};
        DWORD n = GetModuleFileNameW(nullptr, buf, (DWORD)std::size(buf));
        std::wstring p(buf, n);
        size_t pos = p.find_last_of(L"\\/");
        return (pos == std::wstring::npos) ? L"." : p.substr(0, pos);
    }

    namespace fs = std::filesystem;
    std::wstring GetLocalAppDataPrintToolDir() {
        PWSTR p = nullptr;
        if (FAILED(SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, nullptr, &p)) || !p) {
            return L"";
        }

        std::wstring base(p);
        CoTaskMemFree(p);

        fs::path dir = fs::path(base) / L"ExcelToday";
        std::error_code ec;
        fs::create_directories(dir, ec);

        return dir.wstring();
    }

    std::wstring GetIniPath() {
        const std::wstring localDir = GetLocalAppDataPrintToolDir();
        if (!localDir.empty()) {
            return (fs::path(localDir) / L"WorkSupportTool_Print.ini").wstring();
        }
        return GetExeDir() + L"\\WorkSupportTool_Print.ini";
    }

    void AddLog(const std::wstring& line) {
        if (!g_editLog) return;
        int len = GetWindowTextLengthW(g_editLog);
        SendMessageW(g_editLog, EM_SETSEL, (WPARAM)len, (LPARAM)len);
        std::wstring text = line + L"\r\n";
        SendMessageW(g_editLog, EM_REPLACESEL, FALSE, (LPARAM)text.c_str());
    }

    void RefreshFileList() {
        if (!g_listFiles) return;
        SendMessageW(g_listFiles, LB_RESETCONTENT, 0, 0);
        for (const auto& f : g_files) {
            SendMessageW(g_listFiles, LB_ADDSTRING, 0, (LPARAM)f.c_str());
        }
    }


    int GetCurrentFileLineIndex() {
        if (!g_listFiles) return -1;
        int sel = (int)SendMessageW(g_listFiles, LB_GETCURSEL, 0, 0);
        return (sel == LB_ERR) ? -1 : sel;
    }


    void UpdateRemoveTargetLabel() {
        if (!g_staticRemoveTarget) return;

        int idx = GetCurrentFileLineIndex();
        if (idx < 0 || idx >= (int)g_files.size()) {
            SetWindowTextW(g_staticRemoveTarget, L"対象: なし");
            return;
        }

        std::wstring text = L"対象: " + g_files[(size_t)idx];
        SetWindowTextW(g_staticRemoveTarget, text.c_str());
    }

    void RemoveCurrentFileLine() {
        int idx = GetCurrentFileLineIndex();
        if (idx < 0 || idx >= (int)g_files.size()) {
            MessageBoxW(g_hwndPage, L"削除したい行にカーソルを置いてください。", L"印刷ツール", MB_OK | MB_ICONINFORMATION);
            return;
        }

        std::wstring removed = g_files[(size_t)idx];
        g_files.erase(g_files.begin() + idx);
        RefreshFileList();
        if (g_listFiles && !g_files.empty()) {
            int newSel = idx;
            if (newSel >= (int)g_files.size()) newSel = (int)g_files.size() - 1;
            SendMessageW(g_listFiles, LB_SETCURSEL, (WPARAM)newSel, 0);
        }
        UpdateRemoveTargetLabel();
        AddLog(L"削除: " + removed);
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
            if (c == L'\n' || c == L',' || c == L';' || c == 0x3001) flush();
            else cur.push_back(c);
        }
        flush();
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
        VARIANT result;
        VariantInit(&result);
        HRESULT hr = AutoWrap(DISPATCH_METHOD, &result, disp, name, 0);
        VariantClear(&result);
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

    bool SetStringProperty(IDispatch* disp, LPCOLESTR name, const std::wstring& value) {
        VARIANT arg;
        VariantInit(&arg);
        arg.vt = VT_BSTR;
        arg.bstrVal = SysAllocString(value.c_str());

        VARIANT result;
        VariantInit(&result);
        HRESULT hr = AutoWrap(DISPATCH_PROPERTYPUT, &result, disp, name, 1, &arg);
        VariantClear(&arg);
        VariantClear(&result);
        return SUCCEEDED(hr);
    }

    bool SetLongProperty(IDispatch* disp, LPCOLESTR name, long value) {
        VARIANT arg;
        VariantInit(&arg);
        arg.vt = VT_I4;
        arg.lVal = value;

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

    void FreeDevMode() {
        if (g_hDevMode) {
            GlobalFree(g_hDevMode);
            g_hDevMode = nullptr;
        }
    }

    bool CreateDefaultDevModeForPrinter(const std::wstring& printerName) {
        FreeDevMode();

        HANDLE hPrinter = nullptr;
        if (!OpenPrinterW((LPWSTR)printerName.c_str(), &hPrinter, nullptr) || !hPrinter) return false;

        LONG size = DocumentPropertiesW(g_hwndPage, hPrinter, (LPWSTR)printerName.c_str(), nullptr, nullptr, 0);
        if (size < (LONG)sizeof(DEVMODEW)) {
            ClosePrinter(hPrinter);
            return false;
        }

        HGLOBAL h = GlobalAlloc(GHND, (SIZE_T)size);
        if (!h) {
            ClosePrinter(hPrinter);
            return false;
        }

        auto* pDM = reinterpret_cast<DEVMODEW*>(GlobalLock(h));
        LONG ret = DocumentPropertiesW(g_hwndPage, hPrinter, (LPWSTR)printerName.c_str(), pDM, nullptr, DM_OUT_BUFFER);
        GlobalUnlock(h);
        ClosePrinter(hPrinter);

        if (ret != IDOK) {
            GlobalFree(h);
            return false;
        }

        g_hDevMode = h;
        return true;
    }

    void RefreshPaperCombo() {
        if (!g_cmbPaper) return;
        ComboBox_ResetContent(g_cmbPaper);

        if (!g_hDevMode) return;
        auto* pDM = reinterpret_cast<DEVMODEW*>(GlobalLock(g_hDevMode));
        if (!pDM) return;

        static const struct { short code; const wchar_t* name; } kPapers[] = {
            { DMPAPER_A4, L"A4" }, { DMPAPER_A3, L"A3" }, { DMPAPER_B5, L"B5" },
            { DMPAPER_B4, L"B4" }, { DMPAPER_LETTER, L"Letter" }, { DMPAPER_LEGAL, L"Legal" }
        };

        int sel = -1;
        short current = (pDM->dmFields & DM_PAPERSIZE) ? (short)pDM->dmPaperSize : (short)0;
        for (const auto& p : kPapers) {
            int idx = ComboBox_AddString(g_cmbPaper, p.name);
            ComboBox_SetItemData(g_cmbPaper, idx, p.code);
            if (p.code == current) sel = idx;
        }
        if (sel >= 0) ComboBox_SetCurSel(g_cmbPaper, sel);
        else if (ComboBox_GetCount(g_cmbPaper) > 0) ComboBox_SetCurSel(g_cmbPaper, 0);

        GlobalUnlock(g_hDevMode);
    }

    void RefreshPrinterCombo() {
        ComboBox_ResetContent(g_cmbPrinter);
        g_printers.clear();

        DWORD needed = 0, returned = 0;
        EnumPrintersW(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS, nullptr, 2, nullptr, 0, &needed, &returned);
        if (!needed) return;

        std::vector<BYTE> buf(needed);
        if (!EnumPrintersW(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS, nullptr, 2, buf.data(), needed, &needed, &returned)) {
            return;
        }

        auto* info = reinterpret_cast<PRINTER_INFO_2W*>(buf.data());
        int sel = -1;
        for (DWORD i = 0; i < returned; ++i) {
            PrinterInfo pi;
            pi.name = info[i].pPrinterName ? info[i].pPrinterName : L"";
            pi.port = info[i].pPortName ? info[i].pPortName : L"";
            int idx = ComboBox_AddString(g_cmbPrinter, pi.name.c_str());
            ComboBox_SetItemData(g_cmbPrinter, idx, (LPARAM)g_printers.size());
            g_printers.push_back(pi);
            if (!g_selectedPrinter.empty() && g_selectedPrinter == pi.name) sel = idx;
        }

        if (sel < 0 && ComboBox_GetCount(g_cmbPrinter) > 0) sel = 0;
        if (sel >= 0) {
            ComboBox_SetCurSel(g_cmbPrinter, sel);
            int data = (int)ComboBox_GetItemData(g_cmbPrinter, sel);
            if (data >= 0 && data < (int)g_printers.size()) {
                g_selectedPrinter = g_printers[data].name;
                CreateDefaultDevModeForPrinter(g_selectedPrinter);
                RefreshPaperCombo();
            }
        }
    }

    short GetSelectedPaperCode() {
        int sel = ComboBox_GetCurSel(g_cmbPaper);
        if (sel < 0) return 0;
        return (short)ComboBox_GetItemData(g_cmbPaper, sel);
    }

    void SaveSheetSettings() {
        const std::wstring ini = GetIniPath();
        wchar_t sheets[4096]{};
        GetWindowTextW(g_editSheets, sheets, 4095);
        wchar_t copies[32]{};
        GetWindowTextW(g_editCopies, copies, 31);

        std::wstring s = sheets;
        for (auto& ch : s) if (ch == L'\r') ch = L' ';
        for (auto& ch : s) if (ch == L'\n') ch = L'|';

        WritePrivateProfileStringW(L"Print", L"Sheets", s.c_str(), ini.c_str());
        WritePrivateProfileStringW(L"Print", L"Copies", copies, ini.c_str());
        WritePrivateProfileStringW(L"Print", L"Preview", GetCheck(g_chkPreview) ? L"1" : L"0", ini.c_str());
        WritePrivateProfileStringW(L"Print", L"Printer", g_selectedPrinter.c_str(), ini.c_str());
        WritePrivateProfileStringW(L"Print", L"PaperCode", std::to_wstring((int)GetSelectedPaperCode()).c_str(), ini.c_str());
        AddLog(L"印刷設定を保存しました。");
    }

    void LoadSheetSettings() {
        const std::wstring ini = GetIniPath();

        wchar_t bufSheets[4096]{};
        GetPrivateProfileStringW(L"Print", L"Sheets", L"", bufSheets, 4096, ini.c_str());
        std::wstring s = bufSheets;
        for (auto& ch : s) if (ch == L'|') ch = L'\n';
        SetWindowTextW(g_editSheets, s.c_str());

        wchar_t bufCopies[32]{};
        GetPrivateProfileStringW(L"Print", L"Copies", L"1", bufCopies, 32, ini.c_str());
        SetWindowTextW(g_editCopies, bufCopies);

        wchar_t bufPreview[8]{};
        GetPrivateProfileStringW(L"Print", L"Preview", L"0", bufPreview, 8, ini.c_str());
        SendMessageW(g_chkPreview, BM_SETCHECK, (_wtoi(bufPreview) != 0) ? BST_CHECKED : BST_UNCHECKED, 0);

        wchar_t bufPrinter[512]{};
        GetPrivateProfileStringW(L"Print", L"Printer", L"", bufPrinter, 512, ini.c_str());
        g_selectedPrinter = bufPrinter;
    }

    void ShowPrinterProperties() {
        if (g_selectedPrinter.empty()) return;
        if (!g_hDevMode && !CreateDefaultDevModeForPrinter(g_selectedPrinter)) {
            MessageBoxW(g_hwndPage, L"プリンタ設定を取得できませんでした。", L"印刷ツール", MB_OK | MB_ICONERROR);
            return;
        }

        HANDLE hPrinter = nullptr;
        if (!OpenPrinterW((LPWSTR)g_selectedPrinter.c_str(), &hPrinter, nullptr) || !hPrinter) {
            MessageBoxW(g_hwndPage, L"プリンタを開けませんでした。", L"印刷ツール", MB_OK | MB_ICONERROR);
            return;
        }

        auto* pDM = reinterpret_cast<DEVMODEW*>(GlobalLock(g_hDevMode));
        LONG res = DocumentPropertiesW(g_hwndPage, hPrinter, (LPWSTR)g_selectedPrinter.c_str(), pDM, pDM,
            DM_IN_BUFFER | DM_IN_PROMPT | DM_OUT_BUFFER);
        GlobalUnlock(g_hDevMode);
        ClosePrinter(hPrinter);

        if (res == IDOK) {
            RefreshPaperCombo();
            AddLog(L"プリンタの各種設定を更新しました。");
        }
    }

    std::wstring BuildActivePrinterString() {
        if (g_selectedPrinter.empty()) return L"";
        int sel = ComboBox_GetCurSel(g_cmbPrinter);
        if (sel >= 0) {
            int data = (int)ComboBox_GetItemData(g_cmbPrinter, sel);
            if (data >= 0 && data < (int)g_printers.size()) {
                return g_printers[data].name + L" on " + g_printers[data].port;
            }
        }
        return g_selectedPrinter;
    }

    bool ShowWorksheetPrintPreview(IDispatch* ws) {
        return CallMethod(ws, L"PrintPreview");
    }

    bool PrintWorksheet(IDispatch* ws, int copies, bool preview, const std::wstring& activePrinter) {
        if (preview) {
            return ShowWorksheetPrintPreview(ws);
        }

        VARIANT vFrom, vTo, vCopies, vPreview, vActivePrinter;
        VariantInit(&vFrom); VariantInit(&vTo); VariantInit(&vCopies); VariantInit(&vPreview); VariantInit(&vActivePrinter);

        vFrom.vt = VT_ERROR; vFrom.scode = DISP_E_PARAMNOTFOUND;
        vTo.vt = VT_ERROR; vTo.scode = DISP_E_PARAMNOTFOUND;

        vCopies.vt = VT_I4;
        vCopies.lVal = copies;

        vPreview.vt = VT_BOOL;
        vPreview.boolVal = VARIANT_FALSE;

        if (!activePrinter.empty()) {
            vActivePrinter.vt = VT_BSTR;
            vActivePrinter.bstrVal = SysAllocString(activePrinter.c_str());
        }
        else {
            vActivePrinter.vt = VT_ERROR;
            vActivePrinter.scode = DISP_E_PARAMNOTFOUND;
        }

        bool ok = CallMethodN(ws, L"PrintOut", 5, &vFrom, &vTo, &vCopies, &vPreview, &vActivePrinter);
        VariantClear(&vActivePrinter);
        return ok;
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
            outMessage = L"Excel.Application を作成できませんでした。";
            return false;
        }

        IDispatch* app = nullptr;
        hr = CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&app);
        if (FAILED(hr) || !app) {
            outMessage = L"Excel を起動できませんでした。";
            return false;
        }

        std::wstring selectedPrinter = g_selectedPrinter;
        std::wstring oldDefaultPrinter = GetDefaultPrinterName();
        bool changedDefaultPrinter = false;

        if (!selectedPrinter.empty() && oldDefaultPrinter != selectedPrinter) {
            changedDefaultPrinter = SetDefaultPrinterName(selectedPrinter);
        }

        SetBoolProperty(app, L"Visible", preview);
        SetBoolProperty(app, L"DisplayAlerts", false);

        std::wstring activePrinter = BuildActivePrinterString();
        if (!activePrinter.empty()) {
            SetStringProperty(app, L"ActivePrinter", activePrinter);
        }

        IDispatch* workbooks = nullptr;
        if (!GetDispatchProperty(app, L"Workbooks", &workbooks)) {
            outMessage = L"Workbooks の取得に失敗しました。";
            app->Release();
            return false;
        }

        IDispatch* book = nullptr;
        if (!OpenWorkbook(workbooks, filePath, &book)) {
            outMessage = L"ブックを開けませんでした。";
            workbooks->Release();
            app->Release();
            return false;
        }

        IDispatch* worksheets = nullptr;
        bool okWorksheets = GetDispatchProperty(book, L"Worksheets", &worksheets);

        std::vector<std::wstring> printed;
        std::vector<std::wstring> missing;
        short paper = GetSelectedPaperCode();

        if (okWorksheets && worksheets) {
            for (const auto& target : sheetNames) {
                IDispatch* ws = GetWorksheetByName(worksheets, target);
                if (!ws) {
                    missing.push_back(target);
                    continue;
                }

                if (paper != 0) {
                    IDispatch* pageSetup = nullptr;
                    if (GetDispatchProperty(ws, L"PageSetup", &pageSetup)) {
                        SetLongProperty(pageSetup, L"PaperSize", paper);
                        pageSetup->Release();
                    }
                }

                if (PrintWorksheet(ws, copies, preview, activePrinter)) printed.push_back(target);
                else missing.push_back(target + L"(印刷失敗)");

                ws->Release();
            }
        }

        std::wstringstream ss;
        if (!printed.empty()) {
            ss << L"印刷: " << filePath << L" / ";
            for (size_t i = 0; i < printed.size(); ++i) {
                if (i) ss << L", ";
                ss << printed[i];
            }
        }
        else {
            ss << L"対象シートを印刷できませんでした: " << filePath;
        }

        if (!missing.empty()) {
            ss << L" / 見つからない、または失敗: ";
            for (size_t i = 0; i < missing.size(); ++i) {
                if (i) ss << L", ";
                ss << missing[i];
            }
        }
        outMessage = ss.str();

        if (worksheets) worksheets->Release();

        if (book) {
            VARIANT saveChanges;
            VariantInit(&saveChanges);
            saveChanges.vt = VT_BOOL;
            saveChanges.boolVal = VARIANT_FALSE;

            VARIANT result;
            VariantInit(&result);
            AutoWrap(DISPATCH_METHOD, &result, book, L"Close", 1, &saveChanges);
            VariantClear(&result);
            book->Release();
        }

        workbooks->Release();
        CallMethod(app, L"Quit");
        app->Release();

        if (changedDefaultPrinter && !oldDefaultPrinter.empty()) {
            SetDefaultPrinterName(oldDefaultPrinter);
        }

        return !printed.empty();
    }

    
void ShowPrintSetupDialog() {
    PRINTDLGW pd{};
    pd.lStructSize = sizeof(pd);
    pd.hwndOwner = g_hwndPage;
    pd.Flags = PD_PRINTSETUP | PD_USEDEVMODECOPIESANDCOLLATE;

    if (g_hDevMode) pd.hDevMode = g_hDevMode;
    if (g_hDevNames) pd.hDevNames = g_hDevNames;

    if (PrintDlgW(&pd)) {
        if (g_hDevMode && g_hDevMode != pd.hDevMode) GlobalFree(g_hDevMode);
        if (g_hDevNames && g_hDevNames != pd.hDevNames) GlobalFree(g_hDevNames);
        g_hDevMode = pd.hDevMode;
        g_hDevNames = pd.hDevNames;
        AddLog(L"印刷設定を更新しました。");
    }
}

void DoPrint() {
        if (g_files.empty()) {
            MessageBoxW(g_hwndPage, L"先にExcelブックを追加してください。", L"印刷ツール", MB_ICONWARNING);
            return;
        }

        wchar_t sheetBuf[4096]{};
        GetWindowTextW(g_editSheets, sheetBuf, 4095);
        std::vector<std::wstring> sheetNames = SplitSheetNames(sheetBuf);
        if (sheetNames.empty()) {
            MessageBoxW(g_hwndPage, L"印刷するシート名を1つ以上入力してください。", L"印刷ツール", MB_ICONWARNING);
            return;
        }

        SaveSheetSettings();

        const int copies = GetCopies();
        const bool preview = GetCheck(g_chkPreview);

        if (preview) {
            AddLog(L"印刷プレビューを開きます。プレビュー画面で印刷またはキャンセルを選択してください。");
        } else {
            AddLog(L"印刷を開始します...");
        }

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

        AddLog(L"完了しました。");
    }

void LayoutPage(HWND hwnd) {
    RECT rc{};
    GetClientRect(hwnd, &rc);

    const int margin = 12;
    const int gap = 8;
    const int labelW = 100;
    const int copiesW = 68;
    const int rowH = 32;

    int x = margin;
    int y = margin;
    int w = rc.right - rc.left - margin * 2;

    const int minBtnW = 92;
    const int maxBtnW = 132;
    int freeWForTopButtons = w - labelW - 24;
    int btnW = max(minBtnW, min(maxBtnW, (freeWForTopButtons - gap * 2) / 3));

    int btnX3 = x + w - btnW;
    int btnX2 = btnX3 - gap - btnW;
    int btnX1 = btnX2 - gap - btnW;

    HDWP hdwp = BeginDeferWindowPos(20);
    if (!hdwp) return;

    auto D = [&](HWND h, int px, int py, int pw, int ph) {
        hdwp = DeferWindowPos(
            hdwp, h, nullptr, px, py, pw, ph,
            SWP_NOZORDER | SWP_NOACTIVATE
        );
        };

    D(g_staticFiles, x, y + 5, labelW, 22);
    D(g_btnAddFiles, btnX1, y, btnW, rowH);
    D(g_btnRemoveFile, btnX2, y, btnW, rowH);
    D(g_btnClearFiles, btnX3, y, btnW, rowH);
    y += rowH + 4;

    D(g_staticRemoveTarget, x, y, w, 20);
    y += 20 + 6;

    D(g_listFiles, x, y, w, 120);
    y += 120 + gap;

    D(g_staticSheets, x, y + 5, labelW, 22);
    D(g_editSheets, x + labelW, y, w - labelW - maxBtnW - gap, 52);
    D(g_btnSaveSheetSet, x + w - maxBtnW, y, maxBtnW, rowH);
    y += 56;

    D(g_staticSheetsHint, x + labelW, y, w - labelW, 20);
    y += 24;

    D(g_staticCopies, x, y + 5, labelW, 22);
    D(g_editCopies, x + labelW, y, copiesW, rowH);
    D(g_chkPreview, x + labelW + copiesW + gap, y + 4, 120, 24);
    D(g_btnPrinterProp, x + w - maxBtnW, y, maxBtnW, rowH);
    y += rowH + gap;

    D(g_staticPrinter, x, y + 5, labelW, 22);
    D(g_cmbPrinter, x + labelW, y, w - labelW - maxBtnW - gap, 220);
    y += rowH + gap;

    D(g_staticPaper, x, y + 5, labelW, 22);
    D(g_cmbPaper, x + labelW, y, 180, 220);
    y += rowH + gap;

    D(g_log, x, y + 5, labelW, 22);
    D(g_btnPrint, x + w - maxBtnW, y - 2, maxBtnW, rowH + 4);
    y += rowH + gap;

    D(g_editLog, x, y, w, max(80, rc.bottom - y - margin));

    EndDeferWindowPos(hdwp);

    RedrawWindow(
        hwnd, nullptr, nullptr,
        RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW
    );
}

    LRESULT CALLBACK PrintToolPageWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
        switch (msg) {
        case WM_CREATE:
        {
            g_hwndPage = hwnd;

            if (!g_hFontUi) {
                g_hFontUi = CreateFontW(
                    -18, 0, 0, 0, FW_MEDIUM, FALSE, FALSE, FALSE,
                    DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                    DEFAULT_PITCH | FF_DONTCARE, L"Meiryo UI");
            }
            if (!g_hFontUiBold) {
                g_hFontUiBold = CreateFontW(
                    -20, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
                    DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                    DEFAULT_PITCH | FF_DONTCARE, L"Meiryo UI");
            }

            HFONT hFont = g_hFontUi ? g_hFontUi : (HFONT)GetStockObject(DEFAULT_GUI_FONT);
            HFONT hFontBold = g_hFontUiBold ? g_hFontUiBold : hFont;

            g_staticFiles = CreateWindowW(L"STATIC", L"対象ブック",
                WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_STATIC_FILES, g_hInst, nullptr);
            g_listFiles = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
                WS_CHILD | WS_VISIBLE | WS_VSCROLL | LBS_NOTIFY | WS_BORDER,
                0, 0, 0, 0, hwnd, (HMENU)IDC_LIST_FILES, g_hInst, nullptr);
            g_btnAddFiles = CreateWindowW(L"BUTTON", L"ファイル追加...",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_ADD_FILES, g_hInst, nullptr);
            g_btnRemoveFile = CreateWindowW(L"BUTTON", L"項目削除",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_REMOVE_FILE, g_hInst, nullptr);
            g_btnClearFiles = CreateWindowW(L"BUTTON", L"一覧クリア",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_CLEAR_FILES, g_hInst, nullptr);
            g_staticRemoveTarget = CreateWindowW(L"STATIC", L"対象: なし",
                WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_STATIC_REMOVE_TARGET, g_hInst, nullptr);

            g_staticSheets = CreateWindowW(L"STATIC", L"印刷シート名",
                WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_STATIC_SHEETS, g_hInst, nullptr);
            g_editSheets = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                WS_CHILD | WS_VISIBLE | ES_LEFT | ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL,
                0, 0, 0, 0, hwnd, (HMENU)IDC_EDIT_SHEETS, g_hInst, nullptr);
            g_staticSheetsHint = CreateWindowW(L"STATIC", L"複数指定可: 改行 / , / ; 区切り",
                WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_STATIC_SHEETS_HINT, g_hInst, nullptr);
            g_btnSaveSheetSet = CreateWindowW(L"BUTTON", L"設定保存",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_SAVE_SHEETSET, g_hInst, nullptr);

            g_staticCopies = CreateWindowW(L"STATIC", L"部数",
                WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_STATIC_COPIES, g_hInst, nullptr);
            g_editCopies = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"1",
                WS_CHILD | WS_VISIBLE | ES_LEFT | ES_AUTOHSCROLL,
                0, 0, 0, 0, hwnd, (HMENU)IDC_EDIT_COPIES, g_hInst, nullptr);
            g_chkPreview = CreateWindowW(L"BUTTON", L"プレビュー",
                WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_PREVIEW, g_hInst, nullptr);

            g_staticPrinter = CreateWindowW(L"STATIC", L"プリンタ",
                WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_STATIC_PRINTER, g_hInst, nullptr);
            g_cmbPrinter = CreateWindowW(WC_COMBOBOXW, L"",
                WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL,
                0, 0, 0, 0, hwnd, (HMENU)IDC_CMB_PRINTER, g_hInst, nullptr);
            g_btnPrinterProp = CreateWindowW(L"BUTTON", L"各種設定...",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_PRINTER_PROP, g_hInst, nullptr);

            g_staticPaper = CreateWindowW(L"STATIC", L"用紙",
                WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_STATIC_PAPER, g_hInst, nullptr);
            g_cmbPaper = CreateWindowW(WC_COMBOBOXW, L"",
                WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL,
                0, 0, 0, 0, hwnd, (HMENU)IDC_CMB_PAPER, g_hInst, nullptr);

            g_btnPrint = CreateWindowW(L"BUTTON", L"▶ 印刷を実行",
                WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_PRINT, g_hInst, nullptr);

            g_log = CreateWindowW(L"STATIC", L"ログ",
                WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_STATIC_FILES, g_hInst, nullptr);
            g_editLog = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT",
                L"ブック内に存在するシートのみ印刷します。\r\n存在しないシートはスキップします。\r\n",
                WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_LEFT | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY,
                0, 0, 0, 0, hwnd, (HMENU)IDC_EDIT_LOG, g_hInst, nullptr);

            HWND controls[] = {
                g_staticFiles, g_listFiles, g_btnAddFiles, g_btnRemoveFile, g_btnClearFiles, g_staticRemoveTarget,
                g_staticSheets, g_editSheets, g_staticSheetsHint, g_btnSaveSheetSet,
                g_staticCopies, g_editCopies, g_chkPreview,
                g_staticPrinter, g_cmbPrinter, g_btnPrinterProp,
                g_staticPaper, g_cmbPaper, g_btnPrint, g_editLog,g_log
            };
            for (HWND h : controls) {
                SendMessageW(h, WM_SETFONT, (WPARAM)hFont, TRUE);
            }
            SendMessageW(g_btnPrint, WM_SETFONT, (WPARAM)hFontBold, TRUE);

            RefreshPrinterCombo();
            LoadSheetSettings();
            LayoutPage(hwnd);
            UpdateRemoveTargetLabel();
            return 0;
        }

        case WM_SIZE:
            LayoutPage(hwnd);
            return 0;

        case WM_COMMAND:
            switch (LOWORD(wParam)) {
            case IDC_BTN_ADD_FILES:
            {
                std::vector<std::wstring> picked;
                if (PickFiles(picked)) {
                    for (const auto& f : picked) {
                        if (std::find(g_files.begin(), g_files.end(), f) == g_files.end()) {
                            g_files.push_back(f);
                        }
                    }
                    RefreshFileList();
                    if (g_listFiles && !g_files.empty()) {
                        SendMessageW(g_listFiles, LB_SETCURSEL, (WPARAM)(g_files.size() - 1), 0);
                    }
                    UpdateRemoveTargetLabel();
                    AddLog(L"ファイルを追加しました。");
                }
                return 0;
            }

            case IDC_BTN_REMOVE_FILE:
                RemoveCurrentFileLine();
                return 0;

            case IDC_BTN_CLEAR_FILES:
                g_files.clear();
                RefreshFileList();
                UpdateRemoveTargetLabel();
                AddLog(L"一覧をクリアしました。");
                return 0;

            case IDC_BTN_SAVE_SHEETSET:
                SaveSheetSettings();
                return 0;

            case IDC_BTN_PRINTER_PROP:
                ShowPrinterProperties();
                return 0;

            case IDC_LIST_FILES:
                if (HIWORD(wParam) == LBN_SELCHANGE) {
                    UpdateRemoveTargetLabel();
                }
                return 0;

            case IDC_BTN_PRINT:
                DoPrint();
                return 0;

            case IDC_CMB_PRINTER:
                if (HIWORD(wParam) == CBN_SELCHANGE) {
                    int sel = ComboBox_GetCurSel(g_cmbPrinter);
                    if (sel >= 0) {
                        int data = (int)ComboBox_GetItemData(g_cmbPrinter, sel);
                        if (data >= 0 && data < (int)g_printers.size()) {
                            g_selectedPrinter = g_printers[data].name;
                            CreateDefaultDevModeForPrinter(g_selectedPrinter);
                            RefreshPaperCombo();
                        }
                    }
                }
                return 0;

            case IDC_CMB_PAPER:
                if (HIWORD(wParam) == CBN_SELCHANGE) {
                    // no-op, selected paper is read on print
                }
                return 0;
            }
            break;


        case WM_DESTROY:
            SaveSheetSettings();
            FreeDevMode();
            if (g_hDevNames) {
                GlobalFree(g_hDevNames);
                g_hDevNames = nullptr;
            }
            if (g_hFontUi) {
                DeleteObject(g_hFontUi);
                g_hFontUi = nullptr;
            }
            if (g_hFontUiBold) {
                DeleteObject(g_hFontUiBold);
                g_hFontUiBold = nullptr;
            }
            return 0;
        }

        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }
}

// namespace

void PrintToolPage_SetFiles(const std::vector<std::wstring>& files) {
    g_files = files;
    std::sort(g_files.begin(), g_files.end());
    g_files.erase(std::unique(g_files.begin(), g_files.end()), g_files.end());
    RefreshFileList();
    UpdateRemoveTargetLabel();
    AddLog(L"検索結果を印刷対象一覧へ反映しました。");
}

bool RegisterPrintToolPageClass(HINSTANCE hInstance) {
    g_hInst = hInstance;

    WNDCLASSW wc{};
    wc.lpfnWndProc = PrintToolPageWndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"PrintToolPageWindow";
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);

    return RegisterClassW(&wc) != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS;
}

HWND CreatePrintToolPage(HWND parent, HINSTANCE hInstance, const RECT& rc) {
    return CreateWindowExW(
        0,
        L"PrintToolPageWindow",
        L"",
        WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN,
        rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top,
        parent, nullptr, hInstance, nullptr
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
