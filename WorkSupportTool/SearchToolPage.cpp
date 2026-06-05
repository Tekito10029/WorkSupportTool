//V_1.34
#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#include "SearchToolPage.h"
#include "PrintToolPage.h"
#include <windows.h>
#include <windowsx.h>   // マウス座標取得用マクロ
#include <commctrl.h>
#include <shobjidl.h>
#include <ShlObj.h>
#include <shellapi.h>
#include <commdlg.h>
#include <shlwapi.h>
#include <uxtheme.h>

#include <filesystem>
#include <string>
#include <vector>
#include <algorithm>
#include <iterator>
#include <chrono>
#include <atomic>
#include <memory>
#include <fstream>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "uxtheme.lib")

namespace fs = std::filesystem;


namespace Theme {
constexpr COLORREF AppBg = RGB(245, 247, 250);
constexpr COLORREF CardBg = RGB(255, 255, 255);
constexpr COLORREF Border = RGB(221, 227, 234);
constexpr COLORREF Text = RGB(31, 41, 55);
constexpr COLORREF MutedText = RGB(107, 114, 128);
constexpr COLORREF Primary = RGB(37, 99, 235);
constexpr COLORREF PrimaryHot = RGB(29, 78, 216);
constexpr COLORREF NeutralButton = RGB(255, 255, 255);
constexpr COLORREF NeutralButtonHot = RGB(243, 246, 251);
constexpr COLORREF NeutralButtonPressed = RGB(232, 238, 248);
constexpr COLORREF CsvButton = RGB(160, 255, 160);
constexpr COLORREF CsvButtonHot = RGB(96, 235, 96);
constexpr COLORREF CsvButtonPressed = RGB(72, 205, 72);
constexpr COLORREF CsvButtonBorder = RGB(118, 255, 118);
constexpr COLORREF Danger = RGB(220, 38, 38);
constexpr COLORREF DangerHot = RGB(185, 28, 28);
constexpr COLORREF DisabledBg = RGB(229, 231, 235);
constexpr COLORREF DisabledText = RGB(156, 163, 175);
constexpr COLORREF ListAltBg = RGB(248, 250, 252);
constexpr COLORREF SelectedBg = RGB(219, 234, 254);
constexpr COLORREF SelectedText = RGB(30, 64, 175);
constexpr COLORREF ProgressBg = RGB(229, 231, 235);
}


// -------------------- コントロールID --------------------
enum : int {
    IDC_BTN_BROWSE_ROOT = 102,

    // ルート一覧（複数フォルダーを直感的に扱う）
    IDC_LIST_ROOTS,
    IDC_BTN_ROOT_ADD,
    IDC_BTN_ROOT_REMOVE,
    IDC_BTN_ROOT_UP,
    IDC_BTN_ROOT_DOWN,
    IDC_BTN_ROOT_TOGGLE,
    IDC_STATIC_ROOTS_HINT,


    IDC_STATIC_MODE,
    IDC_CMB_MODE,
    IDC_STATIC_DAYS,
    IDC_EDIT_DAYS,

    // 日時基準
    IDC_STATIC_TIMEBASE,
    IDC_CMB_TIMEBASE,

    // 検索条件プリセット
    IDC_STATIC_PRESET,
    IDC_CMB_PRESET,
    IDC_BTN_PRESET_SAVE,
    IDC_BTN_PRESET_LOAD,
    IDC_BTN_PRESET_DELETE,

    // 拡張子
    IDC_GRP_EXT,
    IDC_CHK_XLS,
    IDC_CHK_XLSX,
    IDC_CHK_XLSM,
    IDC_CHK_XLSB,
    IDC_CHK_XLTX,
    IDC_CHK_XLTM,

    // フォルダー除外
    IDC_STATIC_EXCL_FOLDER,
    IDC_CHK_ENABLE_FOLDER_EXCL,
    IDC_LIST_EXCLUDES,
    IDC_BTN_ADD_EXCL_FOLDER,
    IDC_BTN_REMOVE_EXCL,
    IDC_BTN_EXCL_UP,
    IDC_BTN_EXCL_DOWN,
    IDC_BTN_LOAD_EXCL,
    IDC_BTN_SAVE_EXCL,

    IDC_STATIC_EXCL_PATTERN,
    IDC_EDIT_EXCL_PATTERN,
    IDC_BTN_ADD_PATTERN,
    IDC_BTN_UPDATE_PATTERN,

    // ファイル名除外
    IDC_STATIC_EXCL_NAME,
    IDC_CHK_ENABLE_NAME_EXCL,
    IDC_CHK_NAME_INCLUDE_EXT,
    IDC_EDIT_FNAME_PATTERN,
    IDC_BTN_ADD_FNAME,
    IDC_BTN_UPDATE_FNAME,
    IDC_BTN_REMOVE_FNAME,
    IDC_BTN_FNAME_UP,
    IDC_BTN_FNAME_DOWN,
    IDC_LIST_FNAME,

    // ファイル名除外の読み込み/保存
    IDC_BTN_LOAD_FNAME_EXCL,
    IDC_BTN_SAVE_FNAME_EXCL,

    // 操作
    IDC_BTN_SEARCH,
    IDC_BTN_STOP,
    IDC_BTN_EXPORT_CSV,

    // 表示
    IDC_PROGRESS,
    IDC_STATIC_PROGRESS,
    IDC_LIST_RESULTS,
    IDC_STATUS,

    // 結果フィルター
    IDC_STATIC_FILTER,
    IDC_EDIT_FILTER,
    IDC_STATIC_RESULT_DETAIL,

    // カレンダー範囲
    IDC_STATIC_FROM,
    IDC_DTP_FROM,
    IDC_STATIC_TO,
    IDC_DTP_TO,

    // 左ペインのタブ（検索 / 除外）
    IDC_TAB_LEFT,

};

// ポップアップメニュー用コマンド（コントロールIDではない）
enum : int {
    CMD_ROOT_ADD = 40001,
    CMD_ROOT_REMOVE,
    CMD_ROOT_UP,
    CMD_ROOT_DOWN,

    CMD_EXCL_ADD_FOLDER,
    CMD_EXCL_REMOVE,
    CMD_EXCL_UP,
    CMD_EXCL_DOWN,
    CMD_EXCL_LOAD,
    CMD_EXCL_SAVE,

    CMD_FNAME_ADD,
    CMD_FNAME_UPDATE,
    CMD_FNAME_REMOVE,
    CMD_FNAME_UP,
    CMD_FNAME_DOWN,
    CMD_FNAME_LOAD,
    CMD_FNAME_SAVE,

    CMD_RESULT_ADD_TO_PRINT = 41001,
};

static const UINT WM_APP_ADD_HIT = WM_APP + 1;
static const UINT WM_APP_PROGRESS = WM_APP + 2;
static const UINT WM_APP_FINISHED = WM_APP + 3;
static const UINT WM_APP_THREADERR = WM_APP + 4;
static const UINT WM_APP_TOTAL = WM_APP + 5;

static const UINT WM_APP_SCANPATH = WM_APP + 6; // 現在走査中のフォルダーを表示する

static void CloseModernCalendarPopup();

// -------------------- データモデル --------------------
struct Hit {
    std::wstring timeText;        // 表示用（更新/作成どちらでも）
    unsigned long long sizeKB = 0;
    std::wstring fileName;
    std::wstring path;
    // フィルターと並べ替えで毎回小文字化しないよう、検索ヒット登録時にキャッシュする
    std::wstring fileNameLow;
    std::wstring pathLow;
};

enum class ExcludeType { DirPrefix, Wildcard, Substring };
struct ExcludeRule {
    ExcludeType type{};
    std::wstring raw;
    fs::path dirNorm;
    std::wstring pattern;
    // 除外判定は探索中に大量実行されるため、小文字化済みの値を保持して再計算を避ける
    std::wstring rawLow;
    std::wstring dirNormLow;
};

enum class TimeBase { LastWrite = 0, Creation = 1, Either = 2 };
struct SearchParams {
    std::vector<fs::path> roots;
    bool useFolderExcl = false;
    bool useNameExcl = false;
    std::chrono::system_clock::time_point rangeStart{};
    std::chrono::system_clock::time_point rangeEnd{};
    TimeBase timeBase = TimeBase::LastWrite;
};

// -------------------- グローバル変数 --------------------
static HINSTANCE g_hInst = nullptr;
static HWND g_hwndMain = nullptr;

static HFONT g_hFontUi = nullptr;
static HFONT g_hFontUiBold = nullptr;
static HFONT g_hFontTabLeft = nullptr;
static HBRUSH g_hBrushAppBg = nullptr;
static HBRUSH g_hBrushCardBg = nullptr;
static HBRUSH g_hBrushEditBg = nullptr;
static HIMAGELIST g_hResultsRowImageList = nullptr;

static HWND g_tabLeft = nullptr;
static int  g_leftTab = 0; // 0: 検索, 1: 除外

static HWND g_staticRoot = nullptr;
static HWND g_staticMode = nullptr;
static HWND g_staticDays = nullptr;
static HWND g_staticTimeBase = nullptr;
static HWND g_staticExclFolder = nullptr;
static HWND g_staticExclPattern = nullptr;
static HWND g_staticExclName = nullptr;

static HWND g_btnBrowseRoot = nullptr; // 追加ボタンとして再利用
static HWND g_listRoots = nullptr;
static HWND g_btnRootRemove = nullptr;
static HWND g_btnRootUp = nullptr;
static HWND g_btnRootDown = nullptr;
static HWND g_btnRootToggle = nullptr;
static HWND g_staticRootsHint = nullptr;

static HWND g_cmbMode = nullptr;      // 0: 今日, 1: 過去N日
static HWND g_editDays = nullptr;

static HWND g_cmbTimeBase = nullptr;  // 0: 更新日時, 1: 作成日時, 2: 更新OR作成

static HWND g_staticPreset = nullptr;
static HWND g_cmbPreset = nullptr;
static HWND g_btnPresetSave = nullptr;
static HWND g_btnPresetLoad = nullptr;
static HWND g_btnPresetDelete = nullptr;

static HWND g_chkXls = nullptr;
static HWND g_chkXlsx = nullptr;
static HWND g_chkXlsm = nullptr;
static HWND g_chkXlsb = nullptr;
static HWND g_chkXltx = nullptr;
static HWND g_chkXltm = nullptr;

// フォルダー除外
static HWND g_chkEnableFolderExcl = nullptr;
static HWND g_listExcludes = nullptr;
static HWND g_btnAddExclFolder = nullptr;
static HWND g_btnRemoveExcl = nullptr;
static HWND g_btnExclUp = nullptr;
static HWND g_btnExclDown = nullptr;
static HWND g_btnLoadExcl = nullptr;
static HWND g_btnSaveExcl = nullptr;

static HWND g_editExclPattern = nullptr;
static HWND g_btnAddPattern = nullptr;
// 更新ボタンは廃止（編集→Enter/フォーカスアウトで即反映）
static WNDPROC g_oldExclEditProc = nullptr;

// 名前除外
static HWND g_chkEnableNameExcl = nullptr;
static HWND g_chkNameIncludeExt = nullptr;
static HWND g_editFNamePattern = nullptr;
static HWND g_btnAddFName = nullptr;
// 更新ボタンは廃止（編集→Enter/フォーカスアウトで即反映）
static WNDPROC g_oldFNameEditProc = nullptr;
static HWND g_btnRemoveFName = nullptr;
static HWND g_btnFNameUp = nullptr;
static HWND g_btnFNameDown = nullptr;
static HWND g_listFName = nullptr;
static HWND g_btnLoadFNameExcl = nullptr;
static HWND g_btnSaveFNameExcl = nullptr;

// 操作/表示
static HWND g_btnSearch = nullptr;
static HWND g_btnStop = nullptr;
static HWND g_btnExportCsv = nullptr;
static HWND g_progress = nullptr;
static HWND g_staticProgress = nullptr;
static HWND g_listResults = nullptr;
static HWND g_status = nullptr;

// 結果フィルター
static HWND g_staticFilter = nullptr;
static HWND g_editFilter = nullptr;
static HWND g_staticResultDetail = nullptr;
static int g_visibleCount = 0;
static bool g_deferExcludeListRefresh = false;
static bool g_deferFileNameListRefresh = false;

static HWND g_staticFrom = nullptr;
static HWND g_staticTo = nullptr;
static HWND g_dtpFrom = nullptr;
static HWND g_dtpTo = nullptr;
static SYSTEMTIME g_dateFrom{};
static SYSTEMTIME g_dateTo{};
static HWND g_calendarPopup = nullptr;
static HWND g_calendarTarget = nullptr;
static SYSTEMTIME g_calendarMonth{};
static HWND g_hotButton = nullptr;
static HWND g_hotCombo = nullptr;
static HWND g_hotCheckBox = nullptr;

// 詳細表示切り替え

// 詳細枠（視覚的なグループ化）
static HWND g_frameFolderExcl = nullptr;
static HWND g_frameNameExcl = nullptr;

// 実行時状態
static std::vector<ExcludeRule> g_excludeRules;
static std::vector<std::wstring> g_fileNamePatterns;
static std::vector<std::wstring> g_fileNameWild;
static std::vector<std::wstring> g_fileNameSubLow;

static std::vector<std::unique_ptr<Hit>> g_results;
static std::vector<size_t> g_visibleResultIndices;

static std::wstring g_iniPath;
static std::wstring g_currentScanDir; // 走査中フォルダー
static std::wstring g_lastExcludeFile;
static std::wstring g_lastCsvFile;
static std::wstring g_lastNameExcludeFile;
static unsigned long long g_totalScanFiles = 0;

static std::atomic<bool> g_stopRequested{ false };
static std::atomic<bool> g_searching{ false };
static HANDLE g_hThread = nullptr;

// 並べ替え: 0=日時,1=KB,2=ファイル名,3=パス
static int g_sortCol = 0;
static bool g_sortAsc = false;

// -------------------- 補助関数 --------------------
static std::wstring ToLower(std::wstring s) {
    std::transform(s.begin(), s.end(), s.begin(), [](wchar_t c) { return (wchar_t)towlower(c); });
    return s;
}
static std::wstring Trim(const std::wstring& s) {
    size_t a = 0, b = s.size();
    while (a < b && iswspace(s[a])) a++;
    while (b > a && iswspace(s[b - 1])) b--;
    return s.substr(a, b - a);
}



static std::wstring EllipsizePathRight(const std::wstring& s, size_t maxChars)
{
    if (s.size() <= maxChars) return s;
    if (maxChars <= 3) return s.substr(0, maxChars);
    return L"..." + s.substr(s.size() - (maxChars - 3));
}
// ---- 前方宣言（複数ルート補助関数用） ----
static std::wstring ToLower(std::wstring s);
static std::wstring Trim(const std::wstring& s);
static std::wstring GetWindowTextWStr(HWND h);
static void SetWindowTextWStr(HWND h, const std::wstring& s);
static fs::path NormalizePath(const fs::path& p);

// ---- 前方宣言（定義より前で使用する関数） ----
// 一部のハンドラー/確定処理が定義より前にあるため必要
static void RefreshFileNameListBox();
static void RebuildFileNameExcludeCache();
static void SaveSettings();
static void RedrawListBoxIfVisible(HWND hwnd);
// -------------------------------------------------------

// ---- ルート一覧補助関数（複数フォルダーを直感的に扱う） ----
static std::vector<std::wstring> GetRootsFromListBox()
{
    std::vector<std::wstring> out;
    if (!g_listRoots) return out;
    int n = (int)SendMessageW(g_listRoots, LB_GETCOUNT, 0, 0);
    if (n <= 0) return out;
    out.reserve(n);
    for (int i = 0; i < n; ++i) {
        wchar_t buf[2048]{};
        SendMessageW(g_listRoots, LB_GETTEXT, (WPARAM)i, (LPARAM)buf);
        auto t = Trim(std::wstring(buf));
        if (!t.empty()) out.push_back(t);
    }
    return out;
}

static const wchar_t* kRootEnabledPrefix = L"[有効] ";
static const wchar_t* kRootDisabledPrefix = L"[無効] ";

static std::wstring BuildRootDisplayText(const std::wstring& path, bool enabled)
{
    return std::wstring(enabled ? kRootEnabledPrefix : kRootDisabledPrefix) + Trim(path);
}

static void ParseRootDisplayText(const std::wstring& text, std::wstring& outPath, bool& outEnabled)
{
    std::wstring t = Trim(text);
    if (t.rfind(kRootEnabledPrefix, 0) == 0) {
        outEnabled = true;
        outPath = Trim(t.substr(wcslen(kRootEnabledPrefix)));
        return;
    }
    if (t.rfind(kRootDisabledPrefix, 0) == 0) {
        outEnabled = false;
        outPath = Trim(t.substr(wcslen(kRootDisabledPrefix)));
        return;
    }
    outEnabled = true;
    outPath = t;
}

struct RootEntry {
    std::wstring path;
    bool enabled = true;
};

static std::vector<RootEntry> GetRootEntriesFromListBox()
{
    std::vector<RootEntry> out;
    if (!g_listRoots) return out;

    int n = (int)SendMessageW(g_listRoots, LB_GETCOUNT, 0, 0);
    if (n <= 0) return out;

    out.reserve(n);
    for (int i = 0; i < n; ++i) {
        wchar_t buf[2048]{};
        SendMessageW(g_listRoots, LB_GETTEXT, (WPARAM)i, (LPARAM)buf);

        RootEntry e;
        ParseRootDisplayText(buf, e.path, e.enabled);
        if (!e.path.empty()) out.push_back(std::move(e));
    }
    return out;
}

static void SetRootEntriesToListBox(const std::vector<RootEntry>& entries)
{
    if (!g_listRoots) return;
    SendMessageW(g_listRoots, WM_SETREDRAW, FALSE, 0);
    SendMessageW(g_listRoots, LB_RESETCONTENT, 0, 0);
    for (const auto& e : entries) {
        auto t = Trim(e.path);
        if (t.empty()) continue;
        auto disp = BuildRootDisplayText(t, e.enabled);
        SendMessageW(g_listRoots, LB_ADDSTRING, 0, (LPARAM)disp.c_str());
    }
    SendMessageW(g_listRoots, WM_SETREDRAW, TRUE, 0);
    RedrawListBoxIfVisible(g_listRoots);
}

static std::vector<std::wstring> GetEnabledRootsFromListBox()
{
    std::vector<std::wstring> out;
    auto entries = GetRootEntriesFromListBox();
    out.reserve(entries.size());
    for (const auto& e : entries) {
        if (e.enabled && !e.path.empty()) out.push_back(e.path);
    }
    return out;
}

static bool ToggleSelectedRootEnabled()
{
    if (!g_listRoots) return false;

    int sel = (int)SendMessageW(g_listRoots, LB_GETCURSEL, 0, 0);
    if (sel == LB_ERR) return false;

    wchar_t buf[2048]{};
    SendMessageW(g_listRoots, LB_GETTEXT, (WPARAM)sel, (LPARAM)buf);

    std::wstring path;
    bool enabled = true;
    ParseRootDisplayText(buf, path, enabled);
    if (path.empty()) return false;

    std::wstring disp = BuildRootDisplayText(path, !enabled);
    SendMessageW(g_listRoots, LB_DELETESTRING, (WPARAM)sel, 0);
    int idx = (int)SendMessageW(g_listRoots, LB_INSERTSTRING, (WPARAM)sel, (LPARAM)disp.c_str());
    SendMessageW(g_listRoots, LB_SETCURSEL, (WPARAM)idx, 0);

    return true;
}

static bool RootExistsInListBox(const std::wstring& path)
{
    auto low = ToLower(Trim(path));
    auto roots = GetRootEntriesFromListBox();
    for (auto& r : roots) {
        if (ToLower(Trim(r.path)) == low) return true;
    }
    return false;
}

static void AddRootToListBoxDedup(const std::wstring& path)
{
    if (!g_listRoots) return;
    auto t = Trim(path);
    if (t.empty()) return;
    if (RootExistsInListBox(t)) return;
    auto disp = BuildRootDisplayText(t, true);
    SendMessageW(g_listRoots, LB_ADDSTRING, 0, (LPARAM)disp.c_str());
}

static void MoveRootItem(int from, int to)
{
    if (!g_listRoots) return;
    int n = (int)SendMessageW(g_listRoots, LB_GETCOUNT, 0, 0);
    if (from < 0 || from >= n || to < 0 || to >= n || from == to) return;

    wchar_t buf[2048]{};
    SendMessageW(g_listRoots, LB_GETTEXT, (WPARAM)from, (LPARAM)buf);
    std::wstring item = buf;

    // 元の項目を削除
    SendMessageW(g_listRoots, LB_DELETESTRING, (WPARAM)from, 0);

    // 移動先へ再挿入
    int idx = (int)SendMessageW(g_listRoots, LB_INSERTSTRING, (WPARAM)to, (LPARAM)item.c_str());
    SendMessageW(g_listRoots, LB_SETCURSEL, (WPARAM)idx, 0);
}

// ---------------------------------------------------

static std::vector<std::wstring> SplitRootsText(const std::wstring& text)
{
    // 区切り文字は ; と | を受け付ける（UIは1行のまま）。貼り付け時の改行も許可する
    std::vector<std::wstring> out;
    std::wstring cur;
    auto flush = [&]() {
        auto t = Trim(cur);
        if (!t.empty()) out.push_back(t);
        cur.clear();
        };
    for (wchar_t c : text) {
        if (c == L'\r') continue;
        if (c == L'\n' || c == L';' || c == L'|') { flush(); continue; }
        cur.push_back(c);
    }
    flush();

    // 大文字小文字を区別せず重複排除
    std::vector<std::wstring> uniq;
    for (auto& s : out) {
        auto low = ToLower(Trim(s));
        bool exists = false;
        for (auto& u : uniq) {
            if (ToLower(Trim(u)) == low) { exists = true; break; }
        }
        if (!exists) uniq.push_back(Trim(s));
    }
    return uniq;
}

static std::wstring JoinRootsForIni(const std::vector<std::wstring>& roots)
{
    std::wstring out;
    for (size_t i = 0; i < roots.size(); ++i) {
        if (i) out += L"|";
        out += roots[i];
    }
    return out;
}

static std::wstring GetWindowTextWStr(HWND h) {
    int len = GetWindowTextLengthW(h);
    std::wstring s(len, L'\0');
    GetWindowTextW(h, s.data(), len + 1);
    return s;
}
static void SetWindowTextWStr(HWND h, const std::wstring& s) { SetWindowTextW(h, s.c_str()); }
static bool IsChecked(HWND hChk) { return (SendMessageW(hChk, BM_GETCHECK, 0, 0) == BST_CHECKED); }
static void SetChecked(HWND hChk, bool v) { SendMessageW(hChk, BM_SETCHECK, v ? BST_CHECKED : BST_UNCHECKED, 0); }

static void SetStatus(const std::wstring& s) { if (g_status) SendMessageW(g_status, SB_SETTEXTW, 0, (LPARAM)s.c_str()); }

static std::wstring GetExeDir() {
    std::vector<wchar_t> buf(32768, L'\0');
    DWORD n = GetModuleFileNameW(nullptr, buf.data(), (DWORD)buf.size());
    if (n == 0 || n >= buf.size() - 1) return L"";
    fs::path p(buf.data());
    return p.parent_path().wstring();
}
static fs::path NormalizePath(const fs::path& p) {
    std::error_code ec;
    auto abs = fs::absolute(p, ec);
    if (ec) return p;
    auto can = fs::weakly_canonical(abs, ec);
    if (ec) return abs;
    return can;
}

static bool IsPathSeparator(wchar_t c) {
    return c == L'\\' || c == L'/';
}

static bool HasPathPrefixLow(const std::wstring& pathLow, const std::wstring& prefixLow) {
    // 正規化済みパス同士を小文字文字列で比較し、C:\foo と C:\foobar の誤判定を防ぐ
    if (prefixLow.empty() || pathLow.size() < prefixLow.size()) return false;
    if (pathLow.compare(0, prefixLow.size(), prefixLow) != 0) return false;
    if (pathLow.size() == prefixLow.size()) return true;
    if (IsPathSeparator(prefixLow.back())) return true;
    return IsPathSeparator(pathLow[prefixLow.size()]);
}

// ---- 日時処理 ----
static std::wstring FormatLocalTime(const std::chrono::system_clock::time_point& tp) {
    std::time_t tt = std::chrono::system_clock::to_time_t(tp);
    std::tm local{};
    localtime_s(&local, &tt);
    wchar_t buf[64];
    wcsftime(buf, 64, L"%Y-%m-%d %H:%M:%S", &local);
    return buf;
}

static void GetLocalDayRangePastNDays(int nDays,
    std::chrono::system_clock::time_point& outStart,
    std::chrono::system_clock::time_point& outEnd)
{
    using namespace std::chrono;
    if (nDays <= 0) nDays = 1;

    auto now = system_clock::now();
    std::time_t tnow = system_clock::to_time_t(now);

    std::tm localNow{};
    localtime_s(&localNow, &tnow);

    std::tm startTm = localNow;
    startTm.tm_hour = 0; startTm.tm_min = 0; startTm.tm_sec = 0;

    std::time_t today0 = mktime(&startTm); // ローカル日付の0時を time_t（UTC基準）へ変換
    std::time_t startT = today0 - (std::time_t)(nDays - 1) * 24 * 60 * 60;
    std::time_t endT = today0 + 24 * 60 * 60;

    outStart = system_clock::from_time_t(startT);
    outEnd = system_clock::from_time_t(endT);
}

static bool IsWithinLocalRange(const std::chrono::system_clock::time_point& tp,
    const std::chrono::system_clock::time_point& s,
    const std::chrono::system_clock::time_point& e) {
    return (tp >= s && tp < e);
}

static int GetMode() { // 0 今日, 1 過去N日, 2 期間指定
    LRESULT sel = SendMessageW(g_cmbMode, CB_GETCURSEL, 0, 0);
    if (sel == CB_ERR) return 0;
    return (int)sel;
}
static int GetDaysFromEdit() {
    std::wstring s = Trim(GetWindowTextWStr(g_editDays));
    int v = 1;
    if (!s.empty()) { try { v = std::stoi(s); } catch (...) { v = 1; } }
    if (v <= 0) v = 1;
    if (v > 3650) v = 3650;
    return v;
}

static std::chrono::system_clock::time_point LocalMidnightFromDate(const SYSTEMTIME& st)
{
    std::tm tm{};
    tm.tm_year = (int)st.wYear - 1900;
    tm.tm_mon = (int)st.wMonth - 1;
    tm.tm_mday = (int)st.wDay;
    tm.tm_hour = 0; tm.tm_min = 0; tm.tm_sec = 0;
    std::time_t t = mktime(&tm); // ローカル時刻
    return std::chrono::system_clock::from_time_t(t);
}

static void GetActiveDateRange(std::chrono::system_clock::time_point& outS,
    std::chrono::system_clock::time_point& outE)
{
    int mode = GetMode();
    if (mode == 0) {
        GetLocalDayRangePastNDays(1, outS, outE); // 今日
        return;
    }
    if (mode == 1) {
        GetLocalDayRangePastNDays(GetDaysFromEdit(), outS, outE); // 過去N日
        return;
    }

    // mode が 2 の場合: 期間指定（カレンダー）
    SYSTEMTIME stFrom = g_dateFrom;
    SYSTEMTIME stTo = g_dateTo;
    if (stFrom.wYear == 0 || stTo.wYear == 0) {
        // 万一取れない場合は「今日」にフォールバック
        GetLocalDayRangePastNDays(1, outS, outE);
        return;
    }

    auto s = LocalMidnightFromDate(stFrom);
    auto e0 = LocalMidnightFromDate(stTo);

    // 終了日は「その日の終わりまで」を含めたいので、翌日0:00を上限にする（[s, e)）
    auto e = e0 + std::chrono::hours(24);

    if (s > e) std::swap(s, e);

    outS = s;
    outE = e;
}

// -------------------- 進捗表示補助関数 --------------------
static void Progress_SetMarquee(HWND hProg, bool on) {
    if (!hProg) return;

    LONG_PTR style = GetWindowLongPtrW(hProg, GWL_STYLE);

    if (on) {
        // PBS_MARQUEE が設定されていることを保証する
        if (!(style & PBS_MARQUEE)) {
            SetWindowLongPtrW(hProg, GWL_STYLE, style | PBS_MARQUEE);
            SetWindowPos(hProg, nullptr, 0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
        }
        SendMessageW(hProg, PBM_SETMARQUEE, TRUE, 30);
    }
    else {
        // 先にマーキー表示を停止する
        SendMessageW(hProg, PBM_SETMARQUEE, FALSE, 0);

        // PBM_SETPOSでバーが空表示になるよう PBS_MARQUEE を外す
        style = GetWindowLongPtrW(hProg, GWL_STYLE);
        if (style & PBS_MARQUEE) {
            SetWindowLongPtrW(hProg, GWL_STYLE, style & ~PBS_MARQUEE);
            SetWindowPos(hProg, nullptr, 0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
        }

        SendMessageW(hProg, PBM_SETRANGE32, 0, 100);
        SendMessageW(hProg, PBM_SETPOS, 0, 0);

        // 空状態を反映するため強制的に再描画する
        InvalidateRect(hProg, nullptr, TRUE);
        UpdateWindow(hProg);
    }
}

// ---- FILETIME から system_clock への変換 ----
static std::chrono::system_clock::time_point FileTimeToSysClock(const FILETIME& ft)
{
    ULARGE_INTEGER uli;
    uli.LowPart = ft.dwLowDateTime;
    uli.HighPart = ft.dwHighDateTime;

    constexpr int64_t WINDOWS_TICK_PER_SEC = 10'000'000LL;     // 100ナノ秒 * 10,000,000 = 1秒
    constexpr int64_t SEC_TO_UNIX_EPOCH = 11'644'473'600LL; // 1601-01-01 から 1970-01-01 までの秒数

    int64_t total100ns = (int64_t)uli.QuadPart;
    int64_t sec = total100ns / WINDOWS_TICK_PER_SEC;
    int64_t rem = total100ns % WINDOWS_TICK_PER_SEC; // 100ナノ秒単位の余り

    // ここがポイント：system_clock::duration に明示キャストして精度落ちを許可
    auto dur = std::chrono::seconds(sec - SEC_TO_UNIX_EPOCH)
        + std::chrono::duration_cast<std::chrono::system_clock::duration>(
            std::chrono::nanoseconds(rem * 100LL)
        );

    return std::chrono::system_clock::time_point(dur);
}

static TimeBase GetTimeBase() {
    LRESULT sel = SendMessageW(g_cmbTimeBase, CB_GETCURSEL, 0, 0);
    if (sel == CB_ERR) return TimeBase::LastWrite;
    if (sel == 1) return TimeBase::Creation;
    if (sel == 2) return TimeBase::Either;
    return TimeBase::LastWrite;
}

static std::wstring TimeBaseText(TimeBase tb) {
    if (tb == TimeBase::Creation) return L"作成日時";
    if (tb == TimeBase::Either)   return L"更新OR作成";
    return L"更新日時";
}

// 更新日時/作成日時/いずれかを取得
static bool GetFileTimesSysClock(const std::wstring& path,
    std::chrono::system_clock::time_point& outWrite,
    std::chrono::system_clock::time_point& outCreate)
{
    WIN32_FILE_ATTRIBUTE_DATA fad{};
    if (!GetFileAttributesExW(path.c_str(), GetFileExInfoStandard, &fad)) return false;
    outWrite = FileTimeToSysClock(fad.ftLastWriteTime);
    outCreate = FileTimeToSysClock(fad.ftCreationTime);
    return true;
}

// ---- 拡張子 ----
static bool ExtChecked(const std::wstring& extLower) {
    if (extLower == L".xls")  return IsChecked(g_chkXls);
    if (extLower == L".xlsx") return IsChecked(g_chkXlsx);
    if (extLower == L".xlsm") return IsChecked(g_chkXlsm);
    if (extLower == L".xlsb") return IsChecked(g_chkXlsb);
    if (extLower == L".xltx") return IsChecked(g_chkXltx);
    if (extLower == L".xltm") return IsChecked(g_chkXltm);
    return false;
}
static bool AnyExtSelected() {
    return IsChecked(g_chkXls) || IsChecked(g_chkXlsx) || IsChecked(g_chkXlsm) ||
        IsChecked(g_chkXlsb) || IsChecked(g_chkXltx) || IsChecked(g_chkXltm);
}

static bool UseTargetExtensionFilter() {
    return IsChecked(g_chkNameIncludeExt);
}

static void SetTargetExtensionSectionEnabled(bool enabled) {
    HWND hExtGrp = g_hwndMain ? GetDlgItem(g_hwndMain, IDC_GRP_EXT) : nullptr;
    HWND controls[] = { hExtGrp, g_chkXls, g_chkXlsx, g_chkXlsm, g_chkXlsb, g_chkXltx, g_chkXltm };
    for (HWND h : controls) {
        if (!h) continue;
        EnableWindow(h, enabled ? TRUE : FALSE);
        InvalidateRect(h, nullptr, TRUE);
    }
}

static bool IsTargetExcelFile(const fs::path& p) {
    if (!UseTargetExtensionFilter()) return true;

    std::wstring e = ToLower(p.extension().wstring());
    if (e == L".xls" || e == L".xlsx" || e == L".xlsm" || e == L".xlsb" || e == L".xltx" || e == L".xltm") {
        return ExtChecked(e);
    }
    return false;
}

// -------------------- UTF-8 補助関数 --------------------
static std::string WideToUtf8(const std::wstring& w) {
    if (w.empty()) return {};
    int sz = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), nullptr, 0, nullptr, nullptr);
    std::string out(sz, '\0');
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), out.data(), sz, nullptr, nullptr);
    return out;
}
static std::wstring Utf8ToWide(const std::string& s) {
    if (s.empty()) return {};
    int sz = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(), nullptr, 0);
    std::wstring out(sz, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(), out.data(), sz);
    return out;
}
static bool ReadTextFileUtf8(const std::wstring& path, std::wstring& outText) {
    outText.clear();
    std::ifstream ifs(fs::path(path), std::ios::binary);
    if (!ifs) return false;
    std::string bytes((std::istreambuf_iterator<char>(ifs)), std::istreambuf_iterator<char>());
    if (bytes.size() >= 3 && (unsigned char)bytes[0] == 0xEF && (unsigned char)bytes[1] == 0xBB && (unsigned char)bytes[2] == 0xBF) {
        bytes.erase(0, 3);
    }
    outText = Utf8ToWide(bytes);
    return true;
}
static bool WriteTextFileUtf8Bom(const std::wstring& path, const std::wstring& text) {
    std::ofstream ofs(fs::path(path), std::ios::binary);
    if (!ofs) return false;
    unsigned char bom[3] = { 0xEF, 0xBB, 0xBF };
    ofs.write((const char*)bom, 3);
    auto u8 = WideToUtf8(text);
    ofs.write(u8.data(), (std::streamsize)u8.size());
    return true;
}

// -------------------- ダイアログ補助関数 --------------------
static bool PickFolder(HWND owner, std::wstring& outPath) {
    outPath.clear();
    IFileDialog* pfd = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pfd));
    if (FAILED(hr) || !pfd) return false;

    DWORD opts = 0;
    pfd->GetOptions(&opts);
    pfd->SetOptions(opts | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);

    hr = pfd->Show(owner);
    if (SUCCEEDED(hr)) {
        IShellItem* psi = nullptr;
        hr = pfd->GetResult(&psi);
        if (SUCCEEDED(hr) && psi) {
            PWSTR psz = nullptr;
            hr = psi->GetDisplayName(SIGDN_FILESYSPATH, &psz);
            if (SUCCEEDED(hr) && psz) {
                outPath = psz;
                CoTaskMemFree(psz);
            }
            psi->Release();
        }
    }
    pfd->Release();
    return !outPath.empty();
}
static bool PickOpenFile(HWND owner, const wchar_t* filter, std::wstring& inoutPath) {
    std::vector<wchar_t> buf(32768, L'\0');
    if (!inoutPath.empty()) wcsncpy_s(buf.data(), buf.size(), inoutPath.c_str(), _TRUNCATE);

    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = owner;
    ofn.lpstrFilter = filter;
    ofn.lpstrFile = buf.data();
    ofn.nMaxFile = (DWORD)buf.size();
    ofn.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR;

    if (GetOpenFileNameW(&ofn)) {
        inoutPath = buf.data();
        return true;
    }

    DWORD err = CommDlgExtendedError(); // 0 はユーザーキャンセル
    if (err != 0) {
        wchar_t msg[128];
        wsprintfW(msg, L"GetOpenFileNameW failed: 0x%08X", err);
        MessageBoxW(owner, msg, L"Dialog error", MB_OK | MB_ICONERROR);
    }
    return false;
}

static bool PickSaveFile(HWND owner, const wchar_t* filter, const wchar_t* defExt, std::wstring& inoutPath) {
    std::vector<wchar_t> buf(32768, L'\0');
    if (!inoutPath.empty()) wcsncpy_s(buf.data(), buf.size(), inoutPath.c_str(), _TRUNCATE);

    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = owner;
    ofn.lpstrFilter = filter;
    ofn.lpstrDefExt = defExt;
    ofn.lpstrFile = buf.data();
    ofn.nMaxFile = (DWORD)buf.size();
    ofn.Flags = OFN_EXPLORER | OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR;

    if (GetSaveFileNameW(&ofn)) {
        inoutPath = buf.data();
        return true;
    }

    DWORD err = CommDlgExtendedError(); // 0 はユーザーキャンセル
    if (err != 0) {
        wchar_t msg[128];
        wsprintfW(msg, L"GetSaveFileNameW failed: 0x%08X", err);
        MessageBoxW(owner, msg, L"Dialog error", MB_OK | MB_ICONERROR);
    }
    return false;
}

// -------------------- フォルダー除外 --------------------
static std::wstring FolderRuleToDisplay(const ExcludeRule& r) {
    switch (r.type) {
    case ExcludeType::DirPrefix: return L"[DIR] " + r.dirNorm.wstring();
    case ExcludeType::Wildcard:  return L"[WILD] " + r.raw;
    case ExcludeType::Substring: return L"[SUB] " + r.raw;
    default: return r.raw;
    }
}
static void RefreshExcludeListBox() {
    if (g_deferExcludeListRefresh || !g_listExcludes) return;
    SendMessageW(g_listExcludes, WM_SETREDRAW, FALSE, 0);
    SendMessageW(g_listExcludes, LB_RESETCONTENT, 0, 0);
    for (const auto& r : g_excludeRules) {
        auto disp = FolderRuleToDisplay(r);
        SendMessageW(g_listExcludes, LB_ADDSTRING, 0, (LPARAM)disp.c_str());
    }
    SendMessageW(g_listExcludes, WM_SETREDRAW, TRUE, 0);
    RedrawListBoxIfVisible(g_listExcludes);
}
static bool IsExcludedDir(const fs::path& dirNorm) {
    const std::wstring dirStr = dirNorm.wstring();
    const std::wstring dirLower = ToLower(dirStr);

    for (const auto& r : g_excludeRules) {
        if (r.type == ExcludeType::DirPrefix) {
            // パス接頭辞はキャッシュ済み小文字文字列で比較し、探索中のパス分解コストを抑える
            if (HasPathPrefixLow(dirLower, r.dirNormLow)) return true;
        }
        else if (r.type == ExcludeType::Wildcard) {
            if (PathMatchSpecW(dirStr.c_str(), r.raw.c_str())) return true;
        }
        else if (r.type == ExcludeType::Substring) {
            if (!r.pattern.empty() && dirLower.find(r.pattern) != std::wstring::npos) return true;
        }
    }
    return false;
}
static void AddExcludeDirPrefix(const std::wstring& folderPath) {
    ExcludeRule r;
    r.type = ExcludeType::DirPrefix;
    r.dirNorm = NormalizePath(fs::path(folderPath));
    r.raw = r.dirNorm.wstring();
    r.rawLow = ToLower(r.raw);
    r.dirNormLow = r.rawLow;

    for (auto& e : g_excludeRules) {
        if (e.type == ExcludeType::DirPrefix && e.dirNormLow == r.dirNormLow) return;
    }
    g_excludeRules.push_back(std::move(r));
    RefreshExcludeListBox();
}
static void AddOrUpdateExcludePatternOrSubstring(const std::wstring& text, int targetIndexOrMinus1) {
    std::wstring t = Trim(text);
    if (t.empty()) return;

    ExcludeRule r;
    r.raw = t;
    bool hasWild = (t.find(L'*') != std::wstring::npos) || (t.find(L'?') != std::wstring::npos);
    if (hasWild) {
        r.type = ExcludeType::Wildcard;
        r.rawLow = ToLower(r.raw);
        if (targetIndexOrMinus1 < 0) {
            for (auto& e : g_excludeRules) {
                if (e.type == ExcludeType::Wildcard && e.rawLow == r.rawLow) return;
            }
        }
    }
    else {
        r.type = ExcludeType::Substring;
        r.rawLow = ToLower(r.raw);
        r.pattern = r.rawLow;
        if (targetIndexOrMinus1 < 0) {
            for (auto& e : g_excludeRules) {
                if (e.type == ExcludeType::Substring && e.pattern == r.pattern) return;
            }
        }
    }

    if (targetIndexOrMinus1 >= 0 && targetIndexOrMinus1 < (int)g_excludeRules.size()) {
        if (g_excludeRules[(size_t)targetIndexOrMinus1].type == ExcludeType::DirPrefix) return;
        g_excludeRules[(size_t)targetIndexOrMinus1] = std::move(r);
    }
    else {
        g_excludeRules.push_back(std::move(r));
    }
    RefreshExcludeListBox();
}

static bool LoadExcludesFromFile(const std::wstring& filePath, const fs::path& rootForRelative) {
    std::wstring text;
    if (!ReadTextFileUtf8(filePath, text)) return false;

    std::vector<ExcludeRule> loaded;

    size_t pos = 0;
    while (pos <= text.size()) {
        size_t nl = text.find(L'\n', pos);
        std::wstring line = (nl == std::wstring::npos) ? text.substr(pos) : text.substr(pos, nl - pos);
        if (!line.empty() && line.back() == L'\r') line.pop_back();
        pos = (nl == std::wstring::npos) ? text.size() + 1 : nl + 1;

        line = Trim(line);
        if (line.empty()) continue;
        if (line[0] == L'#' || line[0] == L';') continue;

        bool hasWild = (line.find(L'*') != std::wstring::npos) || (line.find(L'?') != std::wstring::npos);
        if (hasWild) {
            ExcludeRule r;
            r.type = ExcludeType::Wildcard;
            r.raw = line;
            r.rawLow = ToLower(r.raw);
            loaded.push_back(std::move(r));
            continue;
        }

        fs::path p(line);
        if (!p.is_absolute()) p = rootForRelative / p;
        std::error_code ec;
        if (fs::exists(p, ec) && fs::is_directory(p, ec)) {
            ExcludeRule r;
            r.type = ExcludeType::DirPrefix;
            r.dirNorm = NormalizePath(p);
            r.raw = r.dirNorm.wstring();
            r.rawLow = ToLower(r.raw);
            r.dirNormLow = r.rawLow;
            loaded.push_back(std::move(r));
        }
        else {
            ExcludeRule r;
            r.type = ExcludeType::Substring;
            r.raw = line;
            r.rawLow = ToLower(r.raw);
            r.pattern = r.rawLow;
            loaded.push_back(std::move(r));
        }
    }

    auto key = [](const ExcludeRule& r) {
        // 読み込み直後に作った小文字キャッシュを重複排除にも使う
        if (r.type == ExcludeType::DirPrefix) return L"DIR:" + r.dirNormLow;
        if (r.type == ExcludeType::Wildcard)  return L"WILD:" + r.rawLow;
        return L"SUB:" + r.pattern;
        };
    std::sort(loaded.begin(), loaded.end(), [&](const ExcludeRule& a, const ExcludeRule& b) { return key(a) < key(b); });
    loaded.erase(std::unique(loaded.begin(), loaded.end(), [&](const ExcludeRule& a, const ExcludeRule& b) { return key(a) == key(b); }), loaded.end());

    g_excludeRules = std::move(loaded);
    RefreshExcludeListBox();
    return true;
}
static bool SaveExcludesToFile(const std::wstring& filePath) {
    std::wstring out;
    for (const auto& r : g_excludeRules) {
        if (r.type == ExcludeType::DirPrefix) out += r.dirNorm.wstring();
        else out += r.raw;
        out += L"\r\n";
    }
    return WriteTextFileUtf8Bom(filePath, out);
}

// -------------------- 除外設定: インライン更新補助関数 --------------------
static void UpdateExcludeDirPrefixAt(int index, const std::wstring& newPath)
{
    if (index < 0 || index >= (int)g_excludeRules.size()) return;
    auto t = Trim(newPath);
    if (t.empty()) return;

    fs::path p(t);
    auto norm = NormalizePath(p);
    auto low = ToLower(norm.wstring());

    // 他のフォルダー除外と重複しないか確認する
    for (int i = 0; i < (int)g_excludeRules.size(); ++i) {
        if (i == index) continue;
        const auto& r = g_excludeRules[(size_t)i];
        if (r.type == ExcludeType::DirPrefix && r.dirNormLow == low) {
            MessageBoxW(g_hwndMain, L"同じ除外フォルダが既に存在します。", L"更新できません", MB_OK | MB_ICONINFORMATION);
            return;
        }
    }

    auto& r = g_excludeRules[(size_t)index];
    r.type = ExcludeType::DirPrefix;
    r.dirNorm = norm;
    r.raw = norm.wstring();
    r.rawLow = low;
    r.dirNormLow = low;
    r.pattern.clear();
    RefreshExcludeListBox();
    SendMessageW(g_listExcludes, LB_SETCURSEL, (WPARAM)index, 0);
}

static void CommitExcludeEditIfNeeded()
{
    if (!g_listExcludes || !g_editExclPattern) return;
    int sel = (int)SendMessageW(g_listExcludes, LB_GETCURSEL, 0, 0);
    if (sel == LB_ERR) return;

    auto t = Trim(GetWindowTextWStr(g_editExclPattern));
    if (t.empty()) return;
    if (sel < 0 || sel >= (int)g_excludeRules.size()) return;

    auto& r = g_excludeRules[(size_t)sel];
    if (r.type == ExcludeType::DirPrefix) {
        UpdateExcludeDirPrefixAt(sel, t);
    }
    else {
        AddOrUpdateExcludePatternOrSubstring(t, sel);
    }
    SaveSettings();
}

static bool FileNamePatternExistsExcept(const std::wstring& value, int exceptIndex)
{
    auto low = ToLower(Trim(value));
    if (low.empty()) return false;
    for (int i = 0; i < (int)g_fileNamePatterns.size(); ++i) {
        if (i == exceptIndex) continue;
        if (ToLower(Trim(g_fileNamePatterns[(size_t)i])) == low) return true;
    }
    return false;
}

static void CommitFileNameEditIfNeeded()
{
    if (!g_listFName || !g_editFNamePattern) return;
    int sel = (int)SendMessageW(g_listFName, LB_GETCURSEL, 0, 0);
    if (sel == LB_ERR) return;
    if (sel < 0 || sel >= (int)g_fileNamePatterns.size()) return;

    auto t = Trim(GetWindowTextWStr(g_editFNamePattern));
    if (t.empty()) return;
    if (FileNamePatternExistsExcept(t, sel)) {
        MessageBoxW(g_hwndMain, L"同じ除外ファイル名が既に存在します。", L"更新できません", MB_OK | MB_ICONINFORMATION);
        return;
    }

    g_fileNamePatterns[(size_t)sel] = t;
    RefreshFileNameListBox();
    SendMessageW(g_listFName, LB_SETCURSEL, (WPARAM)sel, 0);
    RebuildFileNameExcludeCache();
    SaveSettings();
}

static LRESULT CALLBACK ExclEditProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    if (msg == WM_KEYDOWN) {
        // 新しく入力を始めたら一覧選択を解除して「新規追加モード」にする
        if ((wParam >= 0x30 && wParam <= 0x5A) ||
            (wParam >= VK_NUMPAD0 && wParam <= VK_DIVIDE) ||
            wParam == VK_OEM_1 || wParam == VK_OEM_2 || wParam == VK_OEM_3 ||
            wParam == VK_OEM_4 || wParam == VK_OEM_5 || wParam == VK_OEM_6 ||
            wParam == VK_OEM_7 || wParam == VK_OEM_MINUS || wParam == VK_OEM_PLUS ||
            wParam == VK_SPACE || wParam == VK_BACK || wParam == VK_DELETE)
        {
            if (g_listExcludes) {
                SendMessageW(g_listExcludes, LB_SETCURSEL, (WPARAM)-1, 0);
            }
        }

        if (wParam == VK_RETURN) {
            int sel = g_listExcludes ? (int)SendMessageW(g_listExcludes, LB_GETCURSEL, 0, 0) : LB_ERR;
            if (sel == LB_ERR) {
                return 0;
            }

            CommitExcludeEditIfNeeded();
            return 0;
        }

        if (wParam == VK_ESCAPE) {
            int sel = (int)SendMessageW(g_listExcludes, LB_GETCURSEL, 0, 0);
            if (sel != LB_ERR && sel >= 0 && sel < (int)g_excludeRules.size()) {
                SetWindowTextWStr(g_editExclPattern, g_excludeRules[(size_t)sel].raw);
                SendMessageW(g_editExclPattern, EM_SETSEL, 0, -1);
            }
            return 0;
        }
    }
    return CallWindowProcW(g_oldExclEditProc, hWnd, msg, wParam, lParam);
}

static LRESULT CALLBACK FNameEditProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    if (msg == WM_KEYDOWN) {
        // 文字入力を始めたら、一覧の選択を外して「新規追加モード」にする
        if ((wParam >= 0x30 && wParam <= 0x5A) ||   // 0-9、A-Z
            (wParam >= VK_NUMPAD0 && wParam <= VK_DIVIDE) ||
            wParam == VK_OEM_1 || wParam == VK_OEM_2 || wParam == VK_OEM_3 ||
            wParam == VK_OEM_4 || wParam == VK_OEM_5 || wParam == VK_OEM_6 ||
            wParam == VK_OEM_7 || wParam == VK_OEM_MINUS || wParam == VK_OEM_PLUS ||
            wParam == VK_SPACE || wParam == VK_BACK || wParam == VK_DELETE)
        {
            if (g_listFName) {
                SendMessageW(g_listFName, LB_SETCURSEL, (WPARAM)-1, 0);
            }
        }

        if (wParam == VK_RETURN) {
            // 選択が無いなら更新せず、そのまま新規追加扱い
            int sel = g_listFName ? (int)SendMessageW(g_listFName, LB_GETCURSEL, 0, 0) : LB_ERR;
            if (sel == LB_ERR) {
                return 0;
            }

            CommitFileNameEditIfNeeded();
            return 0;
        }

        if (wParam == VK_ESCAPE) {
            // ESCでは一覧選択があるときだけ元の値に戻す
            int sel = (int)SendMessageW(g_listFName, LB_GETCURSEL, 0, 0);
            if (sel != LB_ERR && sel >= 0 && sel < (int)g_fileNamePatterns.size()) {
                SetWindowTextWStr(g_editFNamePattern, g_fileNamePatterns[(size_t)sel]);
                SendMessageW(g_editFNamePattern, EM_SETSEL, 0, -1);
            }
            return 0;
        }
    }
    return CallWindowProcW(g_oldFNameEditProc, hWnd, msg, wParam, lParam);
}


// -------------------- ファイル名除外 --------------------
static void RebuildFileNameExcludeCache() {
    g_fileNameWild.clear();
    g_fileNameSubLow.clear();
    for (const auto& raw : g_fileNamePatterns) {
        auto t = Trim(raw);
        if (t.empty()) continue;
        bool hasWild = (t.find(L'*') != std::wstring::npos) || (t.find(L'?') != std::wstring::npos);
        if (hasWild) g_fileNameWild.push_back(t);
        else g_fileNameSubLow.push_back(ToLower(t));
    }
}
static void RefreshFileNameListBox() {
    if (g_deferFileNameListRefresh || !g_listFName) return;
    SendMessageW(g_listFName, WM_SETREDRAW, FALSE, 0);
    SendMessageW(g_listFName, LB_RESETCONTENT, 0, 0);
    for (const auto& s : g_fileNamePatterns) {
        SendMessageW(g_listFName, LB_ADDSTRING, 0, (LPARAM)s.c_str());
    }
    SendMessageW(g_listFName, WM_SETREDRAW, TRUE, 0);
    RedrawListBoxIfVisible(g_listFName);
}
static bool IsExcludedByFileName(const fs::path& p) {
    std::wstring name = p.filename().wstring();
    if (!IsChecked(g_chkNameIncludeExt)) name = p.stem().wstring();
    std::wstring nameLow = ToLower(name);

    for (const auto& sub : g_fileNameSubLow) {
        if (!sub.empty() && nameLow.find(sub) != std::wstring::npos) return true;
    }
    for (const auto& pat : g_fileNameWild) {
        if (PathMatchSpecW(name.c_str(), pat.c_str())) return true;
    }
    return false;
}


static bool LoadFileNameExcludesFromFile(const std::wstring& filePath) {
    std::wstring text;
    if (!ReadTextFileUtf8(filePath, text)) return false;

    std::vector<std::wstring> loaded;
    size_t pos = 0;
    while (pos <= text.size()) {
        size_t nl = text.find(L'\n', pos);
        std::wstring line = (nl == std::wstring::npos) ? text.substr(pos) : text.substr(pos, nl - pos);
        if (!line.empty() && line.back() == L'\r') line.pop_back();
        pos = (nl == std::wstring::npos) ? text.size() + 1 : nl + 1;

        line = Trim(line);
        if (line.empty()) continue;
        if (line[0] == L'#' || line[0] == L';') continue;

        loaded.push_back(line);
    }

    // 正規化して大文字小文字を区別せず一意化
    auto key = [](const std::wstring& s) { return ToLower(Trim(s)); };
    std::sort(loaded.begin(), loaded.end(), [&](const std::wstring& a, const std::wstring& b) { return key(a) < key(b); });
    loaded.erase(std::unique(loaded.begin(), loaded.end(), [&](const std::wstring& a, const std::wstring& b) { return key(a) == key(b); }), loaded.end());

    g_fileNamePatterns = std::move(loaded);
    if (g_fileNamePatterns.empty()) g_fileNamePatterns.push_back(L"~$");

    RefreshFileNameListBox();
    RebuildFileNameExcludeCache();
    return true;
}

static bool SaveFileNameExcludesToFile(const std::wstring& filePath) {
    std::wstring out;
    for (const auto& s : g_fileNamePatterns) {
        auto t = Trim(s);
        if (t.empty()) continue;
        out += t;
        out += L"\r\n";
    }
    return WriteTextFileUtf8Bom(filePath, out);
}

// -------------------- 結果フィルター --------------------
static std::wstring GetFilterLow() {
    if (!g_editFilter) return L"";
    return ToLower(Trim(GetWindowTextWStr(g_editFilter)));
}
static bool HitMatchesFilterLow(const Hit& h, const std::wstring& fLow) {
    if (fLow.empty()) return true;
    // ヒット作成時に保存した小文字キャッシュを使い、フィルター変更時の再描画を軽くする
    return (h.fileNameLow.find(fLow) != std::wstring::npos) ||
        (h.pathLow.find(fLow) != std::wstring::npos);
}
static void UpdateExportButtonEnabled() {
    if (!g_btnExportCsv) return;
    if (g_searching) {
        EnableWindow(g_btnExportCsv, FALSE);
        return;
    }
    EnableWindow(g_btnExportCsv, (g_visibleCount > 0) ? TRUE : FALSE);
}

// -------------------- CSV出力 --------------------
static std::wstring CsvEscape(const std::wstring& s) {
    bool need = false;
    for (wchar_t c : s) {
        if (c == L',' || c == L'"' || c == L'\n' || c == L'\r') { need = true; break; }
    }
    if (!need) return s;
    std::wstring t;
    t.reserve(s.size() + 2);
    t.push_back(L'"');
    for (wchar_t c : s) {
        if (c == L'"') t += L"\"\"";
        else t.push_back(c);
    }
    t.push_back(L'"');
    return t;
}
static bool ExportResultsCsv(const std::wstring& filePath, TimeBase tb) {
    // 先頭列名も基準に合わせる
    const std::wstring fLow = GetFilterLow();
    std::wstring out;
    out += (tb == TimeBase::Creation) ? L"CreationTime,SizeKB,FileName,Path\r\n"
        : (tb == TimeBase::Either ? L"WriteOrCreateTime,SizeKB,FileName,Path\r\n"
            : L"LastWriteTime,SizeKB,FileName,Path\r\n");
    for (const auto& hp : g_results) {
        const auto& h = *hp;
        if (!HitMatchesFilterLow(h, fLow)) continue;
        out += CsvEscape(h.timeText); out += L",";
        out += std::to_wstring(h.sizeKB); out += L",";
        out += CsvEscape(h.fileName); out += L",";
        out += CsvEscape(h.path); out += L"\r\n";
    }
    return WriteTextFileUtf8Bom(filePath, out);
}

// -------------------- リストビュー --------------------
static void InitListViewColumns(HWND lv) {
    ListView_SetExtendedListViewStyle(lv, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_HEADERDRAGDROP);

    LVCOLUMNW col{};
    col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM | LVCF_FMT;

    col.pszText = (LPWSTR)L"日時";
    col.cx = 170; col.iSubItem = 0; col.fmt = LVCFMT_LEFT;
    ListView_InsertColumn(lv, 0, &col);

    col.pszText = (LPWSTR)L"KB";
    col.cx = 80; col.iSubItem = 1; col.fmt = LVCFMT_RIGHT;
    ListView_InsertColumn(lv, 1, &col);

    col.pszText = (LPWSTR)L"ファイル名";
    col.cx = 240; col.iSubItem = 2; col.fmt = LVCFMT_LEFT;
    ListView_InsertColumn(lv, 2, &col);

    col.pszText = (LPWSTR)L"パス";
    col.cx = 680; col.iSubItem = 3; col.fmt = LVCFMT_LEFT;
    ListView_InsertColumn(lv, 3, &col);
}

static void SetListViewTimeHeader(TimeBase tb) {
    if (!g_listResults) return;
    LVCOLUMNW col{};
    col.mask = LVCF_TEXT;
    std::wstring t = (tb == TimeBase::Creation) ? L"作成日時" : (tb == TimeBase::Either ? L"更新/作成" : L"更新日時");
    col.pszText = (LPWSTR)t.c_str();
    ListView_SetColumn(g_listResults, 0, &col);
}

static void UpdateResultDetailFromSelection();

static void ClearResultsUI() {
    g_results.clear();
    g_visibleResultIndices.clear();
    g_visibleCount = 0;
    ListView_DeleteAllItems(g_listResults);
    UpdateExportButtonEnabled();
    UpdateResultDetailFromSelection();
}

static void AddResultToUI(std::unique_ptr<Hit> hit) {
    g_results.push_back(std::move(hit));
    const size_t realIndex = g_results.size() - 1;
    const Hit& h = *g_results.back();

    const std::wstring fLow = GetFilterLow();
    if (!HitMatchesFilterLow(h, fLow)) {
        return;
    }

    LVITEMW item{};
    item.mask = LVIF_TEXT;
    item.iItem = ListView_GetItemCount(g_listResults);
    item.pszText = (LPWSTR)h.timeText.c_str();
    int idx = ListView_InsertItem(g_listResults, &item);

    std::wstring kb = std::to_wstring(h.sizeKB);
    ListView_SetItemText(g_listResults, idx, 1, kb.data());
    ListView_SetItemText(g_listResults, idx, 2, (LPWSTR)h.fileName.c_str());
    ListView_SetItemText(g_listResults, idx, 3, (LPWSTR)h.path.c_str());

    g_visibleResultIndices.push_back(realIndex);
    g_visibleCount++;
    UpdateExportButtonEnabled();
}

static void RebuildListViewFromResults() {
    const std::wstring fLow = GetFilterLow();
    ListView_DeleteAllItems(g_listResults);
    g_visibleResultIndices.clear();
    g_visibleCount = 0;

    for (size_t i = 0; i < g_results.size(); ++i) {
        const Hit& h = *g_results[i];
        if (!HitMatchesFilterLow(h, fLow)) continue;

        LVITEMW item{};
        item.mask = LVIF_TEXT;
        item.iItem = ListView_GetItemCount(g_listResults);
        item.pszText = (LPWSTR)h.timeText.c_str();
        int idx = ListView_InsertItem(g_listResults, &item);

        std::wstring kb = std::to_wstring(h.sizeKB);
        ListView_SetItemText(g_listResults, idx, 1, kb.data());
        ListView_SetItemText(g_listResults, idx, 2, (LPWSTR)h.fileName.c_str());
        ListView_SetItemText(g_listResults, idx, 3, (LPWSTR)h.path.c_str());

        g_visibleResultIndices.push_back(i);
        g_visibleCount++;
    }
    UpdateExportButtonEnabled();
    UpdateResultDetailFromSelection();
}

static void SortResults(int col, bool asc) {
    if (g_results.empty()) return;

    auto cmp = [&](const std::unique_ptr<Hit>& a, const std::unique_ptr<Hit>& b) {
        const Hit& A = *a;
        const Hit& B = *b;
        int r = 0;
        if (col == 0) {
            // YYYY-MM-DD HH:MM:SS 形式なので文字列比較で問題ない
            r = (A.timeText < B.timeText) ? -1 : (A.timeText > B.timeText ? 1 : 0);
        }
        else if (col == 1) {
            r = (A.sizeKB < B.sizeKB) ? -1 : (A.sizeKB > B.sizeKB ? 1 : 0);
        }
        else if (col == 2) {
            // ファイル名/パスは小文字キャッシュで比較し、ソート中の一時文字列生成を減らす
            r = (A.fileNameLow < B.fileNameLow) ? -1 : (A.fileNameLow > B.fileNameLow ? 1 : 0);
        }
        else {
            r = (A.pathLow < B.pathLow) ? -1 : (A.pathLow > B.pathLow ? 1 : 0);
        }
        return asc ? (r < 0) : (r > 0);
        };

    std::stable_sort(g_results.begin(), g_results.end(), cmp);
    RebuildListViewFromResults();
}

// -------------------- INI設定 --------------------
// settings.ini を UTF-16LE(BOM) で確実に作成し、必要なら旧(サーバー)INIからローカルへ移行する
static bool ContainsQuestionMark(const std::wstring& s) {
    return s.find(L'?') != std::wstring::npos;
}

static void CreateUtf16LeBomFile(const std::wstring& path) {
    try {
        fs::path p(path);
        if (!p.parent_path().empty()) fs::create_directories(p.parent_path());
        std::ofstream f(p, std::ios::binary | std::ios::trunc);
        const unsigned char bom[] = { 0xFF, 0xFE }; // UTF-16LE BOM
        f.write(reinterpret_cast<const char*>(bom), 2);
    }
    catch (...) {
        // 失敗しても処理を継続する
    }
}

static bool IsUnicodeIniFile(const std::wstring& path) {
    try {
        std::ifstream f(fs::path(path), std::ios::binary);
        unsigned char b0 = 0, b1 = 0;
        f.read(reinterpret_cast<char*>(&b0), 1);
        f.read(reinterpret_cast<char*>(&b1), 1);
        return (b0 == 0xFF && b1 == 0xFE) || (b0 == 0xFE && b1 == 0xFF);
    }
    catch (...) {
        return false;
    }
}

static fs::path MakeUniqueBackupPath(const fs::path& p, const std::wstring& suffix) {
    fs::path out = p;
    out += suffix;
    if (!fs::exists(out)) return out;
    for (int i = 1; i < 1000; ++i) {
        fs::path c = p;
        c += suffix;
        c += L"." + std::to_wstring(i);
        if (!fs::exists(c)) return c;
    }
    return out;
}

static std::wstring IniReadStrFrom(const std::wstring& iniPath, const wchar_t* sec, const wchar_t* key, const std::wstring& defv) {
    wchar_t buf[2048]{};
    GetPrivateProfileStringW(sec, key, defv.c_str(), buf, 2048, iniPath.c_str());
    return buf;
}
static int IniReadIntFrom(const std::wstring& iniPath, const wchar_t* sec, const wchar_t* key, int defv) {
    return GetPrivateProfileIntW(sec, key, defv, iniPath.c_str());
}

static std::wstring GetLocalAppDataExcelTodayDir() {
    PWSTR p = nullptr;
    if (FAILED(SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, nullptr, &p)) || !p) return L"";
    std::wstring base(p);
    CoTaskMemFree(p);
    fs::path dir = fs::path(base) / L"ExcelToday";
    std::error_code ec;
    fs::create_directories(dir, ec);
    return dir.wstring();
}

static void InitPaths() {
    const std::wstring exeDir = GetExeDir(); // サーバー上でも利用可能
    const std::wstring localDir = GetLocalAppDataExcelTodayDir();

    if (!localDir.empty()) {
        g_iniPath = (fs::path(localDir) / L"settings.ini").wstring();
        g_lastExcludeFile = (fs::path(localDir) / L"exclude.txt").wstring();
        g_lastCsvFile = (fs::path(localDir) / L"results.csv").wstring();
        g_lastNameExcludeFile = (fs::path(localDir) / L"name_exclude.txt").wstring();
    }
    else {
        // フォールバック（権限問題が再発し得る）
        g_iniPath = exeDir + L"\\settings.ini";
        g_lastExcludeFile = exeDir + L"\\exclude.txt";
        g_lastCsvFile = exeDir + L"\\results.csv";
        g_lastNameExcludeFile = exeDir + L"\\name_exclude.txt";
    }
}

static void EnsureUnicodeIniWithMigration() {
    if (g_iniPath.empty()) return;

    const std::wstring exeDir = GetExeDir();
    const std::wstring legacyIni = exeDir + L"\\settings.ini"; // 旧設定: 実行ファイルフォルダー（読み取りのみ）

    // 1) ローカルINIが無ければ Unicode（BOM付き）で作る
    if (!fs::exists(fs::path(g_iniPath))) {
        CreateUtf16LeBomFile(g_iniPath);

        // 旧INIがあれば読み取りだけして移行（旧INIはリネーム/削除しない）
        if (fs::exists(fs::path(legacyIni))) {
            auto root = IniReadStrFrom(legacyIni, L"Main", L"Root", L"");
            auto excl = IniReadStrFrom(legacyIni, L"Main", L"ExcludeFile", g_lastExcludeFile);
            auto csv = IniReadStrFrom(legacyIni, L"Main", L"CsvFile", g_lastCsvFile);
            auto nxf = IniReadStrFrom(legacyIni, L"Main", L"NameExcludeFile", g_lastNameExcludeFile);

            if (ContainsQuestionMark(excl)) excl = g_lastExcludeFile;
            if (ContainsQuestionMark(csv))  csv = g_lastCsvFile;

            WritePrivateProfileStringW(L"Main", L"Root", root.c_str(), g_iniPath.c_str());
            WritePrivateProfileStringW(L"Main", L"ExcludeFile", excl.c_str(), g_iniPath.c_str());
            WritePrivateProfileStringW(L"Main", L"CsvFile", csv.c_str(), g_iniPath.c_str());
            WritePrivateProfileStringW(L"Main", L"NameExcludeFile", nxf.c_str(), g_iniPath.c_str());
            {
                wchar_t b[64];
                _snwprintf_s(b, _TRUNCATE, L"%d", IniReadIntFrom(legacyIni, L"Main", L"Mode", 0));
                WritePrivateProfileStringW(L"Main", L"Mode", b, g_iniPath.c_str());
                WritePrivateProfileStringW(L"Main", L"Days", IniReadStrFrom(legacyIni, L"Main", L"Days", L"3").c_str(), g_iniPath.c_str());
                _snwprintf_s(b, _TRUNCATE, L"%d", IniReadIntFrom(legacyIni, L"Main", L"TimeBase", 0));
                WritePrivateProfileStringW(L"Main", L"TimeBase", b, g_iniPath.c_str());
            }

            // 拡張子設定
            const wchar_t* extKeys[] = { L"xls", L"xlsx", L"xlsm", L"xlsb", L"xltx", L"xltm" };
            for (auto k : extKeys) {
                wchar_t b[16];
                _snwprintf_s(b, _TRUNCATE, L"%d", IniReadIntFrom(legacyIni, L"Ext", k, 0));
                WritePrivateProfileStringW(L"Ext", k, b, g_iniPath.c_str());
            }

            // フィルター
            const wchar_t* fKeys[] = { L"EnableFolderExclude", L"EnableNameExclude", L"NameIncludeExt" };
            for (auto k : fKeys) {
                wchar_t b[16];
                _snwprintf_s(b, _TRUNCATE, L"%d", IniReadIntFrom(legacyIni, L"Filter", k, 1));
                WritePrivateProfileStringW(L"Filter", k, b, g_iniPath.c_str());
            }

            // 表示設定
            {
                wchar_t b[16];
                _snwprintf_s(b, _TRUNCATE, L"%d", IniReadIntFrom(legacyIni, L"View", L"SortCol", 0));
                WritePrivateProfileStringW(L"View", L"SortCol", b, g_iniPath.c_str());
                _snwprintf_s(b, _TRUNCATE, L"%d", IniReadIntFrom(legacyIni, L"View", L"SortAsc", 0));
                WritePrivateProfileStringW(L"View", L"SortAsc", b, g_iniPath.c_str());
            }

            // ファイル名フィルター設定
            int n = IniReadIntFrom(legacyIni, L"NameFilter", L"Count", 0);
            if (n < 0) n = 0;
            if (n > 200) n = 200;
            {
                wchar_t b[16];
                _snwprintf_s(b, _TRUNCATE, L"%d", n);
                WritePrivateProfileStringW(L"NameFilter", L"Count", b, g_iniPath.c_str());
            }
            for (int i = 0; i < n; ++i) {
                std::wstring key = L"Item" + std::to_wstring(i);
                auto v = IniReadStrFrom(legacyIni, L"NameFilter", key.c_str(), L"");
                WritePrivateProfileStringW(L"NameFilter", key.c_str(), v.c_str(), g_iniPath.c_str());
            }
        }
        return;
    }

    // 2) ローカルINIが存在するが Unicode でない → ローカルでバックアップして作り直し
    if (!IsUnicodeIniFile(g_iniPath)) {
        fs::path iniP(g_iniPath);
        fs::path bakP = MakeUniqueBackupPath(iniP, L".ansi.bak");
        std::error_code ec;
        fs::rename(iniP, bakP, ec);
        if (ec) {
            ec.clear();
            fs::copy_file(iniP, bakP, fs::copy_options::overwrite_existing, ec);
            ec.clear();
            fs::remove(iniP, ec);
        }

        CreateUtf16LeBomFile(g_iniPath);

        // バックアップから可能な範囲で移行
        auto root = IniReadStrFrom(bakP.wstring(), L"Main", L"Root", L"");
        auto excl = IniReadStrFrom(bakP.wstring(), L"Main", L"ExcludeFile", g_lastExcludeFile);
        auto csv = IniReadStrFrom(bakP.wstring(), L"Main", L"CsvFile", g_lastCsvFile);
        if (ContainsQuestionMark(excl)) excl = g_lastExcludeFile;
        if (ContainsQuestionMark(csv))  csv = g_lastCsvFile;

        WritePrivateProfileStringW(L"Main", L"Root", root.c_str(), g_iniPath.c_str());
        WritePrivateProfileStringW(L"Main", L"ExcludeFile", excl.c_str(), g_iniPath.c_str());
        WritePrivateProfileStringW(L"Main", L"CsvFile", csv.c_str(), g_iniPath.c_str());
    }
}


static void IniWriteStr(const wchar_t* sec, const wchar_t* key, const std::wstring& val) {
    WritePrivateProfileStringW(sec, key, val.c_str(), g_iniPath.c_str());
}
static void IniWriteInt(const wchar_t* sec, const wchar_t* key, int val) {
    wchar_t buf[64];
    _snwprintf_s(buf, _TRUNCATE, L"%d", val);
    WritePrivateProfileStringW(sec, key, buf, g_iniPath.c_str());
}
static std::wstring IniReadStr(const wchar_t* sec, const wchar_t* key, const std::wstring& defv) {
    wchar_t buf[2048]{};
    GetPrivateProfileStringW(sec, key, defv.c_str(), buf, 2048, g_iniPath.c_str());
    return buf;
}
static int IniReadInt(const wchar_t* sec, const wchar_t* key, int defv) {
    return GetPrivateProfileIntW(sec, key, defv, g_iniPath.c_str());
}

// -------------------- UIの有効/無効切り替え --------------------
static void UpdateUiEnableStates() {
    bool folderOn = IsChecked(g_chkEnableFolderExcl);
    EnableWindow(g_listExcludes, folderOn);
    EnableWindow(g_btnAddExclFolder, folderOn);
    EnableWindow(g_btnRemoveExcl, folderOn);
    EnableWindow(g_btnExclUp, folderOn);
    EnableWindow(g_btnExclDown, folderOn);
    EnableWindow(g_btnLoadExcl, folderOn);
    EnableWindow(g_btnSaveExcl, folderOn);
    EnableWindow(g_editExclPattern, folderOn);
    EnableWindow(g_btnAddPattern, folderOn);

    bool nameOn = IsChecked(g_chkEnableNameExcl);
    EnableWindow(g_chkNameIncludeExt, !g_searching);
    InvalidateRect(g_chkNameIncludeExt, nullptr, TRUE);
    SetTargetExtensionSectionEnabled(!g_searching && UseTargetExtensionFilter());
    EnableWindow(g_editFNamePattern, nameOn);
    EnableWindow(g_btnAddFName, nameOn);
    EnableWindow(g_btnRemoveFName, nameOn);
    EnableWindow(g_btnFNameUp, nameOn);
    EnableWindow(g_btnFNameDown, nameOn);
    EnableWindow(g_listFName, nameOn);
    EnableWindow(g_btnLoadFNameExcl, nameOn);
    EnableWindow(g_btnSaveFNameExcl, nameOn);

    int mode = GetMode(); // 0:今日 1:過去N日 2:期間指定

    bool useDays = (mode == 1);
    bool useCal = (mode == 2);
    if (g_leftTab != 0 || !useCal) CloseModernCalendarPopup();

    if (g_leftTab != 0) {
        // 除外タブでは、検索期間系は非表示/無効（誤って復活表示しないようにする）
        EnableWindow(g_editDays, FALSE);
        EnableWindow(g_staticDays, FALSE);
        ShowWindow(g_staticFrom, SW_HIDE);
        ShowWindow(g_staticTo, SW_HIDE);
        ShowWindow(g_dtpFrom, SW_HIDE);
        ShowWindow(g_dtpTo, SW_HIDE);
    }
    else {
        EnableWindow(g_editDays, useDays);
        EnableWindow(g_staticDays, useDays);

        EnableWindow(g_dtpFrom, useCal);
        EnableWindow(g_dtpTo, useCal);
        ShowWindow(g_staticFrom, useCal ? SW_SHOW : SW_HIDE);
        ShowWindow(g_staticTo, useCal ? SW_SHOW : SW_HIDE);
        ShowWindow(g_dtpFrom, useCal ? SW_SHOW : SW_HIDE);
        ShowWindow(g_dtpTo, useCal ? SW_SHOW : SW_HIDE);
    }
}


static void EnsureThemeBrushes() {
    if (!g_hBrushAppBg) g_hBrushAppBg = CreateSolidBrush(Theme::AppBg);
    if (!g_hBrushCardBg) g_hBrushCardBg = CreateSolidBrush(Theme::CardBg);
    if (!g_hBrushEditBg) g_hBrushEditBg = CreateSolidBrush(Theme::CardBg);
}

static void ApplyModernControlTheme(HWND hwnd) {
    if (hwnd) {
        SetWindowTheme(hwnd, L"Explorer", nullptr);
    }
}

static void EnableModernOwnerDrawButton(HWND hwnd) {
    if (!hwnd) return;
    LONG_PTR style = GetWindowLongPtrW(hwnd, GWL_STYLE);
    style = (style & ~BS_TYPEMASK) | BS_OWNERDRAW;
    SetWindowLongPtrW(hwnd, GWL_STYLE, style);
    SetWindowPos(hwnd, nullptr, 0, 0, 0, 0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
}

static void ApplyModernListBox(HWND hwnd, int itemHeight = 30) {
    if (!hwnd) return;
    ApplyModernControlTheme(hwnd);
    SendMessageW(hwnd, LB_SETITEMHEIGHT, 0, itemHeight);
}

static void DrawRoundedRect(HDC hdc, const RECT& rc, COLORREF fill, COLORREF border, int radius);

static void UpdateHotControl(HWND& hotControl, HWND hwnd, bool hot) {
    HWND oldHot = hotControl;
    hotControl = hot ? hwnd : (hotControl == hwnd ? nullptr : hotControl);
    if (oldHot && oldHot != hotControl) InvalidateRect(oldHot, nullptr, TRUE);
    if (hotControl && oldHot != hotControl) InvalidateRect(hotControl, nullptr, TRUE);
}

static void StartHoverTracking(HWND hwnd) {
    TRACKMOUSEEVENT tme{};
    tme.cbSize = sizeof(tme);
    tme.dwFlags = TME_LEAVE;
    tme.hwndTrack = hwnd;
    TrackMouseEvent(&tme);
}

static LRESULT CALLBACK ModernButtonHoverProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
    UINT_PTR, DWORD_PTR) {
    switch (msg) {
    case WM_MOUSEMOVE:
        if (g_hotButton != hwnd) {
            UpdateHotControl(g_hotButton, hwnd, true);
        }
        StartHoverTracking(hwnd);
        break;
    case WM_MOUSELEAVE:
        UpdateHotControl(g_hotButton, hwnd, false);
        break;
    case WM_ENABLE:
        if (!IsWindowEnabled(hwnd)) {
            UpdateHotControl(g_hotButton, hwnd, false);
        }
        InvalidateRect(hwnd, nullptr, TRUE);
        break;
    case WM_NCDESTROY:
        UpdateHotControl(g_hotButton, hwnd, false);
        RemoveWindowSubclass(hwnd, ModernButtonHoverProc, 1);
        break;
    }
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

static void EnableButtonHoverHighlight(HWND hwnd) {
    if (!hwnd) return;
    SetWindowSubclass(hwnd, ModernButtonHoverProc, 1, 0);
}

static void DrawModernCheckBoxFace(HWND hwnd, HDC hdc) {
    if (!hwnd || !hdc) return;

    RECT rc{};
    GetClientRect(hwnd, &rc);
    const bool disabled = !IsWindowEnabled(hwnd);
    const bool hot = (g_hotCheckBox == hwnd);
    const bool focused = GetFocus() == hwnd;
    const bool checked = (SendMessageW(hwnd, BM_GETCHECK, 0, 0) == BST_CHECKED);

    HBRUSH bg = CreateSolidBrush(Theme::CardBg);
    FillRect(hdc, &rc, bg);
    DeleteObject(bg);

    constexpr int boxSize = 18;
    RECT boxRc{};
    boxRc.left = 2;
    boxRc.top = rc.top + max(0, ((rc.bottom - rc.top) - boxSize) / 2);
    boxRc.right = boxRc.left + boxSize;
    boxRc.bottom = boxRc.top + boxSize;

    COLORREF boxFill = disabled ? Theme::DisabledBg : (checked ? Theme::Primary : (hot ? Theme::NeutralButtonHot : Theme::CardBg));
    COLORREF boxBorder = disabled ? Theme::Border : (checked || focused ? Theme::Primary : (hot ? RGB(147, 197, 253) : Theme::Border));
    DrawRoundedRect(hdc, boxRc, boxFill, boxBorder, 6);

    if (checked) {
        HPEN checkPen = CreatePen(PS_SOLID, 2, disabled ? Theme::DisabledText : RGB(255, 255, 255));
        HGDIOBJ oldPen = SelectObject(hdc, checkPen);
        MoveToEx(hdc, boxRc.left + 4, boxRc.top + 9, nullptr);
        LineTo(hdc, boxRc.left + 8, boxRc.bottom - 5);
        LineTo(hdc, boxRc.right - 4, boxRc.top + 5);
        SelectObject(hdc, oldPen);
        DeleteObject(checkPen);
    }

    wchar_t text[128]{};
    GetWindowTextW(hwnd, text, static_cast<int>(std::size(text)));
    RECT textRc = rc;
    textRc.left = boxRc.right + 8;
    textRc.right -= 2;

    HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
    HGDIOBJ oldFont = font ? SelectObject(hdc, font) : nullptr;
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, disabled ? Theme::DisabledText : Theme::Text);
    DrawTextW(hdc, text, -1, &textRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
    if (oldFont) SelectObject(hdc, oldFont);

    if (focused) {
        RECT focusRc = rc;
        focusRc.left = 0;
        InflateRect(&focusRc, -1, -1);
        HPEN pen = CreatePen(PS_DOT, 1, Theme::Primary);
        HGDIOBJ oldPen = SelectObject(hdc, pen);
        HGDIOBJ oldBrush = SelectObject(hdc, GetStockObject(NULL_BRUSH));
        RoundRect(hdc, focusRc.left, focusRc.top, focusRc.right, focusRc.bottom, 8, 8);
        SelectObject(hdc, oldBrush);
        SelectObject(hdc, oldPen);
        DeleteObject(pen);
    }
}

static void RedrawModernCheckBoxNow(HWND hwnd) {
    if (!hwnd || !IsWindowVisible(hwnd)) return;
    RedrawWindow(hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_NOERASE);
}

static void ToggleModernCheckBox(HWND hwnd) {
    const LRESULT current = SendMessageW(hwnd, BM_GETCHECK, 0, 0);
    SendMessageW(hwnd, BM_SETCHECK, current == BST_CHECKED ? BST_UNCHECKED : BST_CHECKED, 0);
    RedrawModernCheckBoxNow(hwnd);
    HWND parent = GetParent(hwnd);
    if (parent) {
        SendMessageW(parent, WM_COMMAND, MAKEWPARAM(GetDlgCtrlID(hwnd), BN_CLICKED), reinterpret_cast<LPARAM>(hwnd));
    }
}

static LRESULT CallCheckBoxDefaultWithoutNativePaint(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    SendMessageW(hwnd, WM_SETREDRAW, FALSE, 0);
    LRESULT result = DefSubclassProc(hwnd, msg, wParam, lParam);
    SendMessageW(hwnd, WM_SETREDRAW, TRUE, 0);
    RedrawModernCheckBoxNow(hwnd);
    return result;
}

static LRESULT CALLBACK ModernCheckBoxProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
    UINT_PTR, DWORD_PTR) {
    switch (msg) {
    case WM_PAINT:
    {
        PAINTSTRUCT ps{};
        HDC hdc = BeginPaint(hwnd, &ps);
        DrawModernCheckBoxFace(hwnd, hdc);
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_PRINTCLIENT:
        DrawModernCheckBoxFace(hwnd, reinterpret_cast<HDC>(wParam));
        return 0;
    case WM_ERASEBKGND:
        return 1;
    case WM_MOUSEMOVE:
        if (g_hotCheckBox != hwnd) {
            UpdateHotControl(g_hotCheckBox, hwnd, true);
        }
        StartHoverTracking(hwnd);
        return 0;
    case WM_MOUSELEAVE:
        UpdateHotControl(g_hotCheckBox, hwnd, false);
        return 0;
    case WM_LBUTTONDOWN:
    case WM_LBUTTONDBLCLK:
        if (IsWindowEnabled(hwnd)) {
            SetFocus(hwnd);
            SetCapture(hwnd);
        }
        RedrawModernCheckBoxNow(hwnd);
        return 0;
    case WM_LBUTTONUP:
        if (GetCapture() == hwnd) {
            ReleaseCapture();
            POINT pt{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
            RECT rc{};
            GetClientRect(hwnd, &rc);
            if (PtInRect(&rc, pt) && IsWindowEnabled(hwnd)) {
                ToggleModernCheckBox(hwnd);
            }
        }
        RedrawModernCheckBoxNow(hwnd);
        return 0;
    case WM_KEYDOWN:
        if (wParam == VK_SPACE) return 0;
        break;
    case WM_KEYUP:
        if (wParam == VK_SPACE && IsWindowEnabled(hwnd)) {
            ToggleModernCheckBox(hwnd);
            RedrawModernCheckBoxNow(hwnd);
            return 0;
        }
        break;
    case WM_ENABLE:
        if (!IsWindowEnabled(hwnd)) {
            UpdateHotControl(g_hotCheckBox, hwnd, false);
        }
        RedrawModernCheckBoxNow(hwnd);
        return 0;
    case WM_SETFOCUS:
    case WM_KILLFOCUS:
        RedrawModernCheckBoxNow(hwnd);
        return 0;
    case WM_SETTEXT:
    case BM_SETCHECK:
        return CallCheckBoxDefaultWithoutNativePaint(hwnd, msg, wParam, lParam);
    case WM_NCDESTROY:
        if (GetCapture() == hwnd) ReleaseCapture();
        UpdateHotControl(g_hotCheckBox, hwnd, false);
        RemoveWindowSubclass(hwnd, ModernCheckBoxProc, 1);
        break;
    }
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

static void ApplyModernCheckBox(HWND hwnd) {
    if (!hwnd) return;
    SetWindowSubclass(hwnd, ModernCheckBoxProc, 1, 0);
    InvalidateRect(hwnd, nullptr, TRUE);
}

static void DrawModernComboBoxFace(HWND hwnd, HDC hdc) {
    if (!hwnd || !hdc) return;

    RECT rc{};
    GetClientRect(hwnd, &rc);
    const bool disabled = !IsWindowEnabled(hwnd);
    const bool focused = GetFocus() == hwnd;
    const bool hot = (g_hotCombo == hwnd);
    COLORREF fill = disabled ? Theme::DisabledBg : (hot ? Theme::NeutralButtonHot : Theme::CardBg);
    COLORREF border = focused ? Theme::Primary : (hot ? RGB(147, 197, 253) : Theme::Border);
    COLORREF textColor = disabled ? Theme::DisabledText : Theme::Text;

    HBRUSH clear = CreateSolidBrush(Theme::CardBg);
    FillRect(hdc, &rc, clear);
    DeleteObject(clear);

    RECT boxRc = rc;
    InflateRect(&boxRc, -1, -1);
    DrawRoundedRect(hdc, boxRc, fill, border, 8);

    int sel = static_cast<int>(SendMessageW(hwnd, CB_GETCURSEL, 0, 0));
    std::wstring text;
    if (sel != CB_ERR) {
        int len = static_cast<int>(SendMessageW(hwnd, CB_GETLBTEXTLEN, sel, 0));
        if (len >= 0) {
            text.resize(static_cast<size_t>(len));
            SendMessageW(hwnd, CB_GETLBTEXT, sel, reinterpret_cast<LPARAM>(text.data()));
        }
    }

    RECT textRc = boxRc;
    textRc.left += 10;
    textRc.right -= 28;
    HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
    HGDIOBJ oldFont = font ? SelectObject(hdc, font) : nullptr;
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, textColor);
    DrawTextW(hdc, text.c_str(), -1, &textRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
    if (oldFont) SelectObject(hdc, oldFont);

    if (!disabled) {
        POINT pts[3]{};
        int cx = boxRc.right - 15;
        int cy = boxRc.top + ((boxRc.bottom - boxRc.top) / 2) + 1;
        pts[0] = { cx - 4, cy - 2 };
        pts[1] = { cx + 4, cy - 2 };
        pts[2] = { cx, cy + 3 };
        HBRUSH arrow = CreateSolidBrush(Theme::MutedText);
        HGDIOBJ oldBrush = SelectObject(hdc, arrow);
        HPEN pen = CreatePen(PS_SOLID, 1, Theme::MutedText);
        HGDIOBJ oldPen = SelectObject(hdc, pen);
        Polygon(hdc, pts, 3);
        SelectObject(hdc, oldPen);
        SelectObject(hdc, oldBrush);
        DeleteObject(pen);
        DeleteObject(arrow);
    }
}

static LRESULT CALLBACK ModernComboBoxProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
    UINT_PTR, DWORD_PTR) {
    switch (msg) {
    case WM_PAINT:
    {
        PAINTSTRUCT ps{};
        HDC hdc = BeginPaint(hwnd, &ps);
        DrawModernComboBoxFace(hwnd, hdc);
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_PRINTCLIENT:
        DrawModernComboBoxFace(hwnd, reinterpret_cast<HDC>(wParam));
        return 0;
    case WM_ERASEBKGND:
        return 1;
    case WM_MOUSEMOVE:
        if (g_hotCombo != hwnd) {
            UpdateHotControl(g_hotCombo, hwnd, true);
        }
        StartHoverTracking(hwnd);
        break;
    case WM_MOUSELEAVE:
        UpdateHotControl(g_hotCombo, hwnd, false);
        break;
    case WM_SETFOCUS:
    case WM_KILLFOCUS:
    case WM_ENABLE:
    case WM_LBUTTONUP:
        InvalidateRect(hwnd, nullptr, TRUE);
        break;
    case CB_SETCURSEL:
    case CB_RESETCONTENT:
    case CB_ADDSTRING:
    case CB_INSERTSTRING:
    case CB_DELETESTRING:
    {
        LRESULT result = DefSubclassProc(hwnd, msg, wParam, lParam);
        InvalidateRect(hwnd, nullptr, TRUE);
        return result;
    }
    case WM_NCDESTROY:
        UpdateHotControl(g_hotCombo, hwnd, false);
        RemoveWindowSubclass(hwnd, ModernComboBoxProc, 1);
        break;
    }
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

static std::wstring FormatDateText(const SYSTEMTIME& st) {
    if (st.wYear == 0) return L"";
    wchar_t buf[32]{};
    _snwprintf_s(buf, _TRUNCATE, L"%04u/%02u/%02u", st.wYear, st.wMonth, st.wDay);
    return buf;
}

static SYSTEMTIME& DateForPicker(HWND hwnd) {
    return hwnd == g_dtpTo ? g_dateTo : g_dateFrom;
}

static void UpdateModernDatePickerText(HWND hwnd) {
    if (!hwnd) return;
    SetWindowTextW(hwnd, FormatDateText(DateForPicker(hwnd)).c_str());
    InvalidateRect(hwnd, nullptr, TRUE);
}

static bool IsLeapYear(WORD year) {
    return (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0);
}

static int DaysInMonth(WORD year, WORD month) {
    static constexpr int days[] = { 31,28,31,30,31,30,31,31,30,31,30,31 };
    if (month == 2) return IsLeapYear(year) ? 29 : 28;
    if (month < 1 || month > 12) return 31;
    return days[month - 1];
}

static int DayOfWeek(WORD year, WORD month, WORD day) {
    SYSTEMTIME st{};
    st.wYear = year;
    st.wMonth = month;
    st.wDay = day;
    FILETIME ft{};
    if (!SystemTimeToFileTime(&st, &ft)) return 0;
    SYSTEMTIME back{};
    FileTimeToSystemTime(&ft, &back);
    return back.wDayOfWeek;
}

static void OffsetCalendarMonth(int delta) {
    int month = static_cast<int>(g_calendarMonth.wMonth) + delta;
    int year = static_cast<int>(g_calendarMonth.wYear);
    while (month < 1) { month += 12; --year; }
    while (month > 12) { month -= 12; ++year; }
    g_calendarMonth.wYear = static_cast<WORD>(max(1601, min(9999, year)));
    g_calendarMonth.wMonth = static_cast<WORD>(month);
    g_calendarMonth.wDay = 1;
}

static void CloseModernCalendarPopup() {
    if (g_calendarPopup && IsWindow(g_calendarPopup)) {
        DestroyWindow(g_calendarPopup);
    }
    g_calendarPopup = nullptr;
    g_calendarTarget = nullptr;
}

static void DrawModernDatePickerFace(HWND hwnd, HDC hdc) {
    if (!hwnd || !hdc) return;

    RECT rc{};
    GetClientRect(hwnd, &rc);
    const bool disabled = !IsWindowEnabled(hwnd);
    const bool focused = GetFocus() == hwnd;
    const bool hot = (g_hotCombo == hwnd);
    COLORREF fill = disabled ? Theme::DisabledBg : (hot ? Theme::NeutralButtonHot : Theme::CardBg);
    COLORREF border = focused ? Theme::Primary : (hot ? RGB(147, 197, 253) : Theme::Border);
    COLORREF textColor = disabled ? Theme::DisabledText : Theme::Text;

    HBRUSH clear = CreateSolidBrush(Theme::CardBg);
    FillRect(hdc, &rc, clear);
    DeleteObject(clear);

    RECT boxRc = rc;
    boxRc.right = max(boxRc.left, boxRc.right - 1);
    boxRc.bottom = max(boxRc.top, boxRc.bottom - 1);
    DrawRoundedRect(hdc, boxRc, fill, border, 8);

    wchar_t text[64]{};
    GetWindowTextW(hwnd, text, static_cast<int>(std::size(text)));

    RECT textRc = boxRc;
    textRc.left += 10;
    textRc.right -= 34;
    HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
    HGDIOBJ oldFont = font ? SelectObject(hdc, font) : nullptr;
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, textColor);
    DrawTextW(hdc, text, -1, &textRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
    if (oldFont) SelectObject(hdc, oldFont);

    RECT iconRc = boxRc;
    iconRc.left = max(iconRc.left, iconRc.right - 30);
    iconRc.top += 6;
    iconRc.right -= 9;
    iconRc.bottom -= 6;

    COLORREF iconColor = disabled ? Theme::DisabledText : Theme::MutedText;
    HPEN pen = CreatePen(PS_SOLID, 1, iconColor);
    HGDIOBJ oldPen = SelectObject(hdc, pen);
    HGDIOBJ oldBrush = SelectObject(hdc, GetStockObject(NULL_BRUSH));
    RoundRect(hdc, iconRc.left, iconRc.top + 2, iconRc.right, iconRc.bottom, 4, 4);
    MoveToEx(hdc, iconRc.left, iconRc.top + 7, nullptr);
    LineTo(hdc, iconRc.right, iconRc.top + 7);
    MoveToEx(hdc, iconRc.left + 4, iconRc.top, nullptr);
    LineTo(hdc, iconRc.left + 4, iconRc.top + 4);
    MoveToEx(hdc, iconRc.right - 4, iconRc.top, nullptr);
    LineTo(hdc, iconRc.right - 4, iconRc.top + 4);
    SelectObject(hdc, oldBrush);
    SelectObject(hdc, oldPen);
    DeleteObject(pen);
}

static void ShowModernCalendarPopup(HWND target);

static LRESULT CALLBACK ModernDatePickerProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
    UINT_PTR, DWORD_PTR) {
    switch (msg) {
    case WM_PAINT:
    {
        PAINTSTRUCT ps{};
        HDC hdc = BeginPaint(hwnd, &ps);
        DrawModernDatePickerFace(hwnd, hdc);
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_PRINTCLIENT:
        DrawModernDatePickerFace(hwnd, reinterpret_cast<HDC>(wParam));
        return 0;
    case WM_NCPAINT:
    case WM_ERASEBKGND:
        return 1;
    case WM_MOUSEMOVE:
        if (g_hotCombo != hwnd) {
            UpdateHotControl(g_hotCombo, hwnd, true);
        }
        StartHoverTracking(hwnd);
        break;
    case WM_MOUSELEAVE:
        UpdateHotControl(g_hotCombo, hwnd, false);
        break;
    case WM_SETFOCUS:
    case WM_KILLFOCUS:
    case WM_ENABLE:
        InvalidateRect(hwnd, nullptr, TRUE);
        break;
    case WM_LBUTTONDOWN:
        SetFocus(hwnd);
        ShowModernCalendarPopup(hwnd);
        return 0;
    case WM_KEYDOWN:
        if (wParam == VK_SPACE || wParam == VK_RETURN || wParam == VK_DOWN) {
            ShowModernCalendarPopup(hwnd);
            return 0;
        }
        break;
    case WM_NCDESTROY:
        UpdateHotControl(g_hotCombo, hwnd, false);
        RemoveWindowSubclass(hwnd, ModernDatePickerProc, 1);
        break;
    }
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

static void PaintModernCalendarPopup(HWND hwnd, HDC hdc) {
    RECT rc{};
    GetClientRect(hwnd, &rc);
    FillRect(hdc, &rc, g_hBrushCardBg ? g_hBrushCardBg : reinterpret_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));

    RECT card = rc;
    card.right -= 1;
    card.bottom -= 1;
    DrawRoundedRect(hdc, card, Theme::CardBg, Theme::Border, 14);

    HFONT oldFont = reinterpret_cast<HFONT>(SelectObject(hdc, g_hFontUi ? g_hFontUi : GetStockObject(DEFAULT_GUI_FONT)));
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, Theme::Text);

    RECT titleRc{ 48, 12, rc.right - 48, 42 };
    wchar_t title[32]{};
    _snwprintf_s(title, _TRUNCATE, L"%04u年 %02u月", g_calendarMonth.wYear, g_calendarMonth.wMonth);
    if (g_hFontUiBold) SelectObject(hdc, g_hFontUiBold);
    DrawTextW(hdc, title, -1, &titleRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    SelectObject(hdc, g_hFontUi ? g_hFontUi : GetStockObject(DEFAULT_GUI_FONT));

    RECT prevRc{ 14, 12, 42, 40 };
    RECT nextRc{ rc.right - 42, 12, rc.right - 14, 40 };
    DrawRoundedRect(hdc, prevRc, Theme::NeutralButton, Theme::Border, 8);
    DrawRoundedRect(hdc, nextRc, Theme::NeutralButton, Theme::Border, 8);
    SetTextColor(hdc, Theme::MutedText);
    DrawTextW(hdc, L"‹", -1, &prevRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    DrawTextW(hdc, L"›", -1, &nextRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

    static const wchar_t* week[] = { L"日", L"月", L"火", L"水", L"木", L"金", L"土" };
    const int left = 14;
    const int top = 56;
    const int cellW = 40;
    const int cellH = 32;
    SetTextColor(hdc, Theme::MutedText);
    for (int i = 0; i < 7; ++i) {
        RECT wrc{ left + i * cellW, top, left + (i + 1) * cellW, top + 24 };
        DrawTextW(hdc, week[i], -1, &wrc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    }

    SYSTEMTIME selected = DateForPicker(g_calendarTarget);
    SYSTEMTIME today{};
    GetLocalTime(&today);
    int firstDow = DayOfWeek(g_calendarMonth.wYear, g_calendarMonth.wMonth, 1);
    int dim = DaysInMonth(g_calendarMonth.wYear, g_calendarMonth.wMonth);
    int gridTop = top + 30;

    HPEN gridPen = CreatePen(PS_SOLID, 1, RGB(229, 234, 240));
    HGDIOBJ oldGridPen = SelectObject(hdc, gridPen);
    for (int row = 0; row <= 6; ++row) {
        int y = gridTop + row * cellH;
        MoveToEx(hdc, left, y, nullptr);
        LineTo(hdc, left + cellW * 7, y);
    }
    for (int col = 0; col <= 7; ++col) {
        int x = left + col * cellW;
        MoveToEx(hdc, x, gridTop, nullptr);
        LineTo(hdc, x, gridTop + cellH * 6);
    }
    SelectObject(hdc, oldGridPen);
    DeleteObject(gridPen);

    for (int day = 1; day <= dim; ++day) {
        int pos = firstDow + day - 1;
        int row = pos / 7;
        int col = pos % 7;
        RECT drc{ left + col * cellW + 3, gridTop + row * cellH + 3,
            left + (col + 1) * cellW - 3, gridTop + (row + 1) * cellH - 3 };
        bool isSelected = selected.wYear == g_calendarMonth.wYear && selected.wMonth == g_calendarMonth.wMonth && selected.wDay == day;
        bool isToday = today.wYear == g_calendarMonth.wYear && today.wMonth == g_calendarMonth.wMonth && today.wDay == day;
        if (isSelected) {
            DrawRoundedRect(hdc, drc, Theme::Primary, Theme::PrimaryHot, 8);
            SetTextColor(hdc, RGB(255, 255, 255));
        }
        else if (isToday) {
            DrawRoundedRect(hdc, drc, RGB(239, 246, 255), RGB(147, 197, 253), 8);
            SetTextColor(hdc, Theme::SelectedText);
        }
        else {
            SetTextColor(hdc, (col == 0) ? Theme::Danger : Theme::Text);
        }
        wchar_t d[4]{};
        _snwprintf_s(d, _TRUNCATE, L"%d", day);
        DrawTextW(hdc, d, -1, &drc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    }

    if (oldFont) SelectObject(hdc, oldFont);
}

static int HitTestModernCalendarDay(int x, int y) {
    const int left = 14;
    const int top = 56 + 30;
    const int cellW = 40;
    const int cellH = 32;
    if (x < left || x >= left + cellW * 7 || y < top) return 0;
    int col = (x - left) / cellW;
    int row = (y - top) / cellH;
    if (row < 0 || row >= 6) return 0;
    int firstDow = DayOfWeek(g_calendarMonth.wYear, g_calendarMonth.wMonth, 1);
    int day = row * 7 + col - firstDow + 1;
    int dim = DaysInMonth(g_calendarMonth.wYear, g_calendarMonth.wMonth);
    return (day >= 1 && day <= dim) ? day : 0;
}

static LRESULT CALLBACK ModernCalendarPopupProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_PAINT:
    {
        PAINTSTRUCT ps{};
        HDC hdc = BeginPaint(hwnd, &ps);
        PaintModernCalendarPopup(hwnd, hdc);
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_ERASEBKGND:
        return 1;
    case WM_LBUTTONDOWN:
    {
        int x = GET_X_LPARAM(lParam);
        int y = GET_Y_LPARAM(lParam);
        RECT rc{};
        GetClientRect(hwnd, &rc);
        RECT prevRc{ 14, 12, 42, 40 };
        RECT nextRc{ rc.right - 42, 12, rc.right - 14, 40 };
        POINT pt{ x, y };
        if (PtInRect(&prevRc, pt)) {
            OffsetCalendarMonth(-1);
            InvalidateRect(hwnd, nullptr, TRUE);
            return 0;
        }
        if (PtInRect(&nextRc, pt)) {
            OffsetCalendarMonth(1);
            InvalidateRect(hwnd, nullptr, TRUE);
            return 0;
        }
        int day = HitTestModernCalendarDay(x, y);
        if (day > 0 && g_calendarTarget) {
            SYSTEMTIME& targetDate = DateForPicker(g_calendarTarget);
            targetDate.wYear = g_calendarMonth.wYear;
            targetDate.wMonth = g_calendarMonth.wMonth;
            targetDate.wDay = static_cast<WORD>(day);
            UpdateModernDatePickerText(g_calendarTarget);
            CloseModernCalendarPopup();
            return 0;
        }
        return 0;
    }
    case WM_KEYDOWN:
        if (wParam == VK_ESCAPE) {
            CloseModernCalendarPopup();
            return 0;
        }
        break;
    case WM_ACTIVATE:
        if (LOWORD(wParam) == WA_INACTIVE) {
            CloseModernCalendarPopup();
            return 0;
        }
        break;
    case WM_DESTROY:
        if (g_calendarPopup == hwnd) g_calendarPopup = nullptr;
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static void ShowModernCalendarPopup(HWND target) {
    if (!target) return;
    if (g_calendarPopup && IsWindow(g_calendarPopup) && g_calendarTarget == target) {
        CloseModernCalendarPopup();
        return;
    }
    CloseModernCalendarPopup();
    g_calendarTarget = target;
    g_calendarMonth = DateForPicker(target);
    if (g_calendarMonth.wYear == 0) GetLocalTime(&g_calendarMonth);
    g_calendarMonth.wDay = 1;

    static bool registered = false;
    if (!registered) {
        WNDCLASSW wc{};
        wc.lpfnWndProc = ModernCalendarPopupProc;
        wc.hInstance = g_hInst;
        wc.lpszClassName = L"WorkSupportModernCalendarPopup";
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
        wc.hbrBackground = nullptr;
        RegisterClassW(&wc);
        registered = true;
    }

    RECT tr{};
    GetWindowRect(target, &tr);
    const int w = 308;
    const int h = 284;
    g_calendarPopup = CreateWindowExW(WS_EX_TOOLWINDOW | WS_EX_TOPMOST,
        L"WorkSupportModernCalendarPopup", L"", WS_POPUP,
        tr.left, tr.bottom + 6, w, h, g_hwndMain, nullptr, g_hInst, nullptr);
    if (g_calendarPopup) {
        ShowWindow(g_calendarPopup, SW_SHOW);
        SetFocus(g_calendarPopup);
    }
}

static void ApplyModernComboBox(HWND hwnd, int selectionHeight = 30, int itemHeight = 30) {
    if (!hwnd) return;
    ApplyModernControlTheme(hwnd);
    SendMessageW(hwnd, CB_SETITEMHEIGHT, (WPARAM)-1, selectionHeight);
    SendMessageW(hwnd, CB_SETITEMHEIGHT, 0, itemHeight);
    SendMessageW(hwnd, CB_SETMINVISIBLE, 8, 0);
    SetWindowSubclass(hwnd, ModernComboBoxProc, 1, 0);
    InvalidateRect(hwnd, nullptr, TRUE);
}

static void ApplyModernResultsListView(HWND hwnd) {
    if (!hwnd) return;
    ApplyModernControlTheme(hwnd);
    ListView_SetBkColor(hwnd, Theme::CardBg);
    ListView_SetTextBkColor(hwnd, CLR_NONE);
    ListView_SetTextColor(hwnd, Theme::Text);
    ListView_SetExtendedListViewStyle(hwnd,
        LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_HEADERDRAGDROP | LVS_EX_LABELTIP);
    HWND header = ListView_GetHeader(hwnd);
    if (header) {
        ApplyModernControlTheme(header);
        SendMessageW(header, HDM_SETBITMAPMARGIN, 8, 0);
    }

    if (!g_hResultsRowImageList) {
        g_hResultsRowImageList = ImageList_Create(1, 34, ILC_COLOR32, 1, 1);
        if (g_hResultsRowImageList) {
            HBITMAP bmp = CreateBitmap(1, 34, 1, 32, nullptr);
            ImageList_Add(g_hResultsRowImageList, bmp, nullptr);
            DeleteObject(bmp);
        }
    }
    if (g_hResultsRowImageList) {
        ListView_SetImageList(hwnd, g_hResultsRowImageList, LVSIL_SMALL);
    }
}

static void ApplyModernDatePickerTheme(HWND hwnd) {
    if (!hwnd) return;
    SetWindowSubclass(hwnd, ModernDatePickerProc, 1, 0);
    UpdateModernDatePickerText(hwnd);
}

static bool IsPrimaryButtonId(int id) {
    return id == IDC_BTN_SEARCH;
}

static bool IsDangerButtonId(int id) {
    return id == IDC_BTN_STOP;
}

static bool IsCsvButtonId(int id) {
    return id == IDC_BTN_EXPORT_CSV;
}

static void DrawRoundedRect(HDC hdc, const RECT& rc, COLORREF fill, COLORREF border, int radius) {
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

static bool DrawModernButton(const DRAWITEMSTRUCT* dis) {
    if (!dis || dis->CtlType != ODT_BUTTON) return false;

    const int id = GetDlgCtrlID(dis->hwndItem);
    const bool disabled = (dis->itemState & ODS_DISABLED) != 0;
    const bool pressed = (dis->itemState & ODS_SELECTED) != 0;
    const bool hot = ((dis->itemState & ODS_HOTLIGHT) != 0) || (g_hotButton == dis->hwndItem);
    const bool focused = (dis->itemState & ODS_FOCUS) != 0;
    const bool primary = IsPrimaryButtonId(id);
    const bool danger = IsDangerButtonId(id);
    const bool csv = IsCsvButtonId(id);

    COLORREF fill = Theme::NeutralButton;
    COLORREF border = Theme::Border;
    COLORREF textColor = Theme::Text;

    if (disabled) {
        fill = Theme::DisabledBg;
        border = Theme::Border;
        textColor = Theme::DisabledText;
    }
    else if (primary) {
        fill = (pressed || hot) ? Theme::PrimaryHot : Theme::Primary;
        border = Theme::PrimaryHot;
        textColor = RGB(255, 255, 255);
    }
    else if (danger) {
        fill = pressed ? Theme::DangerHot : (hot ? RGB(254, 242, 242) : Theme::NeutralButton);
        border = hot || pressed ? Theme::Danger : RGB(252, 165, 165);
        textColor = pressed ? RGB(255, 255, 255) : Theme::Danger;
    }
    else if (csv) {
        fill = pressed ? Theme::CsvButtonPressed : (hot ? Theme::CsvButtonHot : Theme::CsvButton);
        border = Theme::CsvButtonBorder;
        textColor = RGB(22, 101, 52);
    }
    else {
        fill = pressed ? Theme::NeutralButtonPressed : (hot ? Theme::NeutralButtonHot : Theme::NeutralButton);
        border = hot || focused ? RGB(147, 197, 253) : Theme::Border;
    }

    RECT rc = dis->rcItem;
    InflateRect(&rc, -1, -1);
    DrawRoundedRect(dis->hDC, rc, fill, border, 10);

    wchar_t text[128]{};
    GetWindowTextW(dis->hwndItem, text, (int)std::size(text));
    HFONT font = reinterpret_cast<HFONT>(SendMessageW(dis->hwndItem, WM_GETFONT, 0, 0));
    HGDIOBJ oldFont = font ? SelectObject(dis->hDC, font) : nullptr;
    SetBkMode(dis->hDC, TRANSPARENT);
    SetTextColor(dis->hDC, textColor);
    DrawTextW(dis->hDC, text, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
    if (oldFont) SelectObject(dis->hDC, oldFont);
    return true;
}

static bool DrawModernListBox(const DRAWITEMSTRUCT* dis) {
    if (!dis || dis->CtlType != ODT_LISTBOX) return false;
    if (dis->itemID == static_cast<UINT>(-1)) return true;

    wchar_t text[2048]{};
    SendMessageW(dis->hwndItem, LB_GETTEXT, dis->itemID, reinterpret_cast<LPARAM>(text));

    const bool selected = (dis->itemState & ODS_SELECTED) != 0;
    const bool focused = (dis->itemState & ODS_FOCUS) != 0;
    const bool disabled = (dis->itemState & ODS_DISABLED) != 0;
    RECT rc = dis->rcItem;

    COLORREF bg = (dis->itemID % 2 == 0) ? Theme::CardBg : Theme::ListAltBg;
    COLORREF fg = Theme::Text;
    if (selected) {
        bg = Theme::SelectedBg;
        fg = Theme::SelectedText;
    }
    if (disabled || wcsncmp(text, L"[無効]", 4) == 0) {
        fg = Theme::DisabledText;
    }

    HBRUSH brush = CreateSolidBrush(bg);
    FillRect(dis->hDC, &rc, brush);
    DeleteObject(brush);

    RECT textRc = rc;
    textRc.left += 12;
    textRc.right -= 10;
    HFONT font = reinterpret_cast<HFONT>(SendMessageW(dis->hwndItem, WM_GETFONT, 0, 0));
    HGDIOBJ oldFont = font ? SelectObject(dis->hDC, font) : nullptr;
    SetBkMode(dis->hDC, TRANSPARENT);
    SetTextColor(dis->hDC, fg);
    DrawTextW(dis->hDC, text, -1, &textRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
    if (oldFont) SelectObject(dis->hDC, oldFont);

    if (focused) {
        HPEN pen = CreatePen(PS_SOLID, 1, Theme::Primary);
        HGDIOBJ oldPen = SelectObject(dis->hDC, pen);
        HGDIOBJ oldBrush = SelectObject(dis->hDC, GetStockObject(NULL_BRUSH));
        RECT focusRc = rc;
        InflateRect(&focusRc, -2, -2);
        RoundRect(dis->hDC, focusRc.left, focusRc.top, focusRc.right, focusRc.bottom, 8, 8);
        SelectObject(dis->hDC, oldBrush);
        SelectObject(dis->hDC, oldPen);
        DeleteObject(pen);
    }
    return true;
}

static LRESULT HandleResultsHeaderCustomDraw(LPNMCUSTOMDRAW cd) {
    if (!cd) return CDRF_DODEFAULT;
    switch (cd->dwDrawStage) {
    case CDDS_PREPAINT:
        return CDRF_NOTIFYITEMDRAW;
    case CDDS_ITEMPREPAINT:
    {
        RECT rc = cd->rc;
        HBRUSH bg = CreateSolidBrush(Theme::ListAltBg);
        FillRect(cd->hdc, &rc, bg);
        DeleteObject(bg);

        HPEN line = CreatePen(PS_SOLID, 1, Theme::Border);
        HGDIOBJ oldPen = SelectObject(cd->hdc, line);
        MoveToEx(cd->hdc, rc.left, rc.bottom - 1, nullptr);
        LineTo(cd->hdc, rc.right, rc.bottom - 1);
        SelectObject(cd->hdc, oldPen);
        DeleteObject(line);

        wchar_t text[128]{};
        HDITEMW item{};
        item.mask = HDI_TEXT | HDI_FORMAT;
        item.pszText = text;
        item.cchTextMax = static_cast<int>(std::size(text));
        Header_GetItem(cd->hdr.hwndFrom, static_cast<int>(cd->dwItemSpec), &item);

        RECT textRc = rc;
        textRc.left += 12;
        textRc.right -= 12;
        HFONT font = g_hFontUiBold ? g_hFontUiBold : reinterpret_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
        HGDIOBJ oldFont = SelectObject(cd->hdc, font);
        SetBkMode(cd->hdc, TRANSPARENT);
        SetTextColor(cd->hdc, Theme::MutedText);
        UINT format = DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS;
        format |= (item.fmt & HDF_RIGHT) ? DT_RIGHT : DT_LEFT;
        DrawTextW(cd->hdc, text, -1, &textRc, format);
        SelectObject(cd->hdc, oldFont);
        return CDRF_SKIPDEFAULT;
    }
    default:
        return CDRF_DODEFAULT;
    }
}

static bool DrawModernComboBox(const DRAWITEMSTRUCT* dis) {
    if (!dis || dis->CtlType != ODT_COMBOBOX) return false;

    RECT rc = dis->rcItem;
    const bool editField = (dis->itemState & ODS_COMBOBOXEDIT) != 0;
    const bool selected = (dis->itemState & ODS_SELECTED) != 0;
    const bool focused = (dis->itemState & ODS_FOCUS) != 0;
    const bool disabled = (dis->itemState & ODS_DISABLED) != 0 || !IsWindowEnabled(dis->hwndItem);
    if (editField) {
        DrawModernComboBoxFace(dis->hwndItem, dis->hDC);
        return true;
    }

    std::wstring text;
    UINT itemId = dis->itemID;
    if (itemId == static_cast<UINT>(-1)) {
        int sel = static_cast<int>(SendMessageW(dis->hwndItem, CB_GETCURSEL, 0, 0));
        if (sel != CB_ERR) itemId = static_cast<UINT>(sel);
    }
    if (itemId != static_cast<UINT>(-1)) {
        int len = static_cast<int>(SendMessageW(dis->hwndItem, CB_GETLBTEXTLEN, itemId, 0));
        if (len >= 0) {
            text.resize(static_cast<size_t>(len));
            SendMessageW(dis->hwndItem, CB_GETLBTEXT, itemId, reinterpret_cast<LPARAM>(text.data()));
        }
    }

    COLORREF bg = Theme::CardBg;
    COLORREF fg = disabled ? Theme::DisabledText : Theme::Text;
    COLORREF border = focused ? Theme::Primary : Theme::Border;
    if (!editField) {
        bg = (itemId != static_cast<UINT>(-1) && (itemId % 2) != 0) ? Theme::ListAltBg : Theme::CardBg;
        if (selected) {
            bg = Theme::SelectedBg;
            fg = Theme::SelectedText;
        }
    }

    if (editField) {
        RECT fillRc = rc;
        InflateRect(&fillRc, -1, -1);
        DrawRoundedRect(dis->hDC, fillRc, disabled ? Theme::DisabledBg : bg, border, 8);
    }
    else {
        HBRUSH brush = CreateSolidBrush(bg);
        FillRect(dis->hDC, &rc, brush);
        DeleteObject(brush);
    }

    RECT textRc = rc;
    textRc.left += editField ? 10 : 12;
    textRc.right -= editField ? 22 : 10;
    HFONT font = reinterpret_cast<HFONT>(SendMessageW(dis->hwndItem, WM_GETFONT, 0, 0));
    HGDIOBJ oldFont = font ? SelectObject(dis->hDC, font) : nullptr;
    SetBkMode(dis->hDC, TRANSPARENT);
    SetTextColor(dis->hDC, fg);
    DrawTextW(dis->hDC, text.c_str(), -1, &textRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
    if (oldFont) SelectObject(dis->hDC, oldFont);

    if (editField && !disabled) {
        POINT pts[3]{};
        int cx = rc.right - 13;
        int cy = rc.top + ((rc.bottom - rc.top) / 2) + 1;
        pts[0] = { cx - 4, cy - 2 };
        pts[1] = { cx + 4, cy - 2 };
        pts[2] = { cx, cy + 3 };
        HBRUSH arrow = CreateSolidBrush(Theme::MutedText);
        HGDIOBJ oldBrush = SelectObject(dis->hDC, arrow);
        HPEN pen = CreatePen(PS_SOLID, 1, Theme::MutedText);
        HGDIOBJ oldPen = SelectObject(dis->hDC, pen);
        Polygon(dis->hDC, pts, 3);
        SelectObject(dis->hDC, oldPen);
        SelectObject(dis->hDC, oldBrush);
        DeleteObject(pen);
        DeleteObject(arrow);
    }

    return true;
}

static LRESULT HandleResultsCustomDraw(LPNMLVCUSTOMDRAW cd) {
    if (!cd) return CDRF_DODEFAULT;
    switch (cd->nmcd.dwDrawStage) {
    case CDDS_PREPAINT:
        return CDRF_NOTIFYITEMDRAW;
    case CDDS_ITEMPREPAINT:
        return CDRF_NOTIFYSUBITEMDRAW;
    case CDDS_SUBITEM | CDDS_ITEMPREPAINT:
    {
        const bool selected = (cd->nmcd.uItemState & CDIS_SELECTED) != 0;
        const bool hot = (cd->nmcd.uItemState & CDIS_HOT) != 0;
        const bool odd = (cd->nmcd.dwItemSpec % 2) != 0;
        cd->clrTextBk = selected ? Theme::SelectedBg : (hot ? RGB(239, 246, 255) : (odd ? Theme::ListAltBg : Theme::CardBg));
        cd->clrText = selected ? Theme::SelectedText : (cd->iSubItem == 3 ? Theme::MutedText : Theme::Text);
        if (cd->iSubItem == 1) cd->clrText = selected ? Theme::SelectedText : RGB(55, 65, 81);
        if (cd->iSubItem == 2 && g_hFontUiBold) {
            SelectObject(cd->nmcd.hdc, g_hFontUiBold);
            return CDRF_NEWFONT;
        }
        return CDRF_DODEFAULT;
    }
    default:
        return CDRF_DODEFAULT;
    }
}


static bool DrawModernTab(const DRAWITEMSTRUCT* dis) {
    if (!dis || dis->CtlType != ODT_TAB || dis->hwndItem != g_tabLeft) return false;

    wchar_t text[64]{};
    TCITEMW item{};
    item.mask = TCIF_TEXT;
    item.pszText = text;
    item.cchTextMax = (int)std::size(text);
    TabCtrl_GetItem(g_tabLeft, (int)dis->itemID, &item);

    const bool selected = ((int)dis->itemID == TabCtrl_GetCurSel(g_tabLeft));
    RECT rc = dis->rcItem;
    InflateRect(&rc, -3, -2);
    DrawRoundedRect(dis->hDC, rc, selected ? Theme::CardBg : RGB(238, 242, 247),
        selected ? Theme::Border : RGB(229, 234, 240), 10);

    if (selected) {
        HBRUSH accent = CreateSolidBrush(Theme::Primary);
        RECT underline{ rc.left + 14, rc.bottom - 4, rc.right - 14, rc.bottom - 2 };
        FillRect(dis->hDC, &underline, accent);
        DeleteObject(accent);
    }

    HFONT font = reinterpret_cast<HFONT>(SendMessageW(g_tabLeft, WM_GETFONT, 0, 0));
    HGDIOBJ oldFont = font ? SelectObject(dis->hDC, font) : nullptr;
    SetBkMode(dis->hDC, TRANSPARENT);
    SetTextColor(dis->hDC, selected ? Theme::Text : Theme::MutedText);
    DrawTextW(dis->hDC, text, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
    if (oldFont) SelectObject(dis->hDC, oldFont);
    return true;
}

static void DrawCard(HDC hdc, const RECT& rc) {
    HBRUSH card = CreateSolidBrush(Theme::CardBg);
    HPEN border = CreatePen(PS_SOLID, 1, Theme::Border);
    HGDIOBJ oldBrush = SelectObject(hdc, card);
    HGDIOBJ oldPen = SelectObject(hdc, border);
    RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, 12, 12);
    SelectObject(hdc, oldPen);
    SelectObject(hdc, oldBrush);
    DeleteObject(border);
    DeleteObject(card);
}

static void PaintSearchBackground(HWND hwnd, HDC hdc) {
    EnsureThemeBrushes();
    RECT rc{};
    GetClientRect(hwnd, &rc);
    FillRect(hdc, &rc, g_hBrushAppBg);

    const int padding = 12;
    const int statusH = 22;
    int winW = rc.right - rc.left;
    int winH = rc.bottom - rc.top;
    int minRightW = 440;
    int leftW = 560;
    if (winW < leftW + minRightW + padding * 3) {
        leftW = max(320, winW - (minRightW + padding * 3));
    }
    int w = leftW + padding * 2;
    int rightX = w + padding;
    RECT leftCard{ 6, 6, rightX - 8, max(80, winH - statusH - 6) };
    RECT rightCard{ rightX - 2, 6, max(rightX + 260, winW - 6), max(80, winH - statusH - 6) };
    DrawCard(hdc, leftCard);
    DrawCard(hdc, rightCard);
}

static void DoLayout(HWND hwnd); // 前方宣言

static RECT GetLeftPaneInvalidRect(HWND hwnd) {
    RECT rc{};
    GetClientRect(hwnd, &rc);
    const int padding = 12;
    int winW = rc.right - rc.left;
    int minRightW = 440;
    int leftW = 560;
    if (winW < leftW + minRightW + padding * 3) {
        leftW = max(320, winW - (minRightW + padding * 3));
    }
    int rightX = leftW + padding * 3;
    return RECT{ 0, 0, min(rightX, rc.right), rc.bottom };
}

static void InvalidateLeftPane(HWND hwnd) {
    if (!hwnd) return;
    RECT leftPane = GetLeftPaneInvalidRect(hwnd);
    InvalidateRect(hwnd, &leftPane, TRUE);
}

static void RedrawListBoxIfVisible(HWND hwnd) {
    if (!hwnd || !IsWindowVisible(hwnd)) return;
    InvalidateRect(hwnd, nullptr, TRUE);
    UpdateWindow(hwnd);
}

// -------------------- 左タブ（検索 / 除外） --------------------
static void ApplyLeftTabVisibility() {
    bool isSearch = (g_leftTab == 0);
    int swSearch = isSearch ? SW_SHOW : SW_HIDE;
    int swExcl = isSearch ? SW_HIDE : SW_SHOW;

    // 検索タブのコントロール
    ShowWindow(g_staticRoot, swSearch);
    ShowWindow(g_listRoots, swSearch);
    ShowWindow(g_btnBrowseRoot, swSearch);
    ShowWindow(g_btnRootRemove, swSearch);
    ShowWindow(g_btnRootUp, swSearch);
    ShowWindow(g_btnRootDown, swSearch);
    ShowWindow(g_btnRootToggle, swSearch);
    ShowWindow(g_staticRootsHint, swSearch);

    ShowWindow(g_staticMode, swSearch);
    ShowWindow(g_cmbMode, swSearch);
    ShowWindow(g_staticDays, swSearch);
    ShowWindow(g_editDays, swSearch);
    ShowWindow(g_staticTimeBase, swSearch);
    ShowWindow(g_cmbTimeBase, swSearch);
    ShowWindow(g_staticPreset, swSearch);
    ShowWindow(g_cmbPreset, swSearch);
    ShowWindow(g_btnPresetSave, swSearch);
    ShowWindow(g_btnPresetLoad, swSearch);
    ShowWindow(g_btnPresetDelete, swSearch);

    ShowWindow(g_staticFrom, swSearch);
    ShowWindow(g_staticTo, swSearch);
    ShowWindow(g_dtpFrom, swSearch);
    ShowWindow(g_dtpTo, swSearch);
    ShowWindow(g_chkNameIncludeExt, swSearch);

    HWND hExtGrp = (g_hwndMain ? GetDlgItem(g_hwndMain, IDC_GRP_EXT) : nullptr);
    if (hExtGrp) ShowWindow(hExtGrp, swSearch);
    ShowWindow(g_chkXls, swSearch);
    ShowWindow(g_chkXlsx, swSearch);
    ShowWindow(g_chkXlsm, swSearch);
    ShowWindow(g_chkXlsb, swSearch);
    ShowWindow(g_chkXltx, swSearch);
    ShowWindow(g_chkXltm, swSearch);

    // 除外タブのコントロール
    ShowWindow(g_frameFolderExcl, swExcl);
    ShowWindow(g_frameNameExcl, swExcl);

    ShowWindow(g_staticExclFolder, swExcl);
    ShowWindow(g_chkEnableFolderExcl, swExcl);
    ShowWindow(g_listExcludes, swExcl);
    ShowWindow(g_btnAddExclFolder, swExcl);
    ShowWindow(g_btnRemoveExcl, swExcl);
    ShowWindow(g_btnExclUp, swExcl);
    ShowWindow(g_btnExclDown, swExcl);
    ShowWindow(g_btnLoadExcl, swExcl);
    ShowWindow(g_btnSaveExcl, swExcl);
    ShowWindow(g_staticExclPattern, swExcl);
    ShowWindow(g_editExclPattern, swExcl);
    ShowWindow(g_btnAddPattern, swExcl);

    ShowWindow(g_staticExclName, swExcl);
    ShowWindow(g_chkEnableNameExcl, swExcl);
    ShowWindow(g_editFNamePattern, swExcl);
    ShowWindow(g_btnAddFName, swExcl);
    ShowWindow(g_btnRemoveFName, swExcl);
    ShowWindow(g_btnFNameUp, swExcl);
    ShowWindow(g_btnFNameDown, swExcl);
    ShowWindow(g_listFName, swExcl);
    ShowWindow(g_btnLoadFNameExcl, swExcl);
    ShowWindow(g_btnSaveFNameExcl, swExcl);

    // 表示中のコントロールの有効/無効状態を整合させる
    UpdateUiEnableStates();
}

static void SetLeftTab(int tab, bool saveIni = true) {
    tab = max(0, min(1, tab));
    g_leftTab = tab;
    if (g_tabLeft) TabCtrl_SetCurSel(g_tabLeft, g_leftTab);
    ApplyLeftTabVisibility();
    if (g_hwndMain) {
        DoLayout(g_hwndMain);
        InvalidateLeftPane(g_hwndMain);
    }
    if (saveIni) {
        IniWriteInt(L"View", L"LeftTab", g_leftTab);
    }
}

// -------------------- 設定の読み込み/保存 --------------------
static void RefreshPresetCombo(const std::wstring& selectName = L"");
static void SaveSearchPreset();
static void LoadSearchPreset();
static void DeleteSearchPreset();

static void LoadSettings() {
    std::vector<RootEntry> entries;

    int rootsCount = IniReadInt(L"Main", L"RootsCount", 0);
    if (rootsCount > 0) {
        for (int i = 0; i < rootsCount; ++i) {
            std::wstring keyPath = L"RootPath" + std::to_wstring(i);
            std::wstring keyEnabled = L"RootEnabled" + std::to_wstring(i);

            auto p = Trim(IniReadStr(L"Main", keyPath.c_str(), L""));
            bool enabled = IniReadInt(L"Main", keyEnabled.c_str(), 1) != 0;

            if (!p.empty()) {
                entries.push_back({ p, enabled });
            }
        }
    }
    else {
        auto rootsIni = IniReadStr(L"Main", L"Roots", L"");
        std::vector<std::wstring> rootsList;
        if (!rootsIni.empty()) {
            rootsList = SplitRootsText(rootsIni);
        }
        else {
            auto one = Trim(IniReadStr(L"Main", L"Root", L""));
            if (!one.empty()) rootsList.push_back(one);
        }

        for (const auto& r : rootsList) {
            auto t = Trim(r);
            if (!t.empty()) entries.push_back({ t, true });
        }
    }

    SetRootEntriesToListBox(entries);

    g_lastExcludeFile = IniReadStr(L"Main", L"ExcludeFile", g_lastExcludeFile.empty() ? (GetExeDir() + L"\\exclude.txt") : g_lastExcludeFile);
    g_lastCsvFile = IniReadStr(L"Main", L"CsvFile", g_lastCsvFile.empty() ? (GetExeDir() + L"\\results.csv") : g_lastCsvFile);
    g_lastNameExcludeFile = IniReadStr(L"Main", L"NameExcludeFile", g_lastNameExcludeFile.empty() ? (GetExeDir() + L"\\name_exclude.txt") : g_lastNameExcludeFile);

    int mode = IniReadInt(L"Main", L"Mode", 0);
    SendMessageW(g_cmbMode, CB_SETCURSEL, (WPARAM)max(0, min(2, mode)), 0);
    SetWindowTextWStr(g_editDays, IniReadStr(L"Main", L"Days", L"3"));

    int tb = IniReadInt(L"Main", L"TimeBase", 0);
    SendMessageW(g_cmbTimeBase, CB_SETCURSEL, (WPARAM)max(0, min(2, tb)), 0);

    SetChecked(g_chkXls, IniReadInt(L"Ext", L"xls", 1) != 0);
    SetChecked(g_chkXlsx, IniReadInt(L"Ext", L"xlsx", 1) != 0);
    SetChecked(g_chkXlsm, IniReadInt(L"Ext", L"xlsm", 1) != 0);
    SetChecked(g_chkXlsb, IniReadInt(L"Ext", L"xlsb", 1) != 0);
    SetChecked(g_chkXltx, IniReadInt(L"Ext", L"xltx", 1) != 0);
    SetChecked(g_chkXltm, IniReadInt(L"Ext", L"xltm", 1) != 0);

    SetChecked(g_chkEnableFolderExcl, IniReadInt(L"Filter", L"EnableFolderExclude", 1) != 0);
    SetChecked(g_chkEnableNameExcl, IniReadInt(L"Filter", L"EnableNameExclude", 1) != 0);
    SetChecked(g_chkNameIncludeExt, IniReadInt(L"Filter", L"NameIncludeExt", 1) != 0);

    g_sortCol = IniReadInt(L"View", L"SortCol", 0);
    g_sortAsc = IniReadInt(L"View", L"SortAsc", 0) != 0;
    g_leftTab = max(0, min(1, IniReadInt(L"View", L"LeftTab", 0)));
    if (g_tabLeft) TabCtrl_SetCurSel(g_tabLeft, g_leftTab);
    if (g_sortCol < 0 || g_sortCol > 3) g_sortCol = 0;

    if (!g_lastExcludeFile.empty() && fs::exists(fs::path(g_lastExcludeFile))) {
        auto rootsForBase = GetRootsFromListBox();
        std::wstring rootStr = rootsForBase.empty() ? L"" : rootsForBase[0];
        fs::path base = rootStr.empty() ? fs::path(GetExeDir()) : NormalizePath(fs::path(rootStr));
        LoadExcludesFromFile(g_lastExcludeFile, base);
    }

    g_fileNamePatterns.clear();
    int n = IniReadInt(L"NameFilter", L"Count", 0);
    if (n <= 0) {
        g_fileNamePatterns.push_back(L"~$");
    }
    else {
        for (int i = 0; i < n; ++i) {
            std::wstring key = L"Item" + std::to_wstring(i);
            auto v = Trim(IniReadStr(L"NameFilter", key.c_str(), L""));
            if (!v.empty()) g_fileNamePatterns.push_back(v);
        }
        if (g_fileNamePatterns.empty()) g_fileNamePatterns.push_back(L"~$");
    }
    RefreshFileNameListBox();
    RebuildFileNameExcludeCache();

    RefreshPresetCombo();
    ApplyLeftTabVisibility();
    UpdateUiEnableStates();
    SetListViewTimeHeader(GetTimeBase());
}

static void SaveSettings() {
    auto rootEntries = GetRootEntriesFromListBox();
    auto enabledRootItems = GetEnabledRootsFromListBox();

    // 新方式: 1件ずつ保存
    int oldCount = IniReadInt(L"Main", L"RootsCount", 0);
    int newCount = (int)rootEntries.size();
    IniWriteInt(L"Main", L"RootsCount", newCount);

    int maxCount = (oldCount > newCount) ? oldCount : newCount;
    if (maxCount < 200) maxCount = 200; // 以前の空キー掃除も兼ねる

    for (int i = 0; i < maxCount; ++i) {
        std::wstring keyPath = L"RootPath" + std::to_wstring(i);
        std::wstring keyEnabled = L"RootEnabled" + std::to_wstring(i);

        if (i < newCount) {
            IniWriteStr(L"Main", keyPath.c_str(), rootEntries[(size_t)i].path);
            IniWriteInt(L"Main", keyEnabled.c_str(), rootEntries[(size_t)i].enabled ? 1 : 0);
        }
        else {
            WritePrivateProfileStringW(L"Main", keyPath.c_str(), nullptr, g_iniPath.c_str());
            WritePrivateProfileStringW(L"Main", keyEnabled.c_str(), nullptr, g_iniPath.c_str());
        }
    }

    // 旧方式の残骸を消す
    WritePrivateProfileStringW(L"Main", L"RootsEx", nullptr, g_iniPath.c_str());

    // 互換用: 有効なものだけ従来キーにも保存
    IniWriteStr(L"Main", L"Roots", JoinRootsForIni(enabledRootItems));
    IniWriteStr(L"Main", L"Root", enabledRootItems.empty() ? L"" : enabledRootItems[0]);

    IniWriteStr(L"Main", L"ExcludeFile", g_lastExcludeFile);
    IniWriteStr(L"Main", L"CsvFile", g_lastCsvFile);
    IniWriteStr(L"Main", L"NameExcludeFile", g_lastNameExcludeFile);
    IniWriteInt(L"Main", L"Mode", GetMode());
    IniWriteStr(L"Main", L"Days", GetWindowTextWStr(g_editDays));
    IniWriteInt(L"Main", L"TimeBase", (GetTimeBase() == TimeBase::Creation) ? 1 : (GetTimeBase() == TimeBase::Either ? 2 : 0));

    IniWriteInt(L"Ext", L"xls", IsChecked(g_chkXls) ? 1 : 0);
    IniWriteInt(L"Ext", L"xlsx", IsChecked(g_chkXlsx) ? 1 : 0);
    IniWriteInt(L"Ext", L"xlsm", IsChecked(g_chkXlsm) ? 1 : 0);
    IniWriteInt(L"Ext", L"xlsb", IsChecked(g_chkXlsb) ? 1 : 0);
    IniWriteInt(L"Ext", L"xltx", IsChecked(g_chkXltx) ? 1 : 0);
    IniWriteInt(L"Ext", L"xltm", IsChecked(g_chkXltm) ? 1 : 0);

    IniWriteInt(L"Filter", L"EnableFolderExclude", IsChecked(g_chkEnableFolderExcl) ? 1 : 0);
    IniWriteInt(L"Filter", L"EnableNameExclude", IsChecked(g_chkEnableNameExcl) ? 1 : 0);
    IniWriteInt(L"Filter", L"NameIncludeExt", IsChecked(g_chkNameIncludeExt) ? 1 : 0);

    IniWriteInt(L"View", L"SortCol", g_sortCol);
    IniWriteInt(L"View", L"SortAsc", g_sortAsc ? 1 : 0);
    IniWriteInt(L"View", L"LeftTab", g_leftTab);

    int oldNameCount = IniReadInt(L"NameFilter", L"Count", 0);
    int newNameCount = (int)g_fileNamePatterns.size();
    IniWriteInt(L"NameFilter", L"Count", newNameCount);
    int maxNameCount = (oldNameCount > newNameCount) ? oldNameCount : newNameCount;
    for (int i = 0; i < maxNameCount; ++i) {
        std::wstring key = L"Item" + std::to_wstring(i);
        if (i < newNameCount) IniWriteStr(L"NameFilter", key.c_str(), g_fileNamePatterns[(size_t)i]);
        else WritePrivateProfileStringW(L"NameFilter", key.c_str(), nullptr, g_iniPath.c_str());
    }

    if (g_lastExcludeFile.empty()) {
        auto ld = GetLocalAppDataExcelTodayDir();
        if (!ld.empty()) g_lastExcludeFile = (fs::path(ld) / L"exclude.txt").wstring();
        else g_lastExcludeFile = GetExeDir() + L"\\exclude.txt";
    }
    SaveExcludesToFile(g_lastExcludeFile);
}

static std::wstring GetPresetSection(const std::wstring& name) {
    std::wstring safe = Trim(name);
    for (auto& ch : safe) {
        if (ch == L'[' || ch == L']' || ch == L'\r' || ch == L'\n') ch = L'_';
    }
    return L"SearchPreset_" + safe;
}

static std::wstring GetCurrentPresetName() {
    if (!g_cmbPreset) return L"";
    return Trim(GetWindowTextWStr(g_cmbPreset));
}

static std::vector<std::wstring> GetPresetNames() {
    std::vector<std::wstring> names;
    int n = IniReadInt(L"SearchPresets", L"Count", 0);
    for (int i = 0; i < n; ++i) {
        std::wstring key = L"Name" + std::to_wstring(i);
        auto name = Trim(IniReadStr(L"SearchPresets", key.c_str(), L""));
        if (!name.empty()) names.push_back(name);
    }
    names.erase(std::remove_if(names.begin(), names.end(), [](const std::wstring& s) { return Trim(s).empty(); }), names.end());
    std::sort(names.begin(), names.end(), [](const auto& a, const auto& b) { return ToLower(a) < ToLower(b); });
    names.erase(std::unique(names.begin(), names.end(), [](const auto& a, const auto& b) { return ToLower(a) == ToLower(b); }), names.end());
    return names;
}

static void SavePresetNames(const std::vector<std::wstring>& names) {
    int oldCount = IniReadInt(L"SearchPresets", L"Count", 0);
    IniWriteInt(L"SearchPresets", L"Count", (int)names.size());
    int maxCount = max(oldCount, (int)names.size());
    for (int i = 0; i < maxCount; ++i) {
        std::wstring key = L"Name" + std::to_wstring(i);
        if (i < (int)names.size()) IniWriteStr(L"SearchPresets", key.c_str(), names[(size_t)i]);
        else WritePrivateProfileStringW(L"SearchPresets", key.c_str(), nullptr, g_iniPath.c_str());
    }
}

static void RefreshPresetCombo(const std::wstring& selectName) {
    if (!g_cmbPreset) return;
    std::wstring keep = selectName.empty() ? GetCurrentPresetName() : selectName;
    ComboBox_ResetContent(g_cmbPreset);
    for (const auto& name : GetPresetNames()) {
        ComboBox_AddString(g_cmbPreset, name.c_str());
    }
    if (!keep.empty()) SetWindowTextW(g_cmbPreset, keep.c_str());
}

static void SaveSearchPreset() {
    std::wstring name = GetCurrentPresetName();
    if (name.empty()) {
        MessageBoxW(g_hwndMain, L"プリセット名を入力してください。", L"検索条件プリセット", MB_OK | MB_ICONINFORMATION);
        return;
    }

    auto names = GetPresetNames();
    if (std::none_of(names.begin(), names.end(), [&](const auto& s) { return ToLower(s) == ToLower(name); })) {
        names.push_back(name);
        std::sort(names.begin(), names.end(), [](const auto& a, const auto& b) { return ToLower(a) < ToLower(b); });
        SavePresetNames(names);
    }

    const std::wstring sec = GetPresetSection(name);
    auto rootEntries = GetRootEntriesFromListBox();
    IniWriteInt(sec.c_str(), L"RootsCount", (int)rootEntries.size());
    for (int i = 0; i < (int)rootEntries.size(); ++i) {
        IniWriteStr(sec.c_str(), (L"RootPath" + std::to_wstring(i)).c_str(), rootEntries[(size_t)i].path);
        IniWriteInt(sec.c_str(), (L"RootEnabled" + std::to_wstring(i)).c_str(), rootEntries[(size_t)i].enabled ? 1 : 0);
    }
    IniWriteInt(sec.c_str(), L"Mode", GetMode());
    IniWriteStr(sec.c_str(), L"Days", GetWindowTextWStr(g_editDays));
    IniWriteInt(sec.c_str(), L"TimeBase", (GetTimeBase() == TimeBase::Creation) ? 1 : (GetTimeBase() == TimeBase::Either ? 2 : 0));
    IniWriteInt(sec.c_str(), L"xls", IsChecked(g_chkXls) ? 1 : 0);
    IniWriteInt(sec.c_str(), L"xlsx", IsChecked(g_chkXlsx) ? 1 : 0);
    IniWriteInt(sec.c_str(), L"xlsm", IsChecked(g_chkXlsm) ? 1 : 0);
    IniWriteInt(sec.c_str(), L"xlsb", IsChecked(g_chkXlsb) ? 1 : 0);
    IniWriteInt(sec.c_str(), L"xltx", IsChecked(g_chkXltx) ? 1 : 0);
    IniWriteInt(sec.c_str(), L"xltm", IsChecked(g_chkXltm) ? 1 : 0);
    IniWriteInt(sec.c_str(), L"EnableFolderExclude", IsChecked(g_chkEnableFolderExcl) ? 1 : 0);
    IniWriteInt(sec.c_str(), L"EnableNameExclude", IsChecked(g_chkEnableNameExcl) ? 1 : 0);
    IniWriteInt(sec.c_str(), L"NameIncludeExt", IsChecked(g_chkNameIncludeExt) ? 1 : 0);
    IniWriteStr(sec.c_str(), L"ResultFilter", GetWindowTextWStr(g_editFilter));

    IniWriteInt(sec.c_str(), L"ExcludeCount", (int)g_excludeRules.size());
    for (int i = 0; i < (int)g_excludeRules.size(); ++i) {
        const auto& r = g_excludeRules[(size_t)i];
        IniWriteInt(sec.c_str(), (L"ExcludeType" + std::to_wstring(i)).c_str(), (r.type == ExcludeType::DirPrefix) ? 0 : (r.type == ExcludeType::Wildcard ? 1 : 2));
        IniWriteStr(sec.c_str(), (L"ExcludeRaw" + std::to_wstring(i)).c_str(), (r.type == ExcludeType::DirPrefix) ? r.dirNorm.wstring() : r.raw);
    }

    IniWriteInt(sec.c_str(), L"NameCount", (int)g_fileNamePatterns.size());
    for (int i = 0; i < (int)g_fileNamePatterns.size(); ++i) {
        IniWriteStr(sec.c_str(), (L"NamePattern" + std::to_wstring(i)).c_str(), g_fileNamePatterns[(size_t)i]);
    }

    RefreshPresetCombo(name);
    SetStatus(L"検索条件プリセットを保存しました: " + name);
}

static void LoadSearchPreset() {
    std::wstring name = GetCurrentPresetName();
    if (name.empty()) return;
    const std::wstring sec = GetPresetSection(name);
    int rootsCount = IniReadInt(sec.c_str(), L"RootsCount", -1);
    if (rootsCount < 0) {
        MessageBoxW(g_hwndMain, L"指定したプリセットが見つかりません。", L"検索条件プリセット", MB_OK | MB_ICONINFORMATION);
        return;
    }

    std::vector<RootEntry> roots;
    for (int i = 0; i < rootsCount; ++i) {
        auto path = Trim(IniReadStr(sec.c_str(), (L"RootPath" + std::to_wstring(i)).c_str(), L""));
        bool enabled = IniReadInt(sec.c_str(), (L"RootEnabled" + std::to_wstring(i)).c_str(), 1) != 0;
        if (!path.empty()) roots.push_back({ path, enabled });
    }
    SetRootEntriesToListBox(roots);

    SendMessageW(g_cmbMode, CB_SETCURSEL, (WPARAM)max(0, min(2, IniReadInt(sec.c_str(), L"Mode", 0))), 0);
    SetWindowTextWStr(g_editDays, IniReadStr(sec.c_str(), L"Days", L"3"));
    SendMessageW(g_cmbTimeBase, CB_SETCURSEL, (WPARAM)max(0, min(2, IniReadInt(sec.c_str(), L"TimeBase", 0))), 0);
    SetChecked(g_chkXls, IniReadInt(sec.c_str(), L"xls", 1) != 0);
    SetChecked(g_chkXlsx, IniReadInt(sec.c_str(), L"xlsx", 1) != 0);
    SetChecked(g_chkXlsm, IniReadInt(sec.c_str(), L"xlsm", 1) != 0);
    SetChecked(g_chkXlsb, IniReadInt(sec.c_str(), L"xlsb", 1) != 0);
    SetChecked(g_chkXltx, IniReadInt(sec.c_str(), L"xltx", 1) != 0);
    SetChecked(g_chkXltm, IniReadInt(sec.c_str(), L"xltm", 1) != 0);
    SetChecked(g_chkEnableFolderExcl, IniReadInt(sec.c_str(), L"EnableFolderExclude", 1) != 0);
    SetChecked(g_chkEnableNameExcl, IniReadInt(sec.c_str(), L"EnableNameExclude", 1) != 0);
    SetChecked(g_chkNameIncludeExt, IniReadInt(sec.c_str(), L"NameIncludeExt", 1) != 0);
    SetWindowTextWStr(g_editFilter, IniReadStr(sec.c_str(), L"ResultFilter", L""));

    g_deferExcludeListRefresh = true;
    g_excludeRules.clear();
    int exclCount = IniReadInt(sec.c_str(), L"ExcludeCount", 0);
    for (int i = 0; i < exclCount; ++i) {
        int type = IniReadInt(sec.c_str(), (L"ExcludeType" + std::to_wstring(i)).c_str(), 2);
        auto raw = Trim(IniReadStr(sec.c_str(), (L"ExcludeRaw" + std::to_wstring(i)).c_str(), L""));
        if (raw.empty()) continue;
        if (type == 0) AddExcludeDirPrefix(raw);
        else AddOrUpdateExcludePatternOrSubstring(raw, -1);
    }
    g_deferExcludeListRefresh = false;
    RefreshExcludeListBox();

    g_deferFileNameListRefresh = true;
    g_fileNamePatterns.clear();
    int nameCount = IniReadInt(sec.c_str(), L"NameCount", 0);
    for (int i = 0; i < nameCount; ++i) {
        auto v = Trim(IniReadStr(sec.c_str(), (L"NamePattern" + std::to_wstring(i)).c_str(), L""));
        if (!v.empty()) g_fileNamePatterns.push_back(v);
    }
    if (g_fileNamePatterns.empty()) g_fileNamePatterns.push_back(L"~$");
    g_deferFileNameListRefresh = false;
    RefreshFileNameListBox();
    RebuildFileNameExcludeCache();

    ApplyLeftTabVisibility();
    if (g_hwndMain) {
        DoLayout(g_hwndMain);
    }
    SetListViewTimeHeader(GetTimeBase());
    RebuildListViewFromResults();
    SaveSettings();
    SetStatus(L"検索条件プリセットを読み込みました: " + name);
}

static void DeleteSearchPreset() {
    std::wstring name = GetCurrentPresetName();
    if (name.empty()) return;
    auto names = GetPresetNames();
    names.erase(std::remove_if(names.begin(), names.end(), [&](const auto& s) { return ToLower(s) == ToLower(name); }), names.end());
    SavePresetNames(names);
    WritePrivateProfileStringW(GetPresetSection(name).c_str(), nullptr, nullptr, g_iniPath.c_str());
    RefreshPresetCombo();
    SetWindowTextW(g_cmbPreset, L"");
    SetStatus(L"検索条件プリセットを削除しました: " + name);
}

static void SetSearchingUi(bool searching) {
    EnableWindow(g_btnSearch, searching ? FALSE : TRUE);
    EnableWindow(g_btnStop, searching ? TRUE : FALSE);
    EnableWindow(g_btnExportCsv, (!searching && g_visibleCount > 0) ? TRUE : FALSE);

    EnableWindow(g_btnBrowseRoot, !searching);
    EnableWindow(g_listRoots, !searching);
    EnableWindow(g_btnRootRemove, !searching);
    EnableWindow(g_btnRootUp, !searching);
    EnableWindow(g_btnRootDown, !searching);
    EnableWindow(g_btnRootToggle, !searching);
    EnableWindow(g_cmbMode, !searching);
    EnableWindow(g_editDays, !searching);
    EnableWindow(g_cmbTimeBase, !searching);
    EnableWindow(g_cmbPreset, !searching);
    EnableWindow(g_btnPresetSave, !searching);
    EnableWindow(g_btnPresetLoad, !searching);
    EnableWindow(g_btnPresetDelete, !searching);

    EnableWindow(g_chkXls, !searching);
    EnableWindow(g_chkXlsx, !searching);
    EnableWindow(g_chkXlsm, !searching);
    EnableWindow(g_chkXlsb, !searching);
    EnableWindow(g_chkXltx, !searching);
    EnableWindow(g_chkXltm, !searching);

    EnableWindow(g_chkEnableFolderExcl, !searching);
    EnableWindow(g_chkEnableNameExcl, !searching);

    EnableWindow(g_tabLeft, !searching);

    UpdateUiEnableStates();

    if (g_progress) {
        Progress_SetMarquee(g_progress, false);
        if (!searching) {
            SendMessageW(g_progress, PBM_SETPOS, 0, 0);
        }
    }
    if (g_staticProgress && !searching) {
        SetWindowTextW(g_staticProgress, L"待機中");
    }
    if (!searching) g_totalScanFiles = 0;
}
// -------------------- レイアウト --------------------

static void DoLayout(HWND hwnd) {
    RECT rc{};
    GetClientRect(hwnd, &rc);

    const int padding = 12;
    const int rowH = 38;
    const int btnH = 38;
    const int comboDropH = 260;
    const int dtpH = max(rowH, 32);
    const int gap = 8;
    const int labelH = 28;

    int winW = rc.right - rc.left;
    int winH = rc.bottom - rc.top;

    // ステータスバーより上のクライアント領域内にコントロールを収める
    const int statusH = 22;
    const int maxBottom = (winH - statusH - padding);

    // 2列レイアウト: 左側は操作欄、右側は結果一覧
    int minRightW = 440;
    int leftW = 560;
    if (winW < leftW + minRightW + padding * 3) {
        leftW = max(320, winW - (minRightW + padding * 3));
    }

    // 仮想的な左ペイン幅（余白を含む）
    int w = leftW + padding * 2;
    int rightX = w + padding;
    int rightW = winW - rightX - padding;
    if (rightW < 260) rightW = 260;

    int leftBoundary = rightX - padding; // 左ペイン内容の最大X座標

    // ---- 左ペイン共通（タブ） ----
    int y = padding;
    const int tabH = 30;
    MoveWindow(g_tabLeft, padding, y, max(160, w - padding * 2), tabH, TRUE);
    y += tabH + gap;

    // ---- 左ペイン: 検索タブ ----
    if (g_leftTab == 0) {
        // ルート見出し
        int labelW = 82;
        int hintW = w - (padding * 2 + labelW);
        hintW = max(80, hintW);
        MoveWindow(g_staticRoot, padding, y + 4, labelW, labelH, TRUE);
        MoveWindow(g_staticRootsHint, padding + labelW, y + 4, hintW, labelH, TRUE);
        y += rowH + gap;

        // ルート一覧とボタン
        int rootListH = max(150, rowH * 5 + gap * 4);
        int btnColW = 112;
        int rootListX = padding;
        int btnX = leftBoundary - btnColW;
        int rootListW = btnX - rootListX - padding;
        if (rootListW < 160) rootListW = 160;

        MoveWindow(g_listRoots, rootListX, y, rootListW, rootListH, TRUE);
        MoveWindow(g_btnBrowseRoot, btnX, y, btnColW, rowH, TRUE); // 追加...
        MoveWindow(g_btnRootRemove, btnX, y + (rowH + gap) * 1, btnColW, rowH, TRUE);
        MoveWindow(g_btnRootUp, btnX, y + (rowH + gap) * 2, btnColW, rowH, TRUE);
        MoveWindow(g_btnRootDown, btnX, y + (rowH + gap) * 3, btnColW, rowH, TRUE);
        MoveWindow(g_btnRootToggle, btnX, y + (rowH + gap) * 4, btnColW, rowH, TRUE);
        y += rootListH + gap;

        // カレンダー行
        int calLabelW = 56;
        int calW = 150;
        int calGap = 12;

        MoveWindow(g_staticFrom, padding, y + 4, calLabelW, labelH, TRUE);
        MoveWindow(g_dtpFrom, padding + calLabelW, y, calW, dtpH, TRUE);

        int x2 = padding + calLabelW + calW + calGap;
        MoveWindow(g_staticTo, x2, y + 4, calLabelW, labelH, TRUE);
        MoveWindow(g_dtpTo, x2 + calLabelW, y, calW, dtpH, TRUE);

        y += rowH + padding;

        // モード / 日数 / 日時基準（レスポンシブ配置）
        int modeLabelW = 82;
        int modeComboW = 214;
        int daysLabelW = 82;
        int daysEditW = 76;
        int tbLabelW = 82;
        int tbComboW = 180;

        int leftMaxX = w - padding;

        MoveWindow(g_staticMode, padding, y + 4, modeLabelW, labelH, TRUE);
        MoveWindow(g_cmbMode, padding + modeLabelW, y, modeComboW, comboDropH, TRUE);

        int x = padding + modeLabelW + modeComboW + padding;

        MoveWindow(g_staticDays, x, y + 4, daysLabelW, labelH, TRUE);
        MoveWindow(g_editDays, x + daysLabelW, y, daysEditW, rowH, TRUE);
        x += daysLabelW + daysEditW + padding;

        int need = tbLabelW + tbComboW;
        if (x + need > leftMaxX) {
            // 折り返し
            y += rowH + gap;
            x = padding;
        }
        int avail = leftMaxX - (x + tbLabelW);
        tbComboW = max(120, min(tbComboW, avail));
        MoveWindow(g_staticTimeBase, x, y + 4, tbLabelW, labelH, TRUE);
        MoveWindow(g_cmbTimeBase, x + tbLabelW, y, tbComboW, comboDropH, TRUE);
        y += rowH + gap;

        int presetLabelW = 82;
        int presetBtnW = 66;
        int presetBtnsW = presetBtnW * 3 + gap * 3;
        int presetComboW = max(150, w - padding * 2 - presetLabelW - presetBtnsW);
        MoveWindow(g_staticPreset, padding, y + 4, presetLabelW, labelH, TRUE);
        MoveWindow(g_cmbPreset, padding + presetLabelW, y, presetComboW, comboDropH, TRUE);
        int px = padding + presetLabelW + presetComboW + gap;
        MoveWindow(g_btnPresetSave, px, y, presetBtnW, rowH, TRUE);
        px += presetBtnW + gap;
        MoveWindow(g_btnPresetLoad, px, y, presetBtnW, rowH, TRUE);
        px += presetBtnW + gap;
        MoveWindow(g_btnPresetDelete, px, y, presetBtnW, rowH, TRUE);
        y += rowH + padding;

        // 拡張子フィルターを検索条件に含めるかどうかの設定。
        // オフのときは対象拡張子セクションを無効化し、拡張子では絞り込まない。
        MoveWindow(g_chkNameIncludeExt, padding, y, max(220, w - padding * 2), rowH, TRUE);
        y += rowH + gap;

        // 拡張子グループ（検索タブ内に保持）
        int extH = 88;
        int extW = w - padding * 2;
        HWND hExtGrp = GetDlgItem(hwnd, IDC_GRP_EXT);
        if (hExtGrp) MoveWindow(hExtGrp, padding, y, extW, extH, TRUE);

        int innerX = padding + 12;
        int innerY = y + 30;
        int extColW = max(160, (extW - 24) / 3);
        int lineH = 28;

        MoveWindow(g_chkXls, innerX + extColW * 0, innerY + lineH * 0, extColW, lineH, TRUE);
        MoveWindow(g_chkXlsx, innerX + extColW * 1, innerY + lineH * 0, extColW, lineH, TRUE);
        MoveWindow(g_chkXlsm, innerX + extColW * 2, innerY + lineH * 0, extColW, lineH, TRUE);

        MoveWindow(g_chkXlsb, innerX + extColW * 0, innerY + lineH * 1, extColW, lineH, TRUE);
        MoveWindow(g_chkXltx, innerX + extColW * 1, innerY + lineH * 1, extColW, lineH, TRUE);
        MoveWindow(g_chkXltm, innerX + extColW * 2, innerY + lineH * 1, extColW, lineH, TRUE);

        y += extH + gap;
    }
    // ---- 左ペイン: 除外タブ ----
    else {
        // 必要に応じて一覧の高さを縮め、最小ウィンドウサイズ内に詳細UIを収める
        int remainingH = maxBottom - y;
        const int fixedMin = 270;
        const int minFolderListH = 80;
        const int minNameListH = 70;
        int varTotal = remainingH - fixedMin;
        if (varTotal < (minFolderListH + minNameListH)) varTotal = (minFolderListH + minNameListH);

        int folderListH = min(160, max(minFolderListH, (varTotal * 2) / 3));
        int nameListHVar = min(140, max(minNameListH, varTotal - folderListH));

        const int gbTitleH = 24;
        const int gbBottomPad = 12;

        // -------- フォルダー除外グループ --------
        int folderGBTop = y;
        int fx = padding - 4;
        int fw = (w - padding * 2) + 8;

        y = folderGBTop + gbTitleH + 6;

        // 見出し行（ファイル名除外と同様にチェックボックスと操作ボタンを整列）
        MoveWindow(g_staticExclFolder, padding, y + 4, 140, labelH, TRUE);

        // セクション構造を揃えるため [フォルダ追加][削除] を右側に配置する
        int hdrBtnW = 110;
        int hdrBtnX2 = (w - padding) - hdrBtnW;                 // 削除
        int hdrBtnX1 = hdrBtnX2 - gap - hdrBtnW;               // フォルダ追加
        if (hdrBtnX1 < padding + 120 + 160) {                  // 幅が狭い場合はボタンを少し縮める
            hdrBtnW = 95;
            hdrBtnX2 = (w - padding) - hdrBtnW;
            hdrBtnX1 = hdrBtnX2 - gap - hdrBtnW;
        }

        MoveWindow(g_btnAddExclFolder, hdrBtnX1, y, hdrBtnW, rowH, TRUE);
        MoveWindow(g_btnRemoveExcl, hdrBtnX2, y, hdrBtnW, rowH, TRUE);

        int chkX0 = padding + 140;
        int chkW0 = hdrBtnX1 - gap - chkX0;
        chkW0 = min(260, max(160, chkW0));
        MoveWindow(g_chkEnableFolderExcl, chkX0, y, chkW0, rowH, TRUE);

        y += rowH + gap;

        // 一覧（全幅）と下部ボタン（ファイル名除外と同じ構造）
        MoveWindow(g_listExcludes, padding, y, w - padding * 2, folderListH, TRUE);
        y += folderListH + gap;

        // 下部ボタン行: 上へ / 下へ / 読込 / 保存
        int btnCountF = 4;
        int bwF = (w - padding * 2 - gap * (btnCountF - 1)) / btnCountF;
        if (bwF < 80) {
            int bw2 = (w - padding * 2 - gap) / 2;
            MoveWindow(g_btnExclUp, padding, y, bw2, rowH, TRUE);
            MoveWindow(g_btnExclDown, padding + bw2 + gap, y, bw2, rowH, TRUE);
            y += rowH + gap;
            MoveWindow(g_btnLoadExcl, padding, y, bw2, rowH, TRUE);
            MoveWindow(g_btnSaveExcl, padding + bw2 + gap, y, bw2, rowH, TRUE);
            y += rowH + gap;
        }
        else {
            MoveWindow(g_btnExclUp, padding + (bwF + gap) * 0, y, bwF, rowH, TRUE);
            MoveWindow(g_btnExclDown, padding + (bwF + gap) * 1, y, bwF, rowH, TRUE);
            MoveWindow(g_btnLoadExcl, padding + (bwF + gap) * 2, y, bwF, rowH, TRUE);
            MoveWindow(g_btnSaveExcl, padding + (bwF + gap) * 3, y, bwF, rowH, TRUE);
            y += rowH + gap;
        }// パターン入力（ファイル名除外と同じ形式）:
//  - タイトル行（静的テキスト）
//  - 次の行: 入力欄 + [追加] ボタン
        MoveWindow(g_staticExclPattern, padding, y + 4, w - padding * 2, labelH, TRUE);
        y += rowH + gap;

        int bW_pat = 110;
        int editW_pat = w - (padding * 3 + bW_pat);
        if (editW_pat >= 200) {
            MoveWindow(g_editExclPattern, padding, y, editW_pat, rowH, TRUE);
            MoveWindow(g_btnAddPattern, padding + editW_pat + padding, y, bW_pat, rowH, TRUE);
            y += rowH + gbBottomPad;
        }
        else {
            MoveWindow(g_editExclPattern, padding, y, w - padding * 2, rowH, TRUE);
            y += rowH + gap;

            int bw2 = max(90, (w - padding * 2));
            MoveWindow(g_btnAddPattern, padding, y, bw2, rowH, TRUE);
            y += rowH + gbBottomPad;
}
        int folderGBBottom = y + gbBottomPad;
        MoveWindow(g_frameFolderExcl, fx, folderGBTop, fw, max(60, folderGBBottom - folderGBTop), TRUE);
        SetWindowPos(g_frameFolderExcl, HWND_BOTTOM, 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

        y = folderGBBottom + gap;

        // -------- ファイル名除外グループ --------
        int nameGBTop = y;
        y = nameGBTop + gbTitleH + 6;
        // 見出し行（チェックボックス）
        MoveWindow(g_staticExclName, padding, y + 4, 140, labelH, TRUE);
        int chkX = padding + 140;
        int leftMaxX2 = w - padding;
        int avail2 = leftMaxX2 - chkX;
        int chk1W = min(260, max(160, avail2));
        MoveWindow(g_chkEnableNameExcl, chkX, y, chk1W, rowH, TRUE);
        y += rowH + gap;

        // パターン行: 入力欄 + 追加/削除（更新ボタン廃止：編集→Enter/フォーカスアウトで即反映）
        int bW2 = 95;
        int btnsW = bW2 * 2 + gap;
        int editW2 = w - (padding * 3 + btnsW);
        if (editW2 >= 200) {
            MoveWindow(g_editFNamePattern, padding, y, editW2, rowH, TRUE);
            int bx2 = padding + editW2 + padding;
            MoveWindow(g_btnAddFName, bx2 + (bW2 + gap) * 0, y, bW2, rowH, TRUE);
            MoveWindow(g_btnRemoveFName, bx2 + (bW2 + gap) * 1, y, bW2, rowH, TRUE);
            y += rowH + gap;
        }
        else {
            MoveWindow(g_editFNamePattern, padding, y, w - padding * 2, rowH, TRUE);
            y += rowH + gap;

            int bw2 = (w - padding * 2 - gap) / 2;
            bw2 = max(90, bw2);
            MoveWindow(g_btnAddFName, padding, y, bw2, rowH, TRUE);
            MoveWindow(g_btnRemoveFName, padding + bw2 + gap, y, bw2, rowH, TRUE);
            y += rowH + gap;
        }

        // 一覧とボタン
        MoveWindow(g_listFName, padding, y, w - padding * 2, nameListHVar, TRUE);
        y += nameListHVar + gap;

        int btnCount = 4;
        int bwN = (w - padding * 2 - gap * (btnCount - 1)) / btnCount;
        if (bwN < 80) {
            int bw2 = (w - padding * 2 - gap) / 2;
            MoveWindow(g_btnFNameUp, padding, y, bw2, rowH, TRUE);
            MoveWindow(g_btnFNameDown, padding + bw2 + gap, y, bw2, rowH, TRUE);
            y += rowH + gap;
            MoveWindow(g_btnLoadFNameExcl, padding, y, bw2, rowH, TRUE);
            MoveWindow(g_btnSaveFNameExcl, padding + bw2 + gap, y, bw2, rowH, TRUE);
            y += rowH + gbBottomPad;
        }
        else {
            MoveWindow(g_btnFNameUp, padding + (bwN + gap) * 0, y, bwN, rowH, TRUE);
            MoveWindow(g_btnFNameDown, padding + (bwN + gap) * 1, y, bwN, rowH, TRUE);
            MoveWindow(g_btnLoadFNameExcl, padding + (bwN + gap) * 2, y, bwN, rowH, TRUE);
            MoveWindow(g_btnSaveFNameExcl, padding + (bwN + gap) * 3, y, bwN, rowH, TRUE);
            y += rowH + gbBottomPad;
        }

        int nameGBBottom = y + gbBottomPad;
        MoveWindow(g_frameNameExcl, fx, nameGBTop, fw, max(60, nameGBBottom - nameGBTop), TRUE);
        SetWindowPos(g_frameNameExcl, HWND_BOTTOM, 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    }

    // ---- 右ペイン（常に表示） ----
    int yR = padding;

    // 操作ボタン行
    int bw = 132;
    int totalW = bw * 3 + gap * 2;
    if (rightW < totalW) bw = max(90, (rightW - gap * 2) / 3);

    MoveWindow(g_btnSearch, rightX + (bw + gap) * 0, yR, bw, btnH, TRUE);
    MoveWindow(g_btnStop, rightX + (bw + gap) * 1, yR, bw, btnH, TRUE);
    MoveWindow(g_btnExportCsv, rightX + (bw + gap) * 2, yR, bw, btnH, TRUE);
    yR += btnH + gap;

    // 進捗
    int progH = 20;
    int progTextW = 300;
    int progW = max(80, rightW - gap - progTextW);
    MoveWindow(g_progress, rightX, yR, progW, progH, TRUE);
    MoveWindow(g_staticProgress, rightX + progW + gap, yR + 1, max(80, rightW - progW - gap), progH, TRUE);
    yR += progH + gap;

    // フィルター
    int filterLabelW = 90;
    MoveWindow(g_staticFilter, rightX, yR + 4, filterLabelW, labelH, TRUE);
    MoveWindow(g_editFilter, rightX + filterLabelW, yR, rightW - filterLabelW, rowH, TRUE);
    yR += rowH + gap;
    MoveWindow(g_staticResultDetail, rightX, yR, rightW, rowH, TRUE);
    yR += rowH + padding;

    // 結果一覧
    int resultsH = maxBottom - yR;
    if (resultsH < 120) resultsH = 120;
    MoveWindow(g_listResults, rightX, yR, rightW, resultsH, TRUE);

    SendMessageW(g_status, WM_SIZE, 0, 0);
}


// -------------------- 右クリックメニュー --------------------
static void CopyTextToClipboard(HWND hwnd, const std::wstring& text) {
    if (!OpenClipboard(hwnd)) return;
    EmptyClipboard();
    size_t bytes = (text.size() + 1) * sizeof(wchar_t);
    HGLOBAL hg = GlobalAlloc(GMEM_MOVEABLE, bytes);
    if (hg) {
        void* p = GlobalLock(hg);
        memcpy(p, text.c_str(), bytes);
        GlobalUnlock(hg);
        SetClipboardData(CF_UNICODETEXT, hg);
    }
    CloseClipboard();
}

static const Hit* GetVisibleHitAt(int visibleIndex) {
    if (visibleIndex < 0 || visibleIndex >= (int)g_visibleResultIndices.size()) return nullptr;
    size_t realIndex = g_visibleResultIndices[(size_t)visibleIndex];
    if (realIndex >= g_results.size() || !g_results[realIndex]) return nullptr;
    return g_results[realIndex].get();
}

static void UpdateResultDetailFromSelection() {
    if (!g_staticResultDetail) return;
    int sel = ListView_GetNextItem(g_listResults, -1, LVNI_SELECTED);
    const Hit* hit = GetVisibleHitAt(sel);
    if (!hit) {
        SetWindowTextW(g_staticResultDetail, L"選択ファイル: なし");
        return;
    }
    std::wstring text = L"選択ファイル: " + hit->fileName + L"  |  " + hit->timeText + L"  |  " + std::to_wstring(hit->sizeKB) + L" KB  |  " + hit->path;
    SetWindowTextW(g_staticResultDetail, text.c_str());
}

static std::vector<std::wstring> GetSelectedVisibleResultPaths() {
    std::vector<std::wstring> paths;
    int sel = -1;
    while ((sel = ListView_GetNextItem(g_listResults, sel, LVNI_SELECTED)) >= 0) {
        if (const Hit* hit = GetVisibleHitAt(sel)) {
            paths.push_back(hit->path);
        }
    }
    std::sort(paths.begin(), paths.end());
    paths.erase(std::unique(paths.begin(), paths.end()), paths.end());
    return paths;
}

static void ShowResultsContextMenu(HWND hwnd, POINT ptScreen) {
    int sel = ListView_GetNextItem(g_listResults, -1, LVNI_SELECTED);
    const Hit* hit = GetVisibleHitAt(sel);
    if (!hit) return;

    const auto selectedPaths = GetSelectedVisibleResultPaths();

    HMENU hMenu = CreatePopupMenu();
    AppendMenuW(hMenu, MF_STRING, 1, L"開く");
    AppendMenuW(hMenu, MF_STRING, 2, L"フォルダを開く");
    AppendMenuW(hMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(hMenu, MF_STRING, CMD_RESULT_ADD_TO_PRINT, selectedPaths.size() > 1 ? L"選択行を印刷対象へ追加" : L"この行を印刷対象へ追加");
    AppendMenuW(hMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(hMenu, MF_STRING, 3, L"パスをコピー");
    AppendMenuW(hMenu, MF_STRING, 4, L"ファイル名をコピー");

    int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_RIGHTBUTTON, ptScreen.x, ptScreen.y, 0, hwnd, nullptr);
    DestroyMenu(hMenu);

    if (cmd == 1) {
        ShellExecuteW(hwnd, L"open", hit->path.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    }
    else if (cmd == 2) {
        fs::path p(hit->path);
        std::wstring folder = p.parent_path().wstring();
        ShellExecuteW(hwnd, L"open", folder.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    }
    else if (cmd == CMD_RESULT_ADD_TO_PRINT) {
        PrintToolPage_AppendFiles(selectedPaths.empty() ? std::vector<std::wstring>{ hit->path } : selectedPaths);
        SetStatus(L"選択した検索結果を印刷対象へ追加しました");
    }
    else if (cmd == 3) {
        CopyTextToClipboard(hwnd, hit->path);
    }
    else if (cmd == 4) {
        CopyTextToClipboard(hwnd, hit->fileName);
    }
}

static void ShowRootsContextMenu(HWND hwnd, POINT ptScreen) {
    int sel = (int)SendMessageW(g_listRoots, LB_GETCURSEL, 0, 0);
    int n = (int)SendMessageW(g_listRoots, LB_GETCOUNT, 0, 0);
    bool hasSel = (sel != LB_ERR);

    HMENU hMenu = CreatePopupMenu();
    AppendMenuW(hMenu, MF_STRING, CMD_ROOT_ADD, L"フォルダ追加...");
    AppendMenuW(hMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(hMenu, MF_STRING | (hasSel ? 0 : MF_GRAYED), CMD_ROOT_REMOVE, L"選択削除");
    AppendMenuW(hMenu, MF_STRING | ((hasSel && sel > 0) ? 0 : MF_GRAYED), CMD_ROOT_UP, L"上へ");
    AppendMenuW(hMenu, MF_STRING | ((hasSel && sel + 1 < n) ? 0 : MF_GRAYED), CMD_ROOT_DOWN, L"下へ");

    int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_RIGHTBUTTON, ptScreen.x, ptScreen.y, 0, hwnd, nullptr);
    DestroyMenu(hMenu);
    if (cmd == 0) return;

    if (cmd == CMD_ROOT_ADD) {
        std::wstring p;
        if (PickFolder(hwnd, p)) {
            AddRootToListBoxDedup(p);
            SaveSettings();
        }
        return;
    }
    if (!hasSel) return;

    if (cmd == CMD_ROOT_REMOVE) {
        SendMessageW(g_listRoots, LB_DELETESTRING, (WPARAM)sel, 0);
        SaveSettings();
    }
    else if (cmd == CMD_ROOT_UP || cmd == CMD_ROOT_DOWN) {
        int tgt = (cmd == CMD_ROOT_UP) ? (sel - 1) : (sel + 1);
        if (tgt >= 0 && tgt < n) {
            MoveRootItem(sel, tgt);
            SaveSettings();
        }
    }
}

static void ShowExcludesContextMenu(HWND hwnd, POINT ptScreen) {
    int sel = (int)SendMessageW(g_listExcludes, LB_GETCURSEL, 0, 0);
    int n = (int)g_excludeRules.size();
    bool hasSel = (sel != LB_ERR);

    HMENU hMenu = CreatePopupMenu();
    AppendMenuW(hMenu, MF_STRING, CMD_EXCL_ADD_FOLDER, L"フォルダ追加...");
    AppendMenuW(hMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(hMenu, MF_STRING | (hasSel ? 0 : MF_GRAYED), CMD_EXCL_REMOVE, L"選択削除");
    AppendMenuW(hMenu, MF_STRING | ((hasSel && sel > 0) ? 0 : MF_GRAYED), CMD_EXCL_UP, L"上へ");
    AppendMenuW(hMenu, MF_STRING | ((hasSel && sel + 1 < n) ? 0 : MF_GRAYED), CMD_EXCL_DOWN, L"下へ");
    AppendMenuW(hMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(hMenu, MF_STRING, CMD_EXCL_LOAD, L"読込...");
    AppendMenuW(hMenu, MF_STRING, CMD_EXCL_SAVE, L"保存...");

    int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_RIGHTBUTTON, ptScreen.x, ptScreen.y, 0, hwnd, nullptr);
    DestroyMenu(hMenu);
    if (cmd == 0) return;

    if (cmd == CMD_EXCL_ADD_FOLDER) {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDC_BTN_ADD_EXCL_FOLDER, 0), 0);
        SaveSettings();
        return;
    }
    if (cmd == CMD_EXCL_LOAD) {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDC_BTN_LOAD_EXCL, 0), 0);
        SaveSettings();
        return;
    }
    if (cmd == CMD_EXCL_SAVE) {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDC_BTN_SAVE_EXCL, 0), 0);
        SaveSettings();
        return;
    }

    if (!hasSel) return;

    if (cmd == CMD_EXCL_REMOVE) {
        if (sel >= 0 && sel < (int)g_excludeRules.size()) {
            g_excludeRules.erase(g_excludeRules.begin() + sel);
            RefreshExcludeListBox();
            SaveSettings();
        }
    }
    else if (cmd == CMD_EXCL_UP || cmd == CMD_EXCL_DOWN) {
        int tgt = (cmd == CMD_EXCL_UP) ? (sel - 1) : (sel + 1);
        if (tgt >= 0 && tgt < (int)g_excludeRules.size()) {
            std::swap(g_excludeRules[(size_t)sel], g_excludeRules[(size_t)tgt]);
            RefreshExcludeListBox();
            SendMessageW(g_listExcludes, LB_SETCURSEL, (WPARAM)tgt, 0);
            SaveSettings();
        }
    }
}

static void ShowFNameContextMenu(HWND hwnd, POINT ptScreen) {
    int sel = (int)SendMessageW(g_listFName, LB_GETCURSEL, 0, 0);
    int n = (int)g_fileNamePatterns.size();
    bool hasSel = (sel != LB_ERR);

    HMENU hMenu = CreatePopupMenu();
    AppendMenuW(hMenu, MF_STRING, CMD_FNAME_ADD, L"追加");
    AppendMenuW(hMenu, MF_STRING | (hasSel ? 0 : MF_GRAYED), CMD_FNAME_REMOVE, L"選択削除");
    AppendMenuW(hMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(hMenu, MF_STRING | ((hasSel && sel > 0) ? 0 : MF_GRAYED), CMD_FNAME_UP, L"上へ");
    AppendMenuW(hMenu, MF_STRING | ((hasSel && sel + 1 < n) ? 0 : MF_GRAYED), CMD_FNAME_DOWN, L"下へ");
    AppendMenuW(hMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(hMenu, MF_STRING, CMD_FNAME_LOAD, L"読込...");
    AppendMenuW(hMenu, MF_STRING, CMD_FNAME_SAVE, L"保存...");

    int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_RIGHTBUTTON, ptScreen.x, ptScreen.y, 0, hwnd, nullptr);
    DestroyMenu(hMenu);
    if (cmd == 0) return;

    if (cmd == CMD_FNAME_ADD) {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDC_BTN_ADD_FNAME, 0), 0);
        SaveSettings();
        return;
    }
    if (cmd == CMD_FNAME_REMOVE) {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDC_BTN_REMOVE_FNAME, 0), 0);
        SaveSettings();
        return;
    }
    if (cmd == CMD_FNAME_UP) {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDC_BTN_FNAME_UP, 0), 0);
        SaveSettings();
        return;
    }
    if (cmd == CMD_FNAME_DOWN) {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDC_BTN_FNAME_DOWN, 0), 0);
        SaveSettings();
        return;
    }
    if (cmd == CMD_FNAME_LOAD) {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDC_BTN_LOAD_FNAME_EXCL, 0), 0);
        SaveSettings();
        return;
    }
    if (cmd == CMD_FNAME_SAVE) {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDC_BTN_SAVE_FNAME_EXCL, 0), 0);
        SaveSettings();
        return;
    }
}

// -------------------- 検索スレッド --------------------
static DWORD WINAPI SearchThreadProc(LPVOID lpParam) {
    std::unique_ptr<SearchParams> params(reinterpret_cast<SearchParams*>(lpParam));
    g_stopRequested = false;
    g_searching = true;

    if (!params || params->roots.empty()) {
        PostMessageW(g_hwndMain, WM_APP_FINISHED, 0, 0);
        return 0;
    }
    std::error_code ec;
    const bool useFolderExcl = params->useFolderExcl;
    const bool useNameExcl = params->useNameExcl;
    const auto s = params->rangeStart;
    const auto e = params->rangeEnd;
    const TimeBase tb = params->timeBase;

    unsigned long long scanned = 0;
    unsigned long long hits = 0;
    unsigned long long totalFiles = 0;

    std::wstring lastDir;
    int dirNotifyCountdown = 0;

    // 本検索の前に通常ファイル数を数え、進捗バーへ全体件数を表示できるようにする
    for (const auto& root : params->roots) {
        if (g_stopRequested) break;
        fs::recursive_directory_iterator it(root, fs::directory_options::skip_permission_denied, ec);
        if (ec) continue;
        for (auto end = fs::recursive_directory_iterator(); it != end; ++it) {
            if (g_stopRequested) break;
            const auto& entry = *it;
            if (entry.is_directory(ec)) {
                // 除外フォルダーに一致した場合は再帰を止め、配下ファイルの無駄な走査を避ける
                if (useFolderExcl) {
                    auto dirNorm = NormalizePath(entry.path());
                    if (IsExcludedDir(dirNorm)) it.disable_recursion_pending();
                }
                continue;
            }
            if (entry.is_regular_file(ec)) totalFiles++;
        }
    }
    if (IsWindow(g_hwndMain)) PostMessageW(g_hwndMain, WM_APP_TOTAL, (WPARAM)totalFiles, 0);

    for (const auto& root : params->roots) {
        if (g_stopRequested) break;
        fs::recursive_directory_iterator it(root, fs::directory_options::skip_permission_denied, ec);
        if (ec) continue;
        for (auto end = fs::recursive_directory_iterator(); it != end; ++it) {
            if (g_stopRequested) break;

            const auto& entry = *it;

            // 走査中フォルダー名の通知は約200件ごとに間引き、UI更新の負荷を抑える
            if (dirNotifyCountdown-- <= 0) {
                dirNotifyCountdown = 200;
                std::wstring curDir = entry.path().parent_path().wstring();
                if (curDir != lastDir) {
                    lastDir = curDir;
                    PostMessageW(g_hwndMain, WM_APP_SCANPATH, 0, (LPARAM)new std::wstring(curDir));
                }
            }

            if (entry.is_directory(ec)) {
                if (useFolderExcl) {
                    auto dirNorm = NormalizePath(entry.path());
                    if (IsExcludedDir(dirNorm)) it.disable_recursion_pending();
                }
                continue;
            }

            if (!entry.is_regular_file(ec)) continue;

            scanned++;
            if ((scanned % 20) == 0 && IsWindow(g_hwndMain)) {
                PostMessageW(g_hwndMain, WM_APP_PROGRESS, (WPARAM)scanned, (LPARAM)hits);
            }

            const auto p = entry.path();
            // 対象拡張子とファイル名除外を先に判定し、不要な時刻取得やサイズ取得を避ける
            if (!IsTargetExcelFile(p)) continue;

            if (useNameExcl && IsExcludedByFileName(p)) continue;

            // 時刻取得（更新/作成/更新OR作成）
            std::chrono::system_clock::time_point w, c;
            if (!GetFileTimesSysClock(p.wstring(), w, c)) continue;

            bool inW = IsWithinLocalRange(w, s, e);
            bool inC = IsWithinLocalRange(c, s, e);

            std::chrono::system_clock::time_point tpShow{};
            if (tb == TimeBase::LastWrite) {
                if (!inW) continue;
                tpShow = w;
            }
            else if (tb == TimeBase::Creation) {
                if (!inC) continue;
                tpShow = c;
            }
            else { // どちらか
                if (!inW && !inC) continue;
                // 両方ヒットなら新しい方を表示
                tpShow = (inW && inC) ? (w > c ? w : c) : (inW ? w : c);
            }

            std::error_code ec3;
            uintmax_t sz = fs::file_size(p, ec3);
            unsigned long long kb = ec3 ? 0ULL : (unsigned long long)((sz + 1023) / 1024);

            auto hit = std::make_unique<Hit>();
            hit->timeText = FormatLocalTime(tpShow);
            hit->sizeKB = kb;
            hit->fileName = p.filename().wstring();
            hit->path = p.wstring();
            hit->fileNameLow = ToLower(hit->fileName);
            hit->pathLow = ToLower(hit->path);

            hits++;
            if (IsWindow(g_hwndMain)) {
                PostMessageW(g_hwndMain, WM_APP_PROGRESS, (WPARAM)scanned, (LPARAM)hits);
            }
            if (IsWindow(g_hwndMain)) {
                PostMessageW(g_hwndMain, WM_APP_ADD_HIT, 0, (LPARAM)hit.release());
            }
        }

    }

    if (IsWindow(g_hwndMain))
        PostMessageW(g_hwndMain, WM_APP_PROGRESS, (WPARAM)scanned, (LPARAM)hits);

    if (IsWindow(g_hwndMain))
        PostMessageW(g_hwndMain, WM_APP_FINISHED, 0, 0);

    return 0;
}

static std::wstring GetModeTextForStatus()
{
    int mode = GetMode();
    if (mode == 0) return L"今日";
    if (mode == 1) return L"過去 " + std::to_wstring(GetDaysFromEdit()) + L" 日";

    SYSTEMTIME stFrom = g_dateFrom;
    SYSTEMTIME stTo = g_dateTo;

    wchar_t a[32], b[32];
    _snwprintf_s(a, _TRUNCATE, L"%04u-%02u-%02u", stFrom.wYear, stFrom.wMonth, stFrom.wDay);
    _snwprintf_s(b, _TRUNCATE, L"%04u-%02u-%02u", stTo.wYear, stTo.wMonth, stTo.wDay);
    return std::wstring(a) + L" ～ " + b;
}

static void StartSearch() {
    if (g_searching) return;

    if (UseTargetExtensionFilter() && !AnyExtSelected()) {
        MessageBoxW(g_hwndMain, L"対象拡張子が1つも選択されていません。", L"確認", MB_OK | MB_ICONWARNING);
        return;
    }
    if (!g_listRoots || (int)SendMessageW(g_listRoots, LB_GETCOUNT, 0, 0) <= 0) {
        MessageBoxW(g_hwndMain, L"検索元フォルダが未設定です（[追加…]でフォルダを追加してください）。", L"確認", MB_OK | MB_ICONWARNING);
        return;
    }
    if (GetEnabledRootsFromListBox().empty()) {
        MessageBoxW(g_hwndMain, L"有効な検索先がありません。検索先を有効にしてください。", L"確認", MB_OK | MB_ICONWARNING);
        return;
    }
    RebuildFileNameExcludeCache();

    ClearResultsUI();
    SetSearchingUi(true);
    g_totalScanFiles = 0;
    if (g_staticProgress) SetWindowTextW(g_staticProgress, L"進捗: 開始中...");

    TimeBase tb = GetTimeBase();
    SetListViewTimeHeader(tb);

    std::wstring modeText = GetModeTextForStatus();
    SetStatus(L"検索中（" + modeText + L" / " + TimeBaseText(tb) + L"）...");

    auto params = std::make_unique<SearchParams>();
    params->useFolderExcl = IsChecked(g_chkEnableFolderExcl);
    params->useNameExcl = IsChecked(g_chkEnableNameExcl);
    GetActiveDateRange(params->rangeStart, params->rangeEnd);
    params->timeBase = tb;

    auto rootItems = GetEnabledRootsFromListBox();
    for (auto& s : rootItems) {
        std::error_code ec;
        fs::path r = NormalizePath(fs::path(s));
        if (fs::exists(r, ec) && fs::is_directory(r, ec)) params->roots.push_back(r);
    }
    if (params->roots.empty()) {
        SetSearchingUi(false);
        SetStatus(L"[ERROR] 検索元フォルダが存在しません");
        return;
    }

    SearchParams* threadParams = params.release();
    DWORD tid = 0;
    g_hThread = CreateThread(nullptr, 0, SearchThreadProc, threadParams, 0, &tid);
    if (!g_hThread) {
        delete threadParams;
        SetSearchingUi(false);
        SetStatus(L"[ERROR] スレッド作成に失敗しました");
    }
}
static void StopSearch() {
    if (!g_searching) return;
    g_stopRequested = true;
    SetStatus(L"停止要求中...");
}

// -------------------- ウィンドウプロシージャ --------------------
static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {

    case WM_GETMINMAXINFO: {
        auto* mmi = (MINMAXINFO*)lParam;
        mmi->ptMinTrackSize.x = 1150;
        mmi->ptMinTrackSize.y = 900;
        return 0;
    }

    case WM_CREATE:
    {
        g_hwndMain = hwnd;
        EnsureThemeBrushes();

        // 静的テキスト
        g_staticRoot = CreateWindowW(L"STATIC", L"検索先", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        g_staticRootsHint = CreateWindowW(L"STATIC", L"複数指定できます。ダブルクリックで有効/無効を切り替えます", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_STATIC_ROOTS_HINT, g_hInst, nullptr);
        g_staticMode = CreateWindowW(L"STATIC", L"期間", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        g_staticDays = CreateWindowW(L"STATIC", L"過去N日", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        g_staticTimeBase = CreateWindowW(L"STATIC", L"日時", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        g_staticPreset = CreateWindowW(L"STATIC", L"プリセット", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        g_staticFrom = CreateWindowW(L"STATIC", L"開始", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        g_staticTo = CreateWindowW(L"STATIC", L"終了", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        g_staticFilter = CreateWindowW(L"STATIC", L"絞り込み", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);

        g_staticExclFolder = CreateWindowW(L"STATIC", L"除外フォルダ:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        g_staticExclPattern = CreateWindowW(L"STATIC", L"フォルダ名部分一致/ワイルドカード", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        g_staticExclName = CreateWindowW(L"STATIC", L"ファイル名除外:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);

        // ルート操作コントロール
        g_btnBrowseRoot = CreateWindowW(L"BUTTON", L"フォルダ追加", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_BROWSE_ROOT, g_hInst, nullptr);

        // ルート一覧（複数フォルダーを直感的に扱う）
        g_listRoots = CreateWindowW(L"LISTBOX", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | LBS_NOTIFY | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS | WS_VSCROLL,
            0, 0, 0, 0, hwnd, (HMENU)IDC_LIST_ROOTS, g_hInst, nullptr);

        g_btnRootRemove = CreateWindowW(L"BUTTON", L"選択削除", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_ROOT_REMOVE, g_hInst, nullptr);
        g_btnRootUp = CreateWindowW(L"BUTTON", L"上へ", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_ROOT_UP, g_hInst, nullptr);
        g_btnRootDown = CreateWindowW(L"BUTTON", L"下へ", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_ROOT_DOWN, g_hInst, nullptr);
        g_btnRootToggle = CreateWindowW(L"BUTTON", L"有効切替", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_ROOT_TOGGLE, g_hInst, nullptr);

        // モード操作コントロール
        g_cmbMode = CreateWindowW(WC_COMBOBOXW, L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS | WS_VSCROLL,
            0, 0, 0, 0, hwnd, (HMENU)IDC_CMB_MODE, g_hInst, nullptr);
        SendMessageW(g_cmbMode, CB_ADDSTRING, 0, (LPARAM)L"今日");
        SendMessageW(g_cmbMode, CB_ADDSTRING, 0, (LPARAM)L"過去N日");
        SendMessageW(g_cmbMode, CB_ADDSTRING, 0, (LPARAM)L"期間指定（カレンダー）");


        g_editDays = CreateWindowW(L"EDIT", L"3", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | ES_NUMBER,
            0, 0, 0, 0, hwnd, (HMENU)IDC_EDIT_DAYS, g_hInst, nullptr);

        // 日時基準
        g_cmbTimeBase = CreateWindowW(WC_COMBOBOXW, L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS | WS_VSCROLL,
            0, 0, 0, 0, hwnd, (HMENU)IDC_CMB_TIMEBASE, g_hInst, nullptr);
        SendMessageW(g_cmbTimeBase, CB_ADDSTRING, 0, (LPARAM)L"更新日時");
        SendMessageW(g_cmbTimeBase, CB_ADDSTRING, 0, (LPARAM)L"作成日時");
        SendMessageW(g_cmbTimeBase, CB_ADDSTRING, 0, (LPARAM)L"更新OR作成");

        g_cmbPreset = CreateWindowW(WC_COMBOBOXW, L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWN | CBS_HASSTRINGS | WS_VSCROLL,
            0, 0, 0, 0, hwnd, (HMENU)IDC_CMB_PRESET, g_hInst, nullptr);
        SendMessageW(g_cmbPreset, EM_SETCUEBANNER, TRUE, (LPARAM)L"例: 月次検索");
        g_btnPresetSave = CreateWindowW(L"BUTTON", L"保存", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_PRESET_SAVE, g_hInst, nullptr);
        g_btnPresetLoad = CreateWindowW(L"BUTTON", L"読込", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_PRESET_LOAD, g_hInst, nullptr);
        g_btnPresetDelete = CreateWindowW(L"BUTTON", L"削除", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_PRESET_DELETE, g_hInst, nullptr);

        g_dtpFrom = CreateWindowW(L"STATIC", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | SS_NOTIFY,
            0, 0, 0, 0, hwnd, (HMENU)IDC_DTP_FROM, g_hInst, nullptr);

        g_dtpTo = CreateWindowW(L"STATIC", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | SS_NOTIFY,
            0, 0, 0, 0, hwnd, (HMENU)IDC_DTP_TO, g_hInst, nullptr);

        if (!g_hFontUi) {
            g_hFontUi = CreateFontW(
                -18, 0, 0, 0, FW_MEDIUM, FALSE, FALSE, FALSE,
                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                DEFAULT_PITCH | FF_DONTCARE, L"Yu Gothic UI");
        }
        if (!g_hFontUiBold) {
            g_hFontUiBold = CreateFontW(
                -20, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                DEFAULT_PITCH | FF_DONTCARE, L"Yu Gothic UI");
        }

        HFONT hUiFont = g_hFontUi ? g_hFontUi : (HFONT)GetStockObject(DEFAULT_GUI_FONT);
        HFONT hUiFontBold = g_hFontUiBold ? g_hFontUiBold : hUiFont;
        SendMessageW(g_cmbMode, WM_SETFONT, (WPARAM)hUiFont, TRUE);
        SendMessageW(g_cmbTimeBase, WM_SETFONT, (WPARAM)hUiFont, TRUE);
        SendMessageW(g_cmbPreset, WM_SETFONT, (WPARAM)hUiFont, TRUE);
        SendMessageW(g_dtpFrom, WM_SETFONT, (WPARAM)hUiFont, TRUE);
        SendMessageW(g_dtpTo, WM_SETFONT, (WPARAM)hUiFont, TRUE);

        // デフォルトは今日
        SYSTEMTIME st{};
        GetLocalTime(&st);
        g_dateFrom = st;
        g_dateTo = st;
        UpdateModernDatePickerText(g_dtpFrom);
        UpdateModernDatePickerText(g_dtpTo);

        // 左ペインのタブ（検索 / 除外）
        if (!g_hFontTabLeft) {
            g_hFontTabLeft = CreateFontW(
                -18, 0, 0, 0, FW_MEDIUM, FALSE, FALSE, FALSE,
                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                DEFAULT_PITCH | FF_DONTCARE, L"Yu Gothic UI");
        }
        g_tabLeft = CreateWindowExW(0, WC_TABCONTROLW, L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | TCS_FOCUSNEVER | TCS_OWNERDRAWFIXED,
            0, 0, 0, 0, hwnd, (HMENU)IDC_TAB_LEFT, g_hInst, nullptr);
        ApplyModernControlTheme(g_tabLeft);
        SendMessageW(g_tabLeft, WM_SETFONT, (WPARAM)(g_hFontTabLeft ? g_hFontTabLeft : hUiFont), TRUE);
        SendMessageW(g_tabLeft, TCM_SETITEMSIZE, 0, MAKELPARAM(132, 30));
        {
            TCITEMW ti{};
            ti.mask = TCIF_TEXT;
            ti.pszText = const_cast<LPWSTR>(L"検索条件");
            TabCtrl_InsertItem(g_tabLeft, 0, &ti);
            ti.pszText = const_cast<LPWSTR>(L"除外条件");
            TabCtrl_InsertItem(g_tabLeft, 1, &ti);
            TabCtrl_SetCurSel(g_tabLeft, 0);
        }

        // 詳細枠（コントロール背面）: セクションを視覚的に分けるために使用
        g_frameFolderExcl = CreateWindowW(L"BUTTON", L"フォルダ除外", WS_CHILD | BS_GROUPBOX,
            0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        g_frameNameExcl = CreateWindowW(L"BUTTON", L"ファイル名除外", WS_CHILD | BS_GROUPBOX,
            0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);

        // フォルダー除外コントロール
        g_chkEnableFolderExcl = CreateWindowW(L"BUTTON", L"フォルダ除外を有効", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
            0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_ENABLE_FOLDER_EXCL, g_hInst, nullptr);

        g_listExcludes = CreateWindowW(L"LISTBOX", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | LBS_NOTIFY | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS | WS_VSCROLL,
            0, 0, 0, 0, hwnd, (HMENU)IDC_LIST_EXCLUDES, g_hInst, nullptr);

        g_btnAddExclFolder = CreateWindowW(L"BUTTON", L"フォルダ追加", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_ADD_EXCL_FOLDER, g_hInst, nullptr);
        g_btnRemoveExcl = CreateWindowW(L"BUTTON", L"選択削除", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_REMOVE_EXCL, g_hInst, nullptr);
        g_btnExclUp = CreateWindowW(L"BUTTON", L"上へ", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_EXCL_UP, g_hInst, nullptr);
        g_btnExclDown = CreateWindowW(L"BUTTON", L"下へ", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_EXCL_DOWN, g_hInst, nullptr);
        g_btnLoadExcl = CreateWindowW(L"BUTTON", L"設定読込", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_LOAD_EXCL, g_hInst, nullptr);
        g_btnSaveExcl = CreateWindowW(L"BUTTON", L"設定保存", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_SAVE_EXCL, g_hInst, nullptr);

        g_editExclPattern = CreateWindowW(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, hwnd, (HMENU)IDC_EDIT_EXCL_PATTERN, g_hInst, nullptr);
        g_btnAddPattern = CreateWindowW(L"BUTTON", L"追加", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_ADD_PATTERN, g_hInst, nullptr);
        // 編集欄: Enter で即更新
        g_oldExclEditProc = (WNDPROC)SetWindowLongPtrW(g_editExclPattern, GWLP_WNDPROC, (LONG_PTR)ExclEditProc);

        // ファイル名除外
        g_chkEnableNameExcl = CreateWindowW(L"BUTTON", L"ファイル名除外を有効", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_ENABLE_NAME_EXCL, g_hInst, nullptr);
        g_chkNameIncludeExt = CreateWindowW(L"BUTTON", L"拡張子を含めて検索", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_NAME_INCLUDE_EXT, g_hInst, nullptr);
        g_editFNamePattern = CreateWindowW(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, hwnd, (HMENU)IDC_EDIT_FNAME_PATTERN, g_hInst, nullptr);
        g_btnAddFName = CreateWindowW(L"BUTTON", L"追加", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_ADD_FNAME, g_hInst, nullptr);
        g_btnRemoveFName = CreateWindowW(L"BUTTON", L"選択削除", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_REMOVE_FNAME, g_hInst, nullptr);
        g_oldFNameEditProc = (WNDPROC)SetWindowLongPtrW(g_editFNamePattern, GWLP_WNDPROC, (LONG_PTR)FNameEditProc);
        g_btnFNameUp = CreateWindowW(L"BUTTON", L"上へ", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_FNAME_UP, g_hInst, nullptr);
        g_btnFNameDown = CreateWindowW(L"BUTTON", L"下へ", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_FNAME_DOWN, g_hInst, nullptr);
        g_listFName = CreateWindowW(L"LISTBOX", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | LBS_NOTIFY | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS | WS_VSCROLL, 0, 0, 0, 0, hwnd, (HMENU)IDC_LIST_FNAME, g_hInst, nullptr);

        g_btnLoadFNameExcl = CreateWindowW(L"BUTTON", L"設定読込", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_LOAD_FNAME_EXCL, g_hInst, nullptr);
        g_btnSaveFNameExcl = CreateWindowW(L"BUTTON", L"設定保存", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_SAVE_FNAME_EXCL, g_hInst, nullptr);

        // 拡張子グループ
        CreateWindowW(L"BUTTON", L"対象拡張子", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 0, 0, 0, 0, hwnd, (HMENU)IDC_GRP_EXT, g_hInst, nullptr);

        g_chkXls = CreateWindowW(L"BUTTON", L".xls", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_XLS, g_hInst, nullptr);
        g_chkXlsx = CreateWindowW(L"BUTTON", L".xlsx", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_XLSX, g_hInst, nullptr);
        g_chkXlsm = CreateWindowW(L"BUTTON", L".xlsm", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_XLSM, g_hInst, nullptr);
        g_chkXlsb = CreateWindowW(L"BUTTON", L".xlsb", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_XLSB, g_hInst, nullptr);
        g_chkXltx = CreateWindowW(L"BUTTON", L".xltx", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_XLTX, g_hInst, nullptr);
        g_chkXltm = CreateWindowW(L"BUTTON", L".xltm", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_XLTM, g_hInst, nullptr);

        // 操作
        g_btnSearch = CreateWindowW(L"BUTTON", L"検索を開始", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_SEARCH, g_hInst, nullptr);
        g_btnStop = CreateWindowW(L"BUTTON", L"停止", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_STOP, g_hInst, nullptr);
        g_btnExportCsv = CreateWindowW(L"BUTTON", L"CSV出力", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_EXPORT_CSV, g_hInst, nullptr);

        // 初期は空(0%)表示。検索中だけマーキー表示を有効化する
        g_progress = CreateWindowW(PROGRESS_CLASSW, nullptr,
            WS_CHILD | WS_VISIBLE | PBS_SMOOTH,
            0, 0, 0, 0, hwnd, (HMENU)IDC_PROGRESS, g_hInst, nullptr);
        SendMessageW(g_progress, PBM_SETBKCOLOR, 0, (LPARAM)Theme::ProgressBg);
        SendMessageW(g_progress, PBM_SETBARCOLOR, 0, (LPARAM)Theme::Primary);
        Progress_SetMarquee(g_progress, false);
        g_staticProgress = CreateWindowW(L"STATIC", L"待機中", WS_CHILD | WS_VISIBLE | SS_LEFT | SS_CENTERIMAGE,
            0, 0, 0, 0, hwnd, (HMENU)IDC_STATIC_PROGRESS, g_hInst, nullptr);

        // 結果フィルター
        g_editFilter = CreateWindowW(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
            0, 0, 0, 0, hwnd, (HMENU)IDC_EDIT_FILTER, g_hInst, nullptr);
        SendMessageW(g_editFilter, EM_SETCUEBANNER, TRUE, (LPARAM)L"例: 入出荷 / 工場 / 2026");
        g_staticResultDetail = CreateWindowW(L"STATIC", L"選択ファイル: なし", WS_CHILD | WS_VISIBLE | SS_LEFT | SS_CENTERIMAGE,
            0, 0, 0, 0, hwnd, (HMENU)IDC_STATIC_RESULT_DETAIL, g_hInst, nullptr);

        // 結果
        g_listResults = CreateWindowW(WC_LISTVIEWW, L"", WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SHOWSELALWAYS, 0, 0, 0, 0, hwnd, (HMENU)IDC_LIST_RESULTS, g_hInst, nullptr);
        InitListViewColumns(g_listResults);

        // ステータス
        g_status = CreateWindowW(STATUSCLASSNAMEW, L"", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_STATUS, g_hInst, nullptr);

        HWND controls[] = {
            g_staticRoot, g_staticRootsHint, g_staticMode, g_staticDays, g_staticTimeBase, g_staticPreset, g_staticFrom, g_staticTo, g_staticFilter,
            g_staticExclFolder, g_staticExclPattern, g_staticExclName,
            g_btnBrowseRoot, g_listRoots, g_btnRootRemove, g_btnRootUp, g_btnRootDown, g_btnRootToggle,
            g_cmbMode, g_editDays, g_cmbTimeBase, g_cmbPreset, g_btnPresetSave, g_btnPresetLoad, g_btnPresetDelete, g_dtpFrom, g_dtpTo,
            g_frameFolderExcl, g_frameNameExcl,
            g_chkEnableFolderExcl, g_listExcludes, g_btnAddExclFolder, g_btnRemoveExcl, g_btnExclUp, g_btnExclDown, g_btnLoadExcl, g_btnSaveExcl,
            g_editExclPattern, g_btnAddPattern,
            g_chkEnableNameExcl, g_chkNameIncludeExt, g_editFNamePattern, g_btnAddFName, g_btnRemoveFName, g_btnFNameUp, g_btnFNameDown, g_listFName, g_btnLoadFNameExcl, g_btnSaveFNameExcl,
            GetDlgItem(hwnd, IDC_GRP_EXT), g_chkXls, g_chkXlsx, g_chkXlsm, g_chkXlsb, g_chkXltx, g_chkXltm,
            g_btnSearch, g_btnStop, g_btnExportCsv, g_progress, g_staticProgress, g_editFilter, g_staticResultDetail, g_listResults, g_status
        };
        for (HWND h : controls) {
            if (h) {
                ApplyModernControlTheme(h);
                SendMessageW(h, WM_SETFONT, (WPARAM)hUiFont, TRUE);
            }
        }
        ApplyModernDatePickerTheme(g_dtpFrom);
        ApplyModernDatePickerTheme(g_dtpTo);
        ApplyModernListBox(g_listRoots);
        ApplyModernListBox(g_listExcludes);
        ApplyModernListBox(g_listFName);
        ApplyModernComboBox(g_cmbMode);
        ApplyModernComboBox(g_cmbTimeBase);
        ApplyModernControlTheme(g_cmbPreset);
        ApplyModernResultsListView(g_listResults);

        HWND modernCheckBoxes[] = {
            g_chkEnableFolderExcl, g_chkEnableNameExcl, g_chkNameIncludeExt,
            g_chkXls, g_chkXlsx, g_chkXlsm, g_chkXlsb, g_chkXltx, g_chkXltm
        };
        for (HWND h : modernCheckBoxes) {
            ApplyModernCheckBox(h);
        }

        SendMessageW(g_btnSearch, WM_SETFONT, (WPARAM)hUiFontBold, TRUE);

        HWND modernButtons[] = {
            g_btnBrowseRoot, g_btnRootRemove, g_btnRootUp, g_btnRootDown, g_btnRootToggle,
            g_btnPresetSave, g_btnPresetLoad, g_btnPresetDelete,
            g_btnAddExclFolder, g_btnRemoveExcl, g_btnExclUp, g_btnExclDown, g_btnLoadExcl, g_btnSaveExcl,
            g_btnAddPattern, g_btnAddFName, g_btnRemoveFName, g_btnFNameUp, g_btnFNameDown,
            g_btnLoadFNameExcl, g_btnSaveFNameExcl, g_btnSearch, g_btnStop, g_btnExportCsv
        };
        for (HWND h : modernButtons) {
            EnableModernOwnerDrawButton(h);
            EnableButtonHoverHighlight(h);
        }

        // サーバー権限問題を避けるため、settings.ini / exclude.txt / results.csv は LocalAppData 配下に保存する
        InitPaths();
        EnsureUnicodeIniWithMigration();

        // 設定を読み込む
        LoadSettings();

        EnableWindow(g_btnStop, FALSE);
        EnableWindow(g_btnExportCsv, FALSE);

        SetSearchingUi(false);
        SetStatus(L"待機中（デフォルト: 今日 / 更新日時）");

        // 最初の WM_SIZE より前でも初期レイアウトを適用する
        DoLayout(hwnd);
        InvalidateRect(hwnd, nullptr, TRUE);

        return 0;
    }

    case WM_SIZE:
        SendMessageW(g_status, WM_SIZE, 0, 0);
        DoLayout(hwnd);
        return 0;

    case WM_CONTEXTMENU:
    {
        HWND src = (HWND)wParam;
        POINT pt{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        if ((short)pt.x == -1 && (short)pt.y == -1) {
            GetCursorPos(&pt);
        }

        if (src == g_listRoots) { ShowRootsContextMenu(hwnd, pt); return 0; }
        if (src == g_listExcludes) { ShowExcludesContextMenu(hwnd, pt); return 0; }
        if (src == g_listFName) { ShowFNameContextMenu(hwnd, pt); return 0; }
        break;
    }


    case WM_PAINT:
    {
        PAINTSTRUCT ps{};
        HDC hdc = BeginPaint(hwnd, &ps);
        PaintSearchBackground(hwnd, hdc);
        EndPaint(hwnd, &ps);
        return 0;
    }

    case WM_ERASEBKGND:
    {
        PaintSearchBackground(hwnd, reinterpret_cast<HDC>(wParam));
        return 1;
    }

    case WM_CTLCOLORSTATIC:
    {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        HWND ctrl = reinterpret_cast<HWND>(lParam);
        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, ctrl == g_staticRootsHint ? Theme::MutedText : Theme::Text);
        return reinterpret_cast<LRESULT>(g_hBrushCardBg ? g_hBrushCardBg : reinterpret_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
    }

    case WM_CTLCOLORBTN:
    {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        HWND ctrl = reinterpret_cast<HWND>(lParam);
        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, IsWindowEnabled(ctrl) ? Theme::Text : Theme::DisabledText);
        return reinterpret_cast<LRESULT>(g_hBrushCardBg ? g_hBrushCardBg : reinterpret_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
    }

    case WM_CTLCOLOREDIT:
    case WM_CTLCOLORLISTBOX:
    {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        SetBkColor(hdc, Theme::CardBg);
        SetTextColor(hdc, Theme::Text);
        return reinterpret_cast<LRESULT>(g_hBrushEditBg ? g_hBrushEditBg : reinterpret_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
    }

    case WM_DRAWITEM:
    {
        auto* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lParam);
        if (DrawModernTab(dis) || DrawModernButton(dis) || DrawModernListBox(dis) || DrawModernComboBox(dis)) return TRUE;
        break;
    }

    case WM_COMMAND:
    {
        int id = LOWORD(wParam);
        int code = HIWORD(wParam);

        // 更新ボタン廃止: 編集欄から直接更新（Enter またはフォーカスアウト）
        if (id == IDC_EDIT_EXCL_PATTERN && code == EN_KILLFOCUS) {
            CommitExcludeEditIfNeeded();
            return 0;
        }
        if (id == IDC_EDIT_FNAME_PATTERN && code == EN_KILLFOCUS) {
            CommitFileNameEditIfNeeded();
            return 0;
        }

        if (id == IDC_BTN_BROWSE_ROOT) {
            std::wstring p;
            if (PickFolder(hwnd, p)) {
                AddRootToListBoxDedup(p);
                SaveSettings();
            }
            return 0;
        }

        if (id == IDC_LIST_ROOTS && code == LBN_DBLCLK) {
            if (ToggleSelectedRootEnabled()) {
                SaveSettings();
            }
            return 0;
        }

        if (id == IDC_BTN_ROOT_REMOVE) {
            int sel = (int)SendMessageW(g_listRoots, LB_GETCURSEL, 0, 0);
            if (sel != LB_ERR) {
                SendMessageW(g_listRoots, LB_DELETESTRING, (WPARAM)sel, 0);
                SaveSettings();
            }
            return 0;
        }

        if (id == IDC_BTN_ROOT_TOGGLE) {
            if (ToggleSelectedRootEnabled()) {
                SaveSettings();
            }
            return 0;
        }
        if (id == IDC_BTN_ROOT_UP || id == IDC_BTN_ROOT_DOWN) {
            int sel = (int)SendMessageW(g_listRoots, LB_GETCURSEL, 0, 0);
            if (sel == LB_ERR) return 0;
            int n = (int)SendMessageW(g_listRoots, LB_GETCOUNT, 0, 0);
            int tgt = (id == IDC_BTN_ROOT_UP) ? (sel - 1) : (sel + 1);
            if (tgt < 0 || tgt >= n) return 0;
            MoveRootItem(sel, tgt);
            SaveSettings();
            return 0;
        }
        if (id == IDC_CMB_MODE && code == CBN_SELCHANGE) {
            UpdateUiEnableStates();
            return 0;
        }

        if (id == IDC_EDIT_FILTER && code == EN_CHANGE) {
            if (!g_searching) {
                RebuildListViewFromResults();
            }
            return 0;
        }

        if (id == IDC_CMB_TIMEBASE && code == CBN_SELCHANGE) {
            SetListViewTimeHeader(GetTimeBase());
            SetStatus(L"日時基準: " + TimeBaseText(GetTimeBase()));
            return 0;
        }
        if (id == IDC_BTN_PRESET_SAVE) {
            SaveSearchPreset();
            return 0;
        }
        if (id == IDC_BTN_PRESET_LOAD) {
            LoadSearchPreset();
            return 0;
        }
        if (id == IDC_BTN_PRESET_DELETE) {
            DeleteSearchPreset();
            return 0;
        }

        if (id == IDC_CHK_ENABLE_FOLDER_EXCL || id == IDC_CHK_ENABLE_NAME_EXCL || id == IDC_CHK_NAME_INCLUDE_EXT) {
            UpdateUiEnableStates();
            SaveSettings();
            return 0;
        }

        if (id == IDC_BTN_ADD_EXCL_FOLDER) {
            std::wstring p;
            if (PickFolder(hwnd, p)) AddExcludeDirPrefix(p);
            return 0;
        }

        if (id == IDC_LIST_EXCLUDES && (code == LBN_SELCHANGE || code == LBN_DBLCLK)) {
            return 0;
        }

        if (id == IDC_BTN_ADD_PATTERN) {
            AddOrUpdateExcludePatternOrSubstring(GetWindowTextWStr(g_editExclPattern), -1);
            SetWindowTextWStr(g_editExclPattern, L"");
            return 0;
        }

        // 更新ボタン廃止（Enter/フォーカスアウトで即反映）

        if (id == IDC_BTN_REMOVE_EXCL) {
            int sel = (int)SendMessageW(g_listExcludes, LB_GETCURSEL, 0, 0);
            if (sel != LB_ERR && sel >= 0 && sel < (int)g_excludeRules.size()) {
                g_excludeRules.erase(g_excludeRules.begin() + sel);
                RefreshExcludeListBox();
            }
            return 0;
        }

        if (id == IDC_BTN_EXCL_UP || id == IDC_BTN_EXCL_DOWN) {
            int sel = (int)SendMessageW(g_listExcludes, LB_GETCURSEL, 0, 0);
            if (sel == LB_ERR) return 0;
            int n = (int)g_excludeRules.size();
            int tgt = (id == IDC_BTN_EXCL_UP) ? (sel - 1) : (sel + 1);
            if (tgt < 0 || tgt >= n) return 0;
            std::swap(g_excludeRules[(size_t)sel], g_excludeRules[(size_t)tgt]);
            RefreshExcludeListBox();
            SendMessageW(g_listExcludes, LB_SETCURSEL, (WPARAM)tgt, 0);
            return 0;
        }

        if (id == IDC_BTN_LOAD_EXCL) {
            const wchar_t filter[] = L"Text Files (*.txt)\0*.txt\0All Files (*.*)\0*.*\0\0";
            std::wstring file = g_lastExcludeFile;
            if (ContainsQuestionMark(file)) file.clear(); // 壊れた初期パス対策
            if (PickOpenFile(hwnd, filter, file)) {
                auto rootsForBase = GetRootsFromListBox();
                std::wstring rootStr = rootsForBase.empty() ? L"" : rootsForBase[0];
                fs::path base = rootStr.empty() ? fs::path(GetExeDir()) : NormalizePath(fs::path(rootStr));
                if (LoadExcludesFromFile(file, base)) {
                    g_lastExcludeFile = file;
                    SetStatus(L"除外フォルダ: 読み込み完了");
                }
                else {
                    SetStatus(L"[ERROR] 除外フォルダ: 読み込み失敗");
                }
            }
            return 0;
        }

        if (id == IDC_BTN_SAVE_EXCL) {
            const wchar_t filter[] = L"Text Files (*.txt)\0*.txt\0All Files (*.*)\0*.*\0\0";
            std::wstring file = g_lastExcludeFile;
            if (ContainsQuestionMark(file)) file.clear(); // 壊れた初期パス対策
            if (PickSaveFile(hwnd, filter, L"txt", file)) {
                if (SaveExcludesToFile(file)) {
                    g_lastExcludeFile = file;
                    SetStatus(L"除外フォルダ: 保存完了");
                }
                else {
                    SetStatus(L"[ERROR] 除外フォルダ: 保存失敗");
                }
            }
            return 0;
        }

        // ファイル名除外の選択
        if (id == IDC_LIST_FNAME && (code == LBN_SELCHANGE || code == LBN_DBLCLK)) {
            return 0;
        }

        if (id == IDC_BTN_ADD_FNAME) {
            auto t = Trim(GetWindowTextWStr(g_editFNamePattern));
            if (!t.empty()) {
                auto low = ToLower(t);
                bool exists = false;
                for (auto& s : g_fileNamePatterns) if (ToLower(s) == low) { exists = true; break; }
                if (!exists) {
                    g_fileNamePatterns.push_back(t);
                    RefreshFileNameListBox();
                    RebuildFileNameExcludeCache();
                }
                SetWindowTextWStr(g_editFNamePattern, L"");
            }
            return 0;
        }


        if (id == IDC_BTN_REMOVE_FNAME) {
            int sel = (int)SendMessageW(g_listFName, LB_GETCURSEL, 0, 0);
            if (sel != LB_ERR && sel >= 0 && sel < (int)g_fileNamePatterns.size()) {
                g_fileNamePatterns.erase(g_fileNamePatterns.begin() + sel);
                if (g_fileNamePatterns.empty()) g_fileNamePatterns.push_back(L"~$");
                RefreshFileNameListBox();
                RebuildFileNameExcludeCache();
            }
            return 0;
        }

        if (id == IDC_BTN_FNAME_UP || id == IDC_BTN_FNAME_DOWN) {
            int sel = (int)SendMessageW(g_listFName, LB_GETCURSEL, 0, 0);
            if (sel == LB_ERR) return 0;
            int n = (int)g_fileNamePatterns.size();
            int tgt = (id == IDC_BTN_FNAME_UP) ? (sel - 1) : (sel + 1);
            if (tgt < 0 || tgt >= n) return 0;
            std::swap(g_fileNamePatterns[(size_t)sel], g_fileNamePatterns[(size_t)tgt]);
            RefreshFileNameListBox();
            SendMessageW(g_listFName, LB_SETCURSEL, (WPARAM)tgt, 0);
            RebuildFileNameExcludeCache();
            return 0;
        }


        if (id == IDC_BTN_LOAD_FNAME_EXCL) {
            const wchar_t filter[] = L"Text Files (*.txt)\0*.txt\0All Files (*.*)\0*.*\0\0";
            std::wstring file = g_lastNameExcludeFile;
            if (ContainsQuestionMark(file)) file.clear(); // 壊れた初期パス対策
            if (PickOpenFile(hwnd, filter, file)) {
                if (LoadFileNameExcludesFromFile(file)) {
                    g_lastNameExcludeFile = file;
                    SetStatus(L"除外ファイル名: 読み込み完了");
                }
                else {
                    SetStatus(L"[ERROR] 除外ファイル名: 読み込み失敗");
                }
            }
            return 0;
        }

        if (id == IDC_BTN_SAVE_FNAME_EXCL) {
            const wchar_t filter[] = L"Text Files (*.txt)\0*.txt\0All Files (*.*)\0*.*\0\0";
            std::wstring file = g_lastNameExcludeFile;
            if (ContainsQuestionMark(file)) file.clear(); // 壊れた初期パス対策
            if (PickSaveFile(hwnd, filter, L"txt", file)) {
                if (SaveFileNameExcludesToFile(file)) {
                    g_lastNameExcludeFile = file;
                    SetStatus(L"除外ファイル名: 保存完了");
                }
                else {
                    SetStatus(L"[ERROR] 除外ファイル名: 保存失敗");
                }
            }
            return 0;
        }

        if (id == IDC_BTN_EXPORT_CSV) {
            const wchar_t filter[] = L"CSV Files (*.csv)\0*.csv\0All Files (*.*)\0*.*\0\0";
            std::wstring file = g_lastCsvFile;
            if (ContainsQuestionMark(file)) file.clear(); // 壊れた初期パスでダイアログが落ちるのを防ぐ
            if (PickSaveFile(hwnd, filter, L"csv", file)) {
                if (ExportResultsCsv(file, GetTimeBase())) {
                    g_lastCsvFile = file;
                    SetStatus(L"CSVを出力しました");
                }
                else {
                    SetStatus(L"[ERROR] CSV出力に失敗しました");
                }
            }
            return 0;
        }

        if ((id == IDC_DTP_FROM || id == IDC_DTP_TO) && code == STN_CLICKED) {
            ShowModernCalendarPopup(reinterpret_cast<HWND>(lParam));
            return 0;
        }

        if (id == IDC_BTN_SEARCH) { StartSearch(); return 0; }
        if (id == IDC_BTN_STOP) { StopSearch();  return 0; }

        return 0;
    }

    case WM_NOTIFY:
    {
        LPNMHDR hdr = (LPNMHDR)lParam;
        if (!hdr) break;

        if (g_listResults && hdr->hwndFrom == ListView_GetHeader(g_listResults) && hdr->code == NM_CUSTOMDRAW) {
            return HandleResultsHeaderCustomDraw(reinterpret_cast<LPNMCUSTOMDRAW>(lParam));
        }

        if (hdr->hwndFrom == g_tabLeft) {
            if (hdr->code == TCN_SELCHANGE) {
                int sel = TabCtrl_GetCurSel(g_tabLeft);
                SetLeftTab(sel);
                return 0;
            }
        }

        if (hdr->hwndFrom == g_listResults) {
            if (hdr->code == NM_CUSTOMDRAW) {
                return HandleResultsCustomDraw(reinterpret_cast<LPNMLVCUSTOMDRAW>(lParam));
            }
            if (hdr->hwndFrom == g_listResults) {
                if (hdr->code == NM_DBLCLK) {
                    int sel = ListView_GetNextItem(g_listResults, -1, LVNI_SELECTED);
                    if (sel >= 0 && sel < (int)g_visibleResultIndices.size()) {
                        size_t realIndex = g_visibleResultIndices[(size_t)sel];
                        if (realIndex < g_results.size() && g_results[realIndex]) {
                            ShellExecuteW(hwnd, L"open", g_results[realIndex]->path.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
                        }
                    }
                    return 0;
                }
            }
            if (hdr->code == LVN_ITEMCHANGED) {
                UpdateResultDetailFromSelection();
                return 0;
            }
            if (hdr->code == LVN_COLUMNCLICK) {
                if (g_searching) return 0;
                auto* p = (NMLISTVIEW*)lParam;
                int col = p->iSubItem;
                if (col == g_sortCol) g_sortAsc = !g_sortAsc;
                else { g_sortCol = col; g_sortAsc = true; }
                SortResults(g_sortCol, g_sortAsc);
                return 0;
            }
            if (hdr->code == NM_RCLICK) {
                DWORD pos = GetMessagePos();
                POINT pt{ GET_X_LPARAM(pos), GET_Y_LPARAM(pos) };
                ShowResultsContextMenu(hwnd, pt);
                return 0;
            }
        }
        break;
    }

    case WM_APP_ADD_HIT:
    {
        std::unique_ptr<Hit> hit((Hit*)lParam);
        if (hit) AddResultToUI(std::move(hit));
        return 0;
    }

    case WM_APP_SCANPATH:
    {
        std::wstring* p = reinterpret_cast<std::wstring*>(lParam);
        if (p) {
            g_currentScanDir = *p;
            delete p;
        }
        else {
            g_currentScanDir.clear();
        }
        return 0;
    }

    case WM_APP_TOTAL:
    {
        g_totalScanFiles = (unsigned long long)wParam;
        return 0;
    }

    case WM_APP_PROGRESS:
    {
        unsigned long long scanned = (unsigned long long)wParam;
        unsigned long long hits = (unsigned long long)lParam;
        int percent = 0;
        if (g_totalScanFiles > 0) {
            percent = (int)((scanned * 100ULL) / g_totalScanFiles);
            if (percent > 100) percent = 100;
        }
        if (g_staticProgress) {
            std::wstring prog = L"進捗: " + std::to_wstring(percent) + L"%  (" + std::to_wstring(scanned) + L"/" + std::to_wstring(g_totalScanFiles) + L")  ヒット: " + std::to_wstring(hits);
            SetWindowTextW(g_staticProgress, prog.c_str());
        }
        if (g_progress) {
            SendMessageW(g_progress, PBM_SETRANGE32, 0, 100);
            SendMessageW(g_progress, PBM_SETPOS, (WPARAM)percent, 0);
        }

        std::wstring modeText = GetModeTextForStatus();
        {
            std::wstring __s = L"検索中（" + modeText + L" / " + TimeBaseText(GetTimeBase()) + L"） 走査: " +
                std::to_wstring(scanned) + L" / ヒット: " + std::to_wstring(hits);
            if (!g_currentScanDir.empty()) {
                __s += L"   |   走査中: " + EllipsizePathRight(g_currentScanDir, 60);
            }
            SetStatus(__s);
        }
        return 0;
    }

    case WM_APP_THREADERR:
    {
        const wchar_t* msgText = (const wchar_t*)lParam;
        if (msgText) MessageBoxW(hwnd, msgText, L"エラー", MB_OK | MB_ICONERROR);
        return 0;
    }

    case WM_APP_FINISHED:
    {
        g_searching = false;

        if (g_hThread) {
            CloseHandle(g_hThread);
            g_hThread = nullptr;
        }

        SetSearchingUi(false);

        SortResults(g_sortCol, g_sortAsc);

        std::wstring s = L"完了: " + std::to_wstring((int)g_results.size()) + L" 件見つかりました";
        if (g_stopRequested) s += L"（途中停止）";
        if (g_staticProgress) SetWindowTextW(g_staticProgress, s.c_str());
        SetStatus(s);

        EnableWindow(g_btnExportCsv, !g_results.empty());
        return 0;
    }

    case WM_CLOSE:
        if (g_searching) {
            StopSearch();
            MessageBoxW(hwnd, L"検索を停止しました。完了表示後に再度閉じてください。", L"停止", MB_OK | MB_ICONINFORMATION);
            return 0;
        }
        DestroyWindow(hwnd);
        return 0;

    case WM_DESTROY:
        CloseModernCalendarPopup();
        SaveSettings();
        if (g_hFontUi) { DeleteObject(g_hFontUi); g_hFontUi = nullptr; }
        if (g_hFontUiBold) { DeleteObject(g_hFontUiBold); g_hFontUiBold = nullptr; }
        if (g_hFontTabLeft) { DeleteObject(g_hFontTabLeft); g_hFontTabLeft = nullptr; }
        if (g_hBrushAppBg) { DeleteObject(g_hBrushAppBg); g_hBrushAppBg = nullptr; }
        if (g_hBrushCardBg) { DeleteObject(g_hBrushCardBg); g_hBrushCardBg = nullptr; }
        if (g_hBrushEditBg) { DeleteObject(g_hBrushEditBg); g_hBrushEditBg = nullptr; }
        if (g_hResultsRowImageList) { ImageList_Destroy(g_hResultsRowImageList); g_hResultsRowImageList = nullptr; }
        if (GetParent(hwnd) == nullptr) {
            PostQuitMessage(0);
        }
        return 0;
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}


// -------------------- ページ公開関数 --------------------
std::vector<std::wstring> SearchToolPage_GetResultPaths() {
    std::vector<std::wstring> out;
    out.reserve(g_results.size());
    for (const auto& hp : g_results) {
        if (!hp) continue;
        if (hp->path.empty()) continue;
        out.push_back(hp->path);
    }
    std::sort(out.begin(), out.end());
    out.erase(std::unique(out.begin(), out.end()), out.end());
    return out;
}

bool RegisterSearchToolPageClass(HINSTANCE hInstance) {
    g_hInst = hInstance;

    INITCOMMONCONTROLSEX icc{};
    icc.dwSize = sizeof(icc);
    icc.dwICC = ICC_WIN95_CLASSES | ICC_LISTVIEW_CLASSES | ICC_BAR_CLASSES | ICC_PROGRESS_CLASS | ICC_DATE_CLASSES | ICC_TAB_CLASSES;
    InitCommonControlsEx(&icc);

    static bool s_registered = false;
    if (s_registered) return true;

    const wchar_t CLASS_NAME[] = L"ExcelFinderAllInOneV5Page";

    WNDCLASSW wc{};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = CreateSolidBrush(Theme::AppBg);

    if (!RegisterClassW(&wc)) {
        DWORD err = GetLastError();
        if (err != ERROR_CLASS_ALREADY_EXISTS) return false;
    }

    s_registered = true;
    return true;
}

HWND CreateSearchToolPage(HWND parent, HINSTANCE hInstance, const RECT& rc) {
    g_hInst = hInstance;
    return CreateWindowW(
        L"ExcelFinderAllInOneV5Page",
        L"",
        WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS,
        rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top,
        parent, nullptr, hInstance, nullptr
    );
}

void SearchToolPage_SetVisible(HWND hwnd, bool visible) {
    if (!hwnd) return;
    ShowWindow(hwnd, visible ? SW_SHOW : SW_HIDE);
    if (visible) {
        InvalidateRect(hwnd, nullptr, TRUE);
        UpdateWindow(hwnd);
    }
}

void SearchToolPage_Resize(HWND hwnd, const RECT& rc) {
    if (!hwnd) return;
    MoveWindow(hwnd, rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top, TRUE);
    InvalidateRect(hwnd, nullptr, TRUE);
}
