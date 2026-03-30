//V_1.34
#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#include "SearchToolPage.h"
#include <windows.h>
#include <windowsx.h>   // GET_X_LPARAM / GET_Y_LPARAM
#include <commctrl.h>
#include <shobjidl.h>
#include <ShlObj.h>
#include <shellapi.h>
#include <commdlg.h>
#include <shlwapi.h>

#include <filesystem>
#include <string>
#include <vector>
#include <algorithm>
#include <chrono>
#include <atomic>
#include <memory>
#include <fstream>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "shlwapi.lib")

namespace fs = std::filesystem;

// -------------------- IDs --------------------
enum : int {
    IDC_EDIT_ROOT = 101,
    IDC_BTN_BROWSE_ROOT,

    // Roots list (NEW - intuitive multi-folder)
    IDC_LIST_ROOTS,
    IDC_BTN_ROOT_ADD,
    IDC_BTN_ROOT_REMOVE,
    IDC_BTN_ROOT_UP,
    IDC_BTN_ROOT_DOWN,
    IDC_STATIC_ROOTS_HINT,


    IDC_STATIC_MODE,
    IDC_CMB_MODE,
    IDC_STATIC_DAYS,
    IDC_EDIT_DAYS,

    // NEW: time base
    IDC_STATIC_TIMEBASE,
    IDC_CMB_TIMEBASE,

    // extensions
    IDC_GRP_EXT,
    IDC_CHK_XLS,
    IDC_CHK_XLSX,
    IDC_CHK_XLSM,
    IDC_CHK_XLSB,
    IDC_CHK_XLTX,
    IDC_CHK_XLTM,

    // folder excludes
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

    // file name excludes
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

    // file-name exclude load/save
    IDC_BTN_LOAD_FNAME_EXCL,
    IDC_BTN_SAVE_FNAME_EXCL,

    // actions
    IDC_BTN_SEARCH,
    IDC_BTN_STOP,
    IDC_BTN_EXPORT_CSV,

    // view
    IDC_PROGRESS,
    IDC_LIST_RESULTS,
    IDC_STATUS,

    // result filter (NEW)
    IDC_STATIC_FILTER,
    IDC_EDIT_FILTER,

    // calendar range (NEW)
    IDC_STATIC_FROM,
    IDC_DTP_FROM,
    IDC_STATIC_TO,
    IDC_DTP_TO,

    // Left panel tab (Search / Excludes)
    IDC_TAB_LEFT,

    // Advanced (legacy; unused in tab UI)
    IDC_BTN_TOGGLE_ADVANCED,
};

// Popup menu commands (not control IDs)
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
};

static const UINT WM_APP_ADD_HIT = WM_APP + 1;
static const UINT WM_APP_PROGRESS = WM_APP + 2;
static const UINT WM_APP_FINISHED = WM_APP + 3;
static const UINT WM_APP_THREADERR = WM_APP + 4;

static const UINT WM_APP_SCANPATH = WM_APP + 5; // NEW: show current scanning folder
// -------------------- Models --------------------
struct Hit {
    std::wstring timeText;        // 表示用（更新/作成どちらでも）
    unsigned long long sizeKB = 0;
    std::wstring fileName;
    std::wstring path;
};

enum class ExcludeType { DirPrefix, Wildcard, Substring };
struct ExcludeRule {
    ExcludeType type{};
    std::wstring raw;
    fs::path dirNorm;
    std::wstring pattern;
};

enum class TimeBase { LastWrite = 0, Creation = 1, Either = 2 };

// -------------------- Globals --------------------
static HINSTANCE g_hInst = nullptr;
static HWND g_hwndMain = nullptr;

static HFONT g_hFontUi = nullptr;
static HFONT g_hFontUiBold = nullptr;

static HWND g_tabLeft = nullptr;
static int  g_leftTab = 0; // 0: 検索, 1: 除外

static HWND g_staticRoot = nullptr;
static HWND g_staticMode = nullptr;
static HWND g_staticDays = nullptr;
static HWND g_staticTimeBase = nullptr;
static HWND g_staticExclFolder = nullptr;
static HWND g_staticExclPattern = nullptr;
static HWND g_staticExclName = nullptr;

static HWND g_editRoot = nullptr;      // legacy (kept but hidden)
static HWND g_btnBrowseRoot = nullptr; // repurposed as Add...
static HWND g_listRoots = nullptr;
static HWND g_btnRootRemove = nullptr;
static HWND g_btnRootUp = nullptr;
static HWND g_btnRootDown = nullptr;
static HWND g_staticRootsHint = nullptr;

static HWND g_cmbMode = nullptr;      // 0: 今日, 1: 過去N日
static HWND g_editDays = nullptr;

static HWND g_cmbTimeBase = nullptr;  // 0: 更新日時, 1: 作成日時, 2: 更新OR作成

static HWND g_chkXls = nullptr;
static HWND g_chkXlsx = nullptr;
static HWND g_chkXlsm = nullptr;
static HWND g_chkXlsb = nullptr;
static HWND g_chkXltx = nullptr;
static HWND g_chkXltm = nullptr;

// folder exclude
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

// name exclude
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

// actions/view
static HWND g_btnSearch = nullptr;
static HWND g_btnStop = nullptr;
static HWND g_btnExportCsv = nullptr;
static HWND g_progress = nullptr;
static HWND g_listResults = nullptr;
static HWND g_status = nullptr;

// result filter (NEW)
static HWND g_staticFilter = nullptr;
static HWND g_editFilter = nullptr;
static int g_visibleCount = 0;

static HWND g_staticFrom = nullptr;
static HWND g_staticTo = nullptr;
static HWND g_dtpFrom = nullptr;
static HWND g_dtpTo = nullptr;

// advanced toggle

// advanced frames (visual grouping)
static HWND g_frameFolderExcl = nullptr;
static HWND g_frameNameExcl = nullptr;

// runtime
static std::vector<ExcludeRule> g_excludeRules;
static std::vector<std::wstring> g_fileNamePatterns;
static std::vector<std::wstring> g_fileNameWild;
static std::vector<std::wstring> g_fileNameSubLow;

static std::vector<std::unique_ptr<Hit>> g_results;

static std::wstring g_iniPath;
static std::wstring g_currentScanDir; // NEW: scanning folder
static std::wstring g_lastExcludeFile;
static std::wstring g_lastCsvFile;
static std::wstring g_lastNameExcludeFile;

static std::atomic<bool> g_stopRequested{ false };
static std::atomic<bool> g_searching{ false };
static HANDLE g_hThread = nullptr;

// sorting: 0=日時,1=KB,2=ファイル名,3=パス
static int g_sortCol = 0;
static bool g_sortAsc = false;

// -------------------- Helpers --------------------
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
// ---- Forward declarations (for multi-root helpers) ----
static std::wstring ToLower(std::wstring s);
static std::wstring Trim(const std::wstring& s);
static std::wstring GetWindowTextWStr(HWND h);
static void SetWindowTextWStr(HWND h, const std::wstring& s);
static fs::path NormalizePath(const fs::path& p);

// ---- Forward declarations (used before definitions) ----
// Needed because some handlers/commit functions appear before these are defined.
static void RefreshFileNameListBox();
static void RebuildFileNameExcludeCache();
static void SaveSettings();
// -------------------------------------------------------

// ---- Roots list helpers (intuitive multi-folder) ----
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

static void SetRootsToListBox(const std::vector<std::wstring>& roots)
{
    if (!g_listRoots) return;
    SendMessageW(g_listRoots, LB_RESETCONTENT, 0, 0);
    for (const auto& s : roots) {
        auto t = Trim(s);
        if (t.empty()) continue;
        SendMessageW(g_listRoots, LB_ADDSTRING, 0, (LPARAM)t.c_str());
    }
}

static bool RootExistsInListBox(const std::wstring& path)
{
    auto low = ToLower(Trim(path));
    auto roots = GetRootsFromListBox();
    for (auto& r : roots) {
        if (ToLower(Trim(r)) == low) return true;
    }
    return false;
}

static void AddRootToListBoxDedup(const std::wstring& path)
{
    if (!g_listRoots) return;
    auto t = Trim(path);
    if (t.empty()) return;
    if (RootExistsInListBox(t)) return;
    SendMessageW(g_listRoots, LB_ADDSTRING, 0, (LPARAM)t.c_str());
}

static void MoveRootItem(int from, int to)
{
    if (!g_listRoots) return;
    int n = (int)SendMessageW(g_listRoots, LB_GETCOUNT, 0, 0);
    if (from < 0 || from >= n || to < 0 || to >= n || from == to) return;

    wchar_t buf[2048]{};
    SendMessageW(g_listRoots, LB_GETTEXT, (WPARAM)from, (LPARAM)buf);
    std::wstring item = buf;

    // delete old
    SendMessageW(g_listRoots, LB_DELETESTRING, (WPARAM)from, 0);

    // re-insert
    int idx = (int)SendMessageW(g_listRoots, LB_INSERTSTRING, (WPARAM)to, (LPARAM)item.c_str());
    SendMessageW(g_listRoots, LB_SETCURSEL, (WPARAM)idx, 0);
}

static std::vector<fs::path> GetRootPathsNormalizedFromUI()
{
    std::vector<fs::path> roots;
    auto lines = GetRootsFromListBox();
    for (auto& s : lines) roots.push_back(NormalizePath(fs::path(s)));
    return roots;
}
// ---------------------------------------------------

static std::vector<std::wstring> SplitRootsText(const std::wstring& text)
{
    // Accept: ';' and '|' separators (keep UI single-line). Also accept newlines if pasted.
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

    // Dedup (case-insensitive)
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

static void SyncLegacyRootEditFromList()
{
    // Search thread currently reads g_editRoot. Keep it in sync with root list.
    auto roots = GetRootsFromListBox();
    SetWindowTextWStr(g_editRoot, JoinRootsForIni(roots)); // '|' separated is OK for SplitRootsText
}

static std::vector<fs::path> GetRootPathsNormalized()
{
    std::vector<fs::path> roots;
    auto items = SplitRootsText(GetWindowTextWStr(g_editRoot));
    for (auto& s : items) roots.push_back(NormalizePath(fs::path(s)));
    return roots;
}

static void AppendRootToEditIfMissing(const std::wstring& folderPath)
{
    auto items = SplitRootsText(GetWindowTextWStr(g_editRoot));
    std::wstring lowNew = ToLower(Trim(folderPath));
    for (auto& s : items) {
        if (ToLower(Trim(s)) == lowNew) return;
    }
    items.push_back(Trim(folderPath));

    // UIは1行のまま：';'区切りで表示
    std::wstring disp;
    for (size_t i = 0; i < items.size(); ++i) {
        if (i) disp += L"; ";
        disp += items[i];
    }
    SetWindowTextWStr(g_editRoot, disp);
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
static bool IsUnderPath(const fs::path& p, const fs::path& prefix) {
    auto pit = p.begin();
    auto eit = prefix.begin();
    for (; eit != prefix.end(); ++eit, ++pit) {
        if (pit == p.end()) return false;
        if (ToLower(pit->wstring()) != ToLower(eit->wstring())) return false;
    }
    return true;
}

// ---- time ----
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

    std::time_t today0 = mktime(&startTm); // local midnight -> time_t(UTC基準)
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
    std::time_t t = mktime(&tm); // local time
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

    // mode == 2: 期間指定（カレンダー）
    SYSTEMTIME stFrom{}, stTo{};
    if (DateTime_GetSystemtime(g_dtpFrom, &stFrom) != GDT_VALID ||
        DateTime_GetSystemtime(g_dtpTo, &stTo) != GDT_VALID)
    {
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

// -------------------- Progress helpers --------------------
static void Progress_SetMarquee(HWND hProg, bool on) {
    if (!hProg) return;

    LONG_PTR style = GetWindowLongPtrW(hProg, GWL_STYLE);

    if (on) {
        // Ensure PBS_MARQUEE is set
        if (!(style & PBS_MARQUEE)) {
            SetWindowLongPtrW(hProg, GWL_STYLE, style | PBS_MARQUEE);
            SetWindowPos(hProg, nullptr, 0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
        }
        SendMessageW(hProg, PBM_SETMARQUEE, TRUE, 30);
    }
    else {
        // Stop marquee first
        SendMessageW(hProg, PBM_SETMARQUEE, FALSE, 0);

        // Remove PBS_MARQUEE so that PBM_SETPOS clears the bar visually
        style = GetWindowLongPtrW(hProg, GWL_STYLE);
        if (style & PBS_MARQUEE) {
            SetWindowLongPtrW(hProg, GWL_STYLE, style & ~PBS_MARQUEE);
            SetWindowPos(hProg, nullptr, 0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
        }

        SendMessageW(hProg, PBM_SETRANGE32, 0, 100);
        SendMessageW(hProg, PBM_SETPOS, 0, 0);

        // Force repaint to reflect the empty state
        InvalidateRect(hProg, nullptr, TRUE);
        UpdateWindow(hProg);
    }
}

// ---- NEW: FILETIME -> system_clock ----
static std::chrono::system_clock::time_point FileTimeToSysClock(const FILETIME& ft)
{
    ULARGE_INTEGER uli;
    uli.LowPart = ft.dwLowDateTime;
    uli.HighPart = ft.dwHighDateTime;

    constexpr int64_t WINDOWS_TICK_PER_SEC = 10'000'000LL;     // 100ns * 10,000,000 = 1 sec
    constexpr int64_t SEC_TO_UNIX_EPOCH = 11'644'473'600LL; // 1601-01-01 -> 1970-01-01 seconds

    int64_t total100ns = (int64_t)uli.QuadPart;
    int64_t sec = total100ns / WINDOWS_TICK_PER_SEC;
    int64_t rem = total100ns % WINDOWS_TICK_PER_SEC; // 100ns ticks

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

// LastWrite/Creation/Either 取得
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

static bool GetFileTimeSysClock(const std::wstring& path, TimeBase tb, std::chrono::system_clock::time_point& outTp)
{
    std::chrono::system_clock::time_point w, c;
    if (!GetFileTimesSysClock(path, w, c)) return false;
    outTp = (tb == TimeBase::Creation) ? c : w;
    return true;
}

// ---- extensions ----
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
static bool IsTargetExcelFile(const fs::path& p) {
    std::wstring e = ToLower(p.extension().wstring());
    if (e == L".xls" || e == L".xlsx" || e == L".xlsm" || e == L".xlsb" || e == L".xltx" || e == L".xltm") {
        return ExtChecked(e);
    }
    return false;
}

// -------------------- UTF-8 helpers --------------------
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

// -------------------- Dialog helpers --------------------
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

    DWORD err = CommDlgExtendedError(); // 0 = user canceled
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

    DWORD err = CommDlgExtendedError(); // 0 = user canceled
    if (err != 0) {
        wchar_t msg[128];
        wsprintfW(msg, L"GetSaveFileNameW failed: 0x%08X", err);
        MessageBoxW(owner, msg, L"Dialog error", MB_OK | MB_ICONERROR);
    }
    return false;
}

// -------------------- Folder excludes --------------------
static std::wstring FolderRuleToDisplay(const ExcludeRule& r) {
    switch (r.type) {
    case ExcludeType::DirPrefix: return L"[DIR] " + r.dirNorm.wstring();
    case ExcludeType::Wildcard:  return L"[WILD] " + r.raw;
    case ExcludeType::Substring: return L"[SUB] " + r.raw;
    default: return r.raw;
    }
}
static void RefreshExcludeListBox() {
    SendMessageW(g_listExcludes, LB_RESETCONTENT, 0, 0);
    for (const auto& r : g_excludeRules) {
        auto disp = FolderRuleToDisplay(r);
        SendMessageW(g_listExcludes, LB_ADDSTRING, 0, (LPARAM)disp.c_str());
    }
}
static bool IsExcludedDir(const fs::path& dirNorm) {
    std::wstring dirStr = dirNorm.wstring();
    std::wstring dirLower = ToLower(dirStr);

    for (const auto& r : g_excludeRules) {
        if (r.type == ExcludeType::DirPrefix) {
            if (IsUnderPath(dirNorm, r.dirNorm)) return true;
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

    auto rawLow = ToLower(r.raw);
    for (auto& e : g_excludeRules) {
        if (e.type == ExcludeType::DirPrefix && ToLower(e.dirNorm.wstring()) == rawLow) return;
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
        if (targetIndexOrMinus1 < 0) {
            auto low = ToLower(r.raw);
            for (auto& e : g_excludeRules) {
                if (e.type == ExcludeType::Wildcard && ToLower(e.raw) == low) return;
            }
        }
    }
    else {
        r.type = ExcludeType::Substring;
        r.pattern = ToLower(t);
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
            loaded.push_back(std::move(r));
        }
        else {
            ExcludeRule r;
            r.type = ExcludeType::Substring;
            r.raw = line;
            r.pattern = ToLower(line);
            loaded.push_back(std::move(r));
        }
    }

    auto key = [](const ExcludeRule& r) {
        if (r.type == ExcludeType::DirPrefix) return L"DIR:" + ToLower(r.dirNorm.wstring());
        if (r.type == ExcludeType::Wildcard)  return L"WILD:" + ToLower(r.raw);
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

// -------------------- Excludes: inline update helpers --------------------
static void UpdateExcludeDirPrefixAt(int index, const std::wstring& newPath)
{
    if (index < 0 || index >= (int)g_excludeRules.size()) return;
    auto t = Trim(newPath);
    if (t.empty()) return;

    fs::path p(t);
    auto norm = NormalizePath(p);
    auto low = ToLower(norm.wstring());

    // de-dup against other DIR rules
    for (int i = 0; i < (int)g_excludeRules.size(); ++i) {
        if (i == index) continue;
        const auto& r = g_excludeRules[(size_t)i];
        if (r.type == ExcludeType::DirPrefix && ToLower(r.dirNorm.wstring()) == low) {
            MessageBoxW(g_hwndMain, L"同じ除外フォルダが既に存在します。", L"更新できません", MB_OK | MB_ICONINFORMATION);
            return;
        }
    }

    auto& r = g_excludeRules[(size_t)index];
    r.type = ExcludeType::DirPrefix;
    r.dirNorm = norm;
    r.raw = norm.wstring();
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
        if (wParam == VK_RETURN) {
            CommitExcludeEditIfNeeded();
            // 余計なビープ防止
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
        if (wParam == VK_RETURN) {
            CommitFileNameEditIfNeeded();
            return 0;
        }
        if (wParam == VK_ESCAPE) {
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


// -------------------- File-name excludes --------------------
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
    SendMessageW(g_listFName, LB_RESETCONTENT, 0, 0);
    for (const auto& s : g_fileNamePatterns) {
        SendMessageW(g_listFName, LB_ADDSTRING, 0, (LPARAM)s.c_str());
    }
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

    // normalize & unique (case-insensitive)
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

// -------------------- Result filter (NEW) --------------------
static std::wstring GetFilterLow() {
    if (!g_editFilter) return L"";
    return ToLower(Trim(GetWindowTextWStr(g_editFilter)));
}
static bool HitMatchesFilterLow(const Hit& h, const std::wstring& fLow) {
    if (fLow.empty()) return true;
    std::wstring n = ToLower(h.fileName);
    std::wstring p = ToLower(h.path);
    return (n.find(fLow) != std::wstring::npos) || (p.find(fLow) != std::wstring::npos);
}
static void UpdateExportButtonEnabled() {
    if (!g_btnExportCsv) return;
    if (g_searching) {
        EnableWindow(g_btnExportCsv, FALSE);
        return;
    }
    EnableWindow(g_btnExportCsv, (g_visibleCount > 0) ? TRUE : FALSE);
}

// -------------------- CSV export --------------------
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

// -------------------- ListView --------------------
static void InitListViewColumns(HWND lv) {
    ListView_SetExtendedListViewStyle(lv, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_DOUBLEBUFFER);

    LVCOLUMNW col{};
    col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;

    col.pszText = (LPWSTR)L"日時";
    col.cx = 170; col.iSubItem = 0;
    ListView_InsertColumn(lv, 0, &col);

    col.pszText = (LPWSTR)L"KB";
    col.cx = 80; col.iSubItem = 1;
    ListView_InsertColumn(lv, 1, &col);

    col.pszText = (LPWSTR)L"ファイル名";
    col.cx = 240; col.iSubItem = 2;
    ListView_InsertColumn(lv, 2, &col);

    col.pszText = (LPWSTR)L"パス";
    col.cx = 680; col.iSubItem = 3;
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

static void ClearResultsUI() {
    g_results.clear();
    g_visibleCount = 0;
    ListView_DeleteAllItems(g_listResults);
    UpdateExportButtonEnabled();
}
static void AddResultToUI(std::unique_ptr<Hit> hit) {
    g_results.push_back(std::move(hit));
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
    g_visibleCount++;
    UpdateExportButtonEnabled();
}
static void RebuildListViewFromResults() {
    const std::wstring fLow = GetFilterLow();
    ListView_DeleteAllItems(g_listResults);
    g_visibleCount = 0;
    for (const auto& hp : g_results) {
        const Hit& h = *hp;
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
        g_visibleCount++;
    }
    UpdateExportButtonEnabled();
}
static void SortResults(int col, bool asc) {
    if (g_results.empty()) return;

    auto cmp = [&](const std::unique_ptr<Hit>& a, const std::unique_ptr<Hit>& b) {
        const Hit& A = *a;
        const Hit& B = *b;
        int r = 0;
        if (col == 0) {
            // YYYY-MM-DD HH:MM:SS なので文字列比較でOK
            r = (A.timeText < B.timeText) ? -1 : (A.timeText > B.timeText ? 1 : 0);
        }
        else if (col == 1) {
            r = (A.sizeKB < B.sizeKB) ? -1 : (A.sizeKB > B.sizeKB ? 1 : 0);
        }
        else if (col == 2) {
            auto aL = ToLower(A.fileName);
            auto bL = ToLower(B.fileName);
            r = (aL < bL) ? -1 : (aL > bL ? 1 : 0);
        }
        else {
            auto aL = ToLower(A.path);
            auto bL = ToLower(B.path);
            r = (aL < bL) ? -1 : (aL > bL ? 1 : 0);
        }
        return asc ? (r < 0) : (r > 0);
        };

    std::stable_sort(g_results.begin(), g_results.end(), cmp);
    RebuildListViewFromResults();
}

// -------------------- INI --------------------
// settings.ini を UTF-16LE(BOM) で確実に作成し、必要なら旧(サーバー)INIからローカルへ移行する
static bool ContainsQuestionMark(const std::wstring& s) {
    return s.find(L'?') != std::wstring::npos;
}

static void CreateUtf16LeBomFile(const std::wstring& path) {
    try {
        fs::path p(path);
        if (!p.parent_path().empty()) fs::create_directories(p.parent_path());
        std::ofstream f(p, std::ios::binary | std::ios::trunc);
        const unsigned char bom[] = { 0xFF, 0xFE }; // UTF-16LE
        f.write(reinterpret_cast<const char*>(bom), 2);
    }
    catch (...) {
        // ignore
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
    const std::wstring exeDir = GetExeDir(); // サーバー上でもOK
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
    const std::wstring legacyIni = exeDir + L"\\settings.ini"; // 旧: exeフォルダ（読むだけ）

    // 1) ローカルINIが無ければ Unicode(BOM) で作る
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

            // Ext
            const wchar_t* extKeys[] = { L"xls", L"xlsx", L"xlsm", L"xlsb", L"xltx", L"xltm" };
            for (auto k : extKeys) {
                wchar_t b[16];
                _snwprintf_s(b, _TRUNCATE, L"%d", IniReadIntFrom(legacyIni, L"Ext", k, 0));
                WritePrivateProfileStringW(L"Ext", k, b, g_iniPath.c_str());
            }

            // Filter
            const wchar_t* fKeys[] = { L"EnableFolderExclude", L"EnableNameExclude", L"NameIncludeExt" };
            for (auto k : fKeys) {
                wchar_t b[16];
                _snwprintf_s(b, _TRUNCATE, L"%d", IniReadIntFrom(legacyIni, L"Filter", k, 1));
                WritePrivateProfileStringW(L"Filter", k, b, g_iniPath.c_str());
            }

            // View
            {
                wchar_t b[16];
                _snwprintf_s(b, _TRUNCATE, L"%d", IniReadIntFrom(legacyIni, L"View", L"SortCol", 0));
                WritePrivateProfileStringW(L"View", L"SortCol", b, g_iniPath.c_str());
                _snwprintf_s(b, _TRUNCATE, L"%d", IniReadIntFrom(legacyIni, L"View", L"SortAsc", 0));
                WritePrivateProfileStringW(L"View", L"SortAsc", b, g_iniPath.c_str());
            }

            // NameFilter
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

// -------------------- UI enable/disable --------------------
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
    EnableWindow(g_chkNameIncludeExt, nameOn);
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

static void DoLayout(HWND hwnd); // forward

// -------------------- Left tab (Search / Excludes) --------------------
static void ApplyLeftTabVisibility() {
    bool isSearch = (g_leftTab == 0);
    int swSearch = isSearch ? SW_SHOW : SW_HIDE;
    int swExcl = isSearch ? SW_HIDE : SW_SHOW;

    // Search tab controls
    ShowWindow(g_staticRoot, swSearch);
    ShowWindow(g_listRoots, swSearch);
    ShowWindow(g_btnBrowseRoot, swSearch);
    ShowWindow(g_btnRootRemove, swSearch);
    ShowWindow(g_btnRootUp, swSearch);
    ShowWindow(g_btnRootDown, swSearch);
    ShowWindow(g_staticRootsHint, swSearch);

    ShowWindow(g_staticMode, swSearch);
    ShowWindow(g_cmbMode, swSearch);
    ShowWindow(g_staticDays, swSearch);
    ShowWindow(g_editDays, swSearch);
    ShowWindow(g_staticTimeBase, swSearch);
    ShowWindow(g_cmbTimeBase, swSearch);

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

    // Exclude tab controls
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
    ShowWindow(g_chkEnableNameExcl, swExcl);    ShowWindow(g_editFNamePattern, swExcl);
    ShowWindow(g_btnAddFName, swExcl);
    ShowWindow(g_btnRemoveFName, swExcl);
    ShowWindow(g_btnFNameUp, swExcl);
    ShowWindow(g_btnFNameDown, swExcl);
    ShowWindow(g_listFName, swExcl);
    ShowWindow(g_btnLoadFNameExcl, swExcl);
    ShowWindow(g_btnSaveFNameExcl, swExcl);

    // Keep enable/disable consistent when visible.
    UpdateUiEnableStates();
    if (g_hwndMain) InvalidateRect(g_hwndMain, nullptr, TRUE);
}

static void SetLeftTab(int tab, bool saveIni = true) {
    tab = max(0, min(1, tab));
    g_leftTab = tab;
    if (g_tabLeft) TabCtrl_SetCurSel(g_tabLeft, g_leftTab);
    ApplyLeftTabVisibility();
    if (g_hwndMain) {
        DoLayout(g_hwndMain);
        InvalidateRect(g_hwndMain, nullptr, TRUE);
    }
    if (saveIni) {
        IniWriteInt(L"View", L"LeftTab", g_leftTab);
    }
}

// -------------------- Settings load/save --------------------
static void LoadSettings() {
    auto rootsIni = IniReadStr(L"Main", L"Roots", L"");
    std::vector<std::wstring> rootsList;
    if (!rootsIni.empty()) {
        rootsList = SplitRootsText(rootsIni);
    }
    else {
        auto one = Trim(IniReadStr(L"Main", L"Root", L""));
        if (!one.empty()) rootsList.push_back(one);
    }
    SetRootsToListBox(rootsList);
    SyncLegacyRootEditFromList();
    // legacy edit is kept hidden
    if (!rootsList.empty()) SetWindowTextWStr(g_editRoot, rootsList[0]);
    g_lastExcludeFile = IniReadStr(L"Main", L"ExcludeFile", g_lastExcludeFile.empty() ? (GetExeDir() + L"\\exclude.txt") : g_lastExcludeFile);
    g_lastCsvFile = IniReadStr(L"Main", L"CsvFile", g_lastCsvFile.empty() ? (GetExeDir() + L"\\results.csv") : g_lastCsvFile);
    g_lastNameExcludeFile = IniReadStr(L"Main", L"NameExcludeFile", g_lastNameExcludeFile.empty() ? (GetExeDir() + L"\\name_exclude.txt") : g_lastNameExcludeFile);

    int mode = IniReadInt(L"Main", L"Mode", 0);
    // 0:今日 1:過去N日 2:期間指定
    SendMessageW(g_cmbMode, CB_SETCURSEL, (WPARAM)max(0, min(2, mode)), 0);
    SetWindowTextWStr(g_editDays, IniReadStr(L"Main", L"Days", L"3"));

    int tb = IniReadInt(L"Main", L"TimeBase", 0); // 0更新,1作成,2更新OR作成
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

    // load folder exclude
    if (!g_lastExcludeFile.empty() && fs::exists(fs::path(g_lastExcludeFile))) {
        auto rootsForBase = GetRootsFromListBox();
        std::wstring rootStr = rootsForBase.empty() ? L"" : rootsForBase[0];
        fs::path base = rootStr.empty() ? fs::path(GetExeDir()) : NormalizePath(fs::path(rootStr));
        LoadExcludesFromFile(g_lastExcludeFile, base);
    }

    // file-name patterns
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

    ApplyLeftTabVisibility();
    UpdateUiEnableStates();
    SetListViewTimeHeader(GetTimeBase());
}

static void SaveSettings() {
    auto rootItems = SplitRootsText(GetWindowTextWStr(g_editRoot));
    IniWriteStr(L"Main", L"Roots", JoinRootsForIni(rootItems));
    IniWriteStr(L"Main", L"Root", rootItems.empty() ? L"" : rootItems[0]);
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

    int oldCount = IniReadInt(L"NameFilter", L"Count", 0);
    int newCount = (int)g_fileNamePatterns.size();
    IniWriteInt(L"NameFilter", L"Count", newCount);
    int maxCount = (oldCount > newCount) ? oldCount : newCount;
    for (int i = 0; i < maxCount; ++i) {
        std::wstring key = L"Item" + std::to_wstring(i);
        if (i < newCount) IniWriteStr(L"NameFilter", key.c_str(), g_fileNamePatterns[(size_t)i]);
        else WritePrivateProfileStringW(L"NameFilter", key.c_str(), nullptr, g_iniPath.c_str());
    }

    if (g_lastExcludeFile.empty()) {
        auto ld = GetLocalAppDataExcelTodayDir();
        if (!ld.empty()) g_lastExcludeFile = (fs::path(ld) / L"exclude.txt").wstring();
        else g_lastExcludeFile = GetExeDir() + L"\\exclude.txt";
    }
    SaveExcludesToFile(g_lastExcludeFile);
}

// -------------------- Search UI state --------------------
static void SetSearchingUi(bool searching) {
    EnableWindow(g_btnSearch, searching ? FALSE : TRUE);
    EnableWindow(g_btnStop, searching ? TRUE : FALSE);
    EnableWindow(g_btnExportCsv, (!searching && g_visibleCount > 0) ? TRUE : FALSE);

    EnableWindow(g_btnBrowseRoot, !searching);
    EnableWindow(g_listRoots, !searching);
    EnableWindow(g_btnRootRemove, !searching);
    EnableWindow(g_btnRootUp, !searching);
    EnableWindow(g_btnRootDown, !searching);
    EnableWindow(g_editRoot, !searching);
    EnableWindow(g_cmbMode, !searching);
    EnableWindow(g_editDays, !searching);
    EnableWindow(g_cmbTimeBase, !searching);

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
        Progress_SetMarquee(g_progress, searching);
    }
}

// -------------------- Layout --------------------

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

    // Keep controls within client area above the status bar.
    const int statusH = 22;
    const int maxBottom = (winH - statusH - padding);

    // Two-column layout: left = controls, right = results list
    int minRightW = 440;
    int leftW = 560;
    if (winW < leftW + minRightW + padding * 3) {
        leftW = max(320, winW - (minRightW + padding * 3));
    }

    // Virtual left panel width (includes padding margins)
    int w = leftW + padding * 2;
    int rightX = w + padding;
    int rightW = winW - rightX - padding;
    if (rightW < 260) rightW = 260;

    int leftBoundary = rightX - padding; // max x for left panel content

    // ---- Left panel common (tab) ----
    int y = padding;
    const int tabH = 40;
    MoveWindow(g_tabLeft, padding, y, max(160, w - padding * 2), tabH, TRUE);
    TabCtrl_SetItemSize(g_tabLeft, 0, MAKELPARAM(96, 30));
    y += tabH + gap;

    // ---- Left panel: Search tab ----
    if (g_leftTab == 0) {
        // Root header
        int labelW = 82;
        int hintW = w - (padding * 2 + labelW);
        hintW = max(80, hintW);
        MoveWindow(g_staticRoot, padding, y + 4, labelW, labelH, TRUE);
        MoveWindow(g_staticRootsHint, padding + labelW, y + 4, hintW, labelH, TRUE);
        y += rowH + gap;

        // Roots list + buttons
        int rootListH = max(112, rowH * 4 + gap * 3);
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
        y += rootListH + gap;

        // Calendar row
        int calLabelW = 56;
        int calW = 150;
        int calGap = 12;

        MoveWindow(g_staticFrom, padding, y + 4, calLabelW, labelH, TRUE);
        MoveWindow(g_dtpFrom, padding + calLabelW, y, calW, dtpH, TRUE);

        int x2 = padding + calLabelW + calW + calGap;
        MoveWindow(g_staticTo, x2, y + 4, calLabelW, labelH, TRUE);
        MoveWindow(g_dtpTo, x2 + calLabelW, y, calW, dtpH, TRUE);

        y += rowH + padding;

        // Mode / Days / TimeBase (responsive)
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
            // wrap
            y += rowH + gap;
            x = padding;
        }
        int avail = leftMaxX - (x + tbLabelW);
        tbComboW = max(120, min(tbComboW, avail));
        MoveWindow(g_staticTimeBase, x, y + 4, tbLabelW, labelH, TRUE);
        MoveWindow(g_cmbTimeBase, x + tbLabelW, y, tbComboW, comboDropH, TRUE);
        y += rowH + padding;
        // Name-exclude option (判定に拡張子を含める)
        // これは「ファイル名除外」で拡張子を含めて比較するかどうかの設定。
        MoveWindow(g_chkNameIncludeExt, padding, y, max(220, w - padding * 2), rowH, TRUE);
        y += rowH + gap;

        // Extensions group (kept in Search tab)
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
    // ---- Left panel: Excludes tab ----
    else {
        // Fit advanced UI within the minimum window size by shrinking list heights when needed.
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

        // -------- Folder exclude group --------
        int folderGBTop = y;
        int fx = padding - 4;
        int fw = (w - padding * 2) + 8;

        y = folderGBTop + gbTitleH + 6;

        // Header row (similar to file-name exclude: checkbox + action buttons aligned)
        MoveWindow(g_staticExclFolder, padding, y + 4, 140, labelH, TRUE);

        // Place [フォルダ追加][削除] on the right to keep the section structure consistent.
        int hdrBtnW = 110;
        int hdrBtnX2 = (w - padding) - hdrBtnW;                 // 削除
        int hdrBtnX1 = hdrBtnX2 - gap - hdrBtnW;               // フォルダ追加
        if (hdrBtnX1 < padding + 120 + 160) {                  // too narrow → shrink buttons a bit
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

        // List (full width) + bottom buttons (same structure as file-name exclude)
        MoveWindow(g_listExcludes, padding, y, w - padding * 2, folderListH, TRUE);
        y += folderListH + gap;

        // Bottom buttons row: 上へ / 下へ / 読込 / 保存
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
        }// Pattern input (match file-name exclude style):
//  - Title line (static)
//  - Next line: edit + [追加] button
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

        // -------- File name exclude group --------
        int nameGBTop = y;
        y = nameGBTop + gbTitleH + 6;
        // Header row (checkbox)
        MoveWindow(g_staticExclName, padding, y + 4, 140, labelH, TRUE);
        int chkX = padding + 140;
        int leftMaxX2 = w - padding;
        int avail2 = leftMaxX2 - chkX;
        int chk1W = min(260, max(160, avail2));
        MoveWindow(g_chkEnableNameExcl, chkX, y, chk1W, rowH, TRUE);
        y += rowH + gap;

        // Pattern row: edit + add/remove（更新ボタン廃止：編集→Enter/フォーカスアウトで即反映）
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

        // List + buttons
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

    // ---- Right panel (always visible) ----
    int yR = padding;

    // Action buttons row
    int bw = 132;
    int totalW = bw * 3 + gap * 2;
    if (rightW < totalW) bw = max(90, (rightW - gap * 2) / 3);

    MoveWindow(g_btnSearch, rightX + (bw + gap) * 0, yR, bw, btnH, TRUE);
    MoveWindow(g_btnStop, rightX + (bw + gap) * 1, yR, bw, btnH, TRUE);
    MoveWindow(g_btnExportCsv, rightX + (bw + gap) * 2, yR, bw, btnH, TRUE);
    yR += btnH + gap;

    // Progress
    int progH = 20;
    MoveWindow(g_progress, rightX, yR, max(80, rightW), progH, TRUE);
    yR += progH + gap;

    // Filter
    int filterLabelW = 90;
    MoveWindow(g_staticFilter, rightX, yR + 4, filterLabelW, labelH, TRUE);
    MoveWindow(g_editFilter, rightX + filterLabelW, yR, rightW - filterLabelW, rowH, TRUE);
    yR += rowH + padding;

    // Results list
    int resultsH = maxBottom - yR;
    if (resultsH < 120) resultsH = 120;
    MoveWindow(g_listResults, rightX, yR, rightW, resultsH, TRUE);

    SendMessageW(g_status, WM_SIZE, 0, 0);
}


// -------------------- Right-click menus --------------------
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

static void ShowResultsContextMenu(HWND hwnd, POINT ptScreen) {
    int sel = ListView_GetNextItem(g_listResults, -1, LVNI_SELECTED);
    if (sel < 0 || sel >= (int)g_results.size()) return;
    const auto& hit = *g_results[(size_t)sel];

    HMENU hMenu = CreatePopupMenu();
    AppendMenuW(hMenu, MF_STRING, 1, L"開く");
    AppendMenuW(hMenu, MF_STRING, 2, L"フォルダを開く");
    AppendMenuW(hMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(hMenu, MF_STRING, 3, L"パスをコピー");
    AppendMenuW(hMenu, MF_STRING, 4, L"ファイル名をコピー");

    int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_RIGHTBUTTON, ptScreen.x, ptScreen.y, 0, hwnd, nullptr);
    DestroyMenu(hMenu);

    if (cmd == 1) {
        ShellExecuteW(hwnd, L"open", hit.path.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    }
    else if (cmd == 2) {
        fs::path p(hit.path);
        std::wstring folder = p.parent_path().wstring();
        ShellExecuteW(hwnd, L"open", folder.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    }
    else if (cmd == 3) {
        CopyTextToClipboard(hwnd, hit.path);
    }
    else if (cmd == 4) {
        CopyTextToClipboard(hwnd, hit.fileName);
    }
}

static void ShowRootsContextMenu(HWND hwnd, POINT ptScreen) {
    int sel = (int)SendMessageW(g_listRoots, LB_GETCURSEL, 0, 0);
    int n = (int)SendMessageW(g_listRoots, LB_GETCOUNT, 0, 0);
    bool hasSel = (sel != LB_ERR);

    HMENU hMenu = CreatePopupMenu();
    AppendMenuW(hMenu, MF_STRING, CMD_ROOT_ADD, L"フォルダ追加...");
    AppendMenuW(hMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(hMenu, MF_STRING | (hasSel ? 0 : MF_GRAYED), CMD_ROOT_REMOVE, L"削除");
    AppendMenuW(hMenu, MF_STRING | ((hasSel && sel > 0) ? 0 : MF_GRAYED), CMD_ROOT_UP, L"上へ");
    AppendMenuW(hMenu, MF_STRING | ((hasSel && sel + 1 < n) ? 0 : MF_GRAYED), CMD_ROOT_DOWN, L"下へ");

    int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_RIGHTBUTTON, ptScreen.x, ptScreen.y, 0, hwnd, nullptr);
    DestroyMenu(hMenu);
    if (cmd == 0) return;

    if (cmd == CMD_ROOT_ADD) {
        std::wstring p;
        if (PickFolder(hwnd, p)) {
            AddRootToListBoxDedup(p);
            SyncLegacyRootEditFromList();
            SaveSettings();
        }
        return;
    }
    if (!hasSel) return;

    if (cmd == CMD_ROOT_REMOVE) {
        SendMessageW(g_listRoots, LB_DELETESTRING, (WPARAM)sel, 0);
        SyncLegacyRootEditFromList();
        SaveSettings();
    }
    else if (cmd == CMD_ROOT_UP || cmd == CMD_ROOT_DOWN) {
        int tgt = (cmd == CMD_ROOT_UP) ? (sel - 1) : (sel + 1);
        if (tgt >= 0 && tgt < n) {
            MoveRootItem(sel, tgt);
            SyncLegacyRootEditFromList();
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
    AppendMenuW(hMenu, MF_STRING | (hasSel ? 0 : MF_GRAYED), CMD_EXCL_REMOVE, L"削除");
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
    AppendMenuW(hMenu, MF_STRING | (hasSel ? 0 : MF_GRAYED), CMD_FNAME_REMOVE, L"削除");
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

// -------------------- Search thread --------------------
static DWORD WINAPI SearchThreadProc(LPVOID) {
    g_stopRequested = false;
    g_searching = true;

    auto rootItems = SplitRootsText(GetWindowTextWStr(g_editRoot));
    if (rootItems.empty()) {
        PostMessageW(g_hwndMain, WM_APP_FINISHED, 0, 0);
        return 0;
    }
    std::vector<fs::path> roots;
    std::error_code ec;
    for (auto& s : rootItems) {
        fs::path r = NormalizePath(fs::path(s));
        if (fs::exists(r, ec) && fs::is_directory(r, ec)) roots.push_back(r);
    }
    if (roots.empty()) {
        PostMessageW(g_hwndMain, WM_APP_THREADERR, 0, (LPARAM)L"検索元フォルダが存在しません");
        PostMessageW(g_hwndMain, WM_APP_FINISHED, 0, 0);
        return 0;
    }
    const bool useFolderExcl = IsChecked(g_chkEnableFolderExcl);
    const bool useNameExcl = IsChecked(g_chkEnableNameExcl);

    // date range
    std::chrono::system_clock::time_point s, e;
    GetActiveDateRange(s, e);

    // time base
    TimeBase tb = GetTimeBase();

    unsigned long long scanned = 0;
    unsigned long long hits = 0;

    std::wstring lastDir;
    int dirNotifyCountdown = 0;

    for (const auto& root : roots) {
        if (g_stopRequested) break;
        fs::recursive_directory_iterator it(root, fs::directory_options::skip_permission_denied, ec);
        if (ec) continue;
        for (auto end = fs::recursive_directory_iterator(); it != end; ++it) {
            if (g_stopRequested) break;

            const auto& entry = *it;

            // Notify current scanning folder (throttled)
            if (dirNotifyCountdown-- <= 0) {
                dirNotifyCountdown = 200; // every ~200 entries
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
            if ((scanned % 300) == 0 && IsWindow(g_hwndMain)) {
                PostMessageW(g_hwndMain, WM_APP_PROGRESS, (WPARAM)scanned, (LPARAM)hits);
            }

            const auto p = entry.path();
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
            else { // Either
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

            hits++;
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

    SYSTEMTIME stFrom{}, stTo{};
    DateTime_GetSystemtime(g_dtpFrom, &stFrom);
    DateTime_GetSystemtime(g_dtpTo, &stTo);

    wchar_t a[32], b[32];
    _snwprintf_s(a, _TRUNCATE, L"%04u-%02u-%02u", stFrom.wYear, stFrom.wMonth, stFrom.wDay);
    _snwprintf_s(b, _TRUNCATE, L"%04u-%02u-%02u", stTo.wYear, stTo.wMonth, stTo.wDay);
    return std::wstring(a) + L" ～ " + b;
}

static void StartSearch() {
    if (g_searching) return;

    if (!AnyExtSelected()) {
        MessageBoxW(g_hwndMain, L"対象拡張子が1つも選択されていません。", L"確認", MB_OK | MB_ICONWARNING);
        return;
    }
    if (!g_listRoots || (int)SendMessageW(g_listRoots, LB_GETCOUNT, 0, 0) <= 0) {
        MessageBoxW(g_hwndMain, L"検索元フォルダが未設定です（[追加…]でフォルダを追加してください）。", L"確認", MB_OK | MB_ICONWARNING);
        return;
    }
    RebuildFileNameExcludeCache();

    ClearResultsUI();
    SetSearchingUi(true);

    TimeBase tb = GetTimeBase();
    SetListViewTimeHeader(tb);

    // Sync multi-root list to legacy edit so worker thread can read it safely
    SyncLegacyRootEditFromList();

    std::wstring modeText = GetModeTextForStatus();
    SetStatus(L"検索中（" + modeText + L" / " + TimeBaseText(tb) + L"）...");

    DWORD tid = 0;
    g_hThread = CreateThread(nullptr, 0, SearchThreadProc, nullptr, 0, &tid);
    if (!g_hThread) {
        SetSearchingUi(false);
        SetStatus(L"[ERROR] スレッド作成に失敗しました");
    }
}
static void StopSearch() {
    if (!g_searching) return;
    g_stopRequested = true;
    SetStatus(L"停止要求中...");
}

// -------------------- WndProc --------------------
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

        // statics
        g_staticRoot = CreateWindowW(L"STATIC", L"検索先:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        g_staticRootsHint = CreateWindowW(L"STATIC", L"（複数可）下のリストに追加してください", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_STATIC_ROOTS_HINT, g_hInst, nullptr);
        g_staticMode = CreateWindowW(L"STATIC", L"期間:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        g_staticDays = CreateWindowW(L"STATIC", L"過去N日:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        g_staticTimeBase = CreateWindowW(L"STATIC", L"日時:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        g_staticFrom = CreateWindowW(L"STATIC", L"開始:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        g_staticTo = CreateWindowW(L"STATIC", L"終了:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        g_staticFilter = CreateWindowW(L"STATIC", L"絞り込み:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);

        g_staticExclFolder = CreateWindowW(L"STATIC", L"除外フォルダ:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        g_staticExclPattern = CreateWindowW(L"STATIC", L"フォルダ名部分一致/ワイルドカード", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        g_staticExclName = CreateWindowW(L"STATIC", L"ファイル名除外:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);

        // Root controls
        g_editRoot = CreateWindowW(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
            0, 0, 0, 0, hwnd, (HMENU)IDC_EDIT_ROOT, g_hInst, nullptr);
        g_btnBrowseRoot = CreateWindowW(L"BUTTON", L"追加...", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_BROWSE_ROOT, g_hInst, nullptr);

        // Roots list (intuitive multi-folder)
        g_listRoots = CreateWindowW(L"LISTBOX", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | LBS_NOTIFY | WS_VSCROLL,
            0, 0, 0, 0, hwnd, (HMENU)IDC_LIST_ROOTS, g_hInst, nullptr);

        g_btnRootRemove = CreateWindowW(L"BUTTON", L"削除", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_ROOT_REMOVE, g_hInst, nullptr);
        g_btnRootUp = CreateWindowW(L"BUTTON", L"上へ", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_ROOT_UP, g_hInst, nullptr);
        g_btnRootDown = CreateWindowW(L"BUTTON", L"下へ", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_ROOT_DOWN, g_hInst, nullptr);

        // Mode controls
        g_cmbMode = CreateWindowW(WC_COMBOBOXW, L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
            0, 0, 0, 0, hwnd, (HMENU)IDC_CMB_MODE, g_hInst, nullptr);
        SendMessageW(g_cmbMode, CB_ADDSTRING, 0, (LPARAM)L"今日");
        SendMessageW(g_cmbMode, CB_ADDSTRING, 0, (LPARAM)L"過去N日");
        SendMessageW(g_cmbMode, CB_ADDSTRING, 0, (LPARAM)L"期間指定（カレンダー）");


        g_editDays = CreateWindowW(L"EDIT", L"3", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | ES_NUMBER,
            0, 0, 0, 0, hwnd, (HMENU)IDC_EDIT_DAYS, g_hInst, nullptr);

        // NEW: time base
        g_cmbTimeBase = CreateWindowW(WC_COMBOBOXW, L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
            0, 0, 0, 0, hwnd, (HMENU)IDC_CMB_TIMEBASE, g_hInst, nullptr);
        SendMessageW(g_cmbTimeBase, CB_ADDSTRING, 0, (LPARAM)L"更新日時");
        SendMessageW(g_cmbTimeBase, CB_ADDSTRING, 0, (LPARAM)L"作成日時");
        SendMessageW(g_cmbTimeBase, CB_ADDSTRING, 0, (LPARAM)L"更新OR作成");

        g_dtpFrom = CreateWindowW(DATETIMEPICK_CLASSW, L"", WS_CHILD | WS_VISIBLE | DTS_SHORTDATEFORMAT,
            0, 0, 0, 0, hwnd, (HMENU)IDC_DTP_FROM, g_hInst, nullptr);

        g_dtpTo = CreateWindowW(DATETIMEPICK_CLASSW, L"", WS_CHILD | WS_VISIBLE | DTS_SHORTDATEFORMAT,
            0, 0, 0, 0, hwnd, (HMENU)IDC_DTP_TO, g_hInst, nullptr);

        if (!g_hFontUi) {
            g_hFontUi = CreateFontW(
                -18, 0, 0, 0, FW_MEDIUM, FALSE, FALSE, FALSE,
                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
        }
        if (!g_hFontUiBold) {
            g_hFontUiBold = CreateFontW(
                -20, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
        }

        HFONT hUiFont = g_hFontUi ? g_hFontUi : (HFONT)GetStockObject(DEFAULT_GUI_FONT);
        HFONT hUiFontBold = g_hFontUiBold ? g_hFontUiBold : hUiFont;
        SendMessageW(g_cmbMode, WM_SETFONT, (WPARAM)hUiFont, TRUE);
        SendMessageW(g_cmbTimeBase, WM_SETFONT, (WPARAM)hUiFont, TRUE);
        SendMessageW(g_dtpFrom, WM_SETFONT, (WPARAM)hUiFont, TRUE);
        SendMessageW(g_dtpTo, WM_SETFONT, (WPARAM)hUiFont, TRUE);

        // 表示フォーマット（好みで変更可）
        DateTime_SetFormat(g_dtpFrom, L"yyyy/MM/dd");
        DateTime_SetFormat(g_dtpTo, L"yyyy/MM/dd");

        // デフォルトは今日
        SYSTEMTIME st{};
        GetLocalTime(&st);
        DateTime_SetSystemtime(g_dtpFrom, GDT_VALID, &st);
        DateTime_SetSystemtime(g_dtpTo, GDT_VALID, &st);

        // Left panel tab (Search / Excludes)
        g_tabLeft = CreateWindowExW(0, WC_TABCONTROLW, L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS,
            0, 0, 0, 0, hwnd, (HMENU)IDC_TAB_LEFT, g_hInst, nullptr);
        SendMessageW(g_tabLeft, WM_SETFONT, (WPARAM)hUiFont, TRUE);
        {
            TCITEMW ti{};
            ti.mask = TCIF_TEXT;
            ti.pszText = const_cast<LPWSTR>(L"検索");
            TabCtrl_InsertItem(g_tabLeft, 0, &ti);
            ti.pszText = const_cast<LPWSTR>(L"除外");
            TabCtrl_InsertItem(g_tabLeft, 1, &ti);
            TabCtrl_SetCurSel(g_tabLeft, 0);
            TabCtrl_SetItemSize(g_tabLeft, 0, MAKELPARAM(96, 30));
        }

        // Advanced frames (behind controls) - used to visually separate sections
        g_frameFolderExcl = CreateWindowW(L"BUTTON", L"フォルダ除外", WS_CHILD | BS_GROUPBOX,
            0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);
        g_frameNameExcl = CreateWindowW(L"BUTTON", L"ファイル名除外", WS_CHILD | BS_GROUPBOX,
            0, 0, 0, 0, hwnd, nullptr, g_hInst, nullptr);

        // Folder exclude controls
        g_chkEnableFolderExcl = CreateWindowW(L"BUTTON", L"除外フォルダを有効", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
            0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_ENABLE_FOLDER_EXCL, g_hInst, nullptr);

        g_listExcludes = CreateWindowW(L"LISTBOX", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | LBS_NOTIFY | WS_VSCROLL,
            0, 0, 0, 0, hwnd, (HMENU)IDC_LIST_EXCLUDES, g_hInst, nullptr);

        g_btnAddExclFolder = CreateWindowW(L"BUTTON", L"フォルダ追加", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_ADD_EXCL_FOLDER, g_hInst, nullptr);
        g_btnRemoveExcl = CreateWindowW(L"BUTTON", L"削除", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_REMOVE_EXCL, g_hInst, nullptr);
        g_btnExclUp = CreateWindowW(L"BUTTON", L"上へ", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_EXCL_UP, g_hInst, nullptr);
        g_btnExclDown = CreateWindowW(L"BUTTON", L"下へ", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_EXCL_DOWN, g_hInst, nullptr);
        g_btnLoadExcl = CreateWindowW(L"BUTTON", L"読込", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_LOAD_EXCL, g_hInst, nullptr);
        g_btnSaveExcl = CreateWindowW(L"BUTTON", L"保存", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_SAVE_EXCL, g_hInst, nullptr);

        g_editExclPattern = CreateWindowW(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, hwnd, (HMENU)IDC_EDIT_EXCL_PATTERN, g_hInst, nullptr);
        g_btnAddPattern = CreateWindowW(L"BUTTON", L"追加", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_ADD_PATTERN, g_hInst, nullptr);
        // 編集欄: Enter で即更新
        g_oldExclEditProc = (WNDPROC)SetWindowLongPtrW(g_editExclPattern, GWLP_WNDPROC, (LONG_PTR)ExclEditProc);

        // Name excludes
        g_chkEnableNameExcl = CreateWindowW(L"BUTTON", L"ファイル名除外を有効", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_ENABLE_NAME_EXCL, g_hInst, nullptr);
        g_chkNameIncludeExt = CreateWindowW(L"BUTTON", L"拡張子を含めて判定", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_NAME_INCLUDE_EXT, g_hInst, nullptr);
        g_editFNamePattern = CreateWindowW(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, hwnd, (HMENU)IDC_EDIT_FNAME_PATTERN, g_hInst, nullptr);
        g_btnAddFName = CreateWindowW(L"BUTTON", L"追加", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_ADD_FNAME, g_hInst, nullptr);
        g_btnRemoveFName = CreateWindowW(L"BUTTON", L"削除", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_REMOVE_FNAME, g_hInst, nullptr);
        g_oldFNameEditProc = (WNDPROC)SetWindowLongPtrW(g_editFNamePattern, GWLP_WNDPROC, (LONG_PTR)FNameEditProc);
        g_btnFNameUp = CreateWindowW(L"BUTTON", L"上へ", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_FNAME_UP, g_hInst, nullptr);
        g_btnFNameDown = CreateWindowW(L"BUTTON", L"下へ", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_FNAME_DOWN, g_hInst, nullptr);
        g_listFName = CreateWindowW(L"LISTBOX", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | LBS_NOTIFY | WS_VSCROLL, 0, 0, 0, 0, hwnd, (HMENU)IDC_LIST_FNAME, g_hInst, nullptr);

        g_btnLoadFNameExcl = CreateWindowW(L"BUTTON", L"読込", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_LOAD_FNAME_EXCL, g_hInst, nullptr);
        g_btnSaveFNameExcl = CreateWindowW(L"BUTTON", L"保存", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_SAVE_FNAME_EXCL, g_hInst, nullptr);

        // Extensions group
        CreateWindowW(L"BUTTON", L"対象拡張子", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 0, 0, 0, 0, hwnd, (HMENU)IDC_GRP_EXT, g_hInst, nullptr);

        g_chkXls = CreateWindowW(L"BUTTON", L".xls", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_XLS, g_hInst, nullptr);
        g_chkXlsx = CreateWindowW(L"BUTTON", L".xlsx", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_XLSX, g_hInst, nullptr);
        g_chkXlsm = CreateWindowW(L"BUTTON", L".xlsm", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_XLSM, g_hInst, nullptr);
        g_chkXlsb = CreateWindowW(L"BUTTON", L".xlsb", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_XLSB, g_hInst, nullptr);
        g_chkXltx = CreateWindowW(L"BUTTON", L".xltx", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_XLTX, g_hInst, nullptr);
        g_chkXltm = CreateWindowW(L"BUTTON", L".xltm", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)IDC_CHK_XLTM, g_hInst, nullptr);

        // Actions
        g_btnSearch = CreateWindowW(L"BUTTON", L"検索", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_SEARCH, g_hInst, nullptr);
        g_btnStop = CreateWindowW(L"BUTTON", L"停止", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_STOP, g_hInst, nullptr);
        g_btnExportCsv = CreateWindowW(L"BUTTON", L"CSV出力...", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_BTN_EXPORT_CSV, g_hInst, nullptr);

        // 初期は空(0%)表示。検索中だけ Marquee を有効化する
        g_progress = CreateWindowW(PROGRESS_CLASSW, nullptr,
            WS_CHILD | WS_VISIBLE | PBS_SMOOTH,
            0, 0, 0, 0, hwnd, (HMENU)IDC_PROGRESS, g_hInst, nullptr);
        Progress_SetMarquee(g_progress, false);

        // Result filter (NEW)
        g_editFilter = CreateWindowW(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
            0, 0, 0, 0, hwnd, (HMENU)IDC_EDIT_FILTER, g_hInst, nullptr);
        SendMessageW(g_editFilter, EM_SETCUEBANNER, TRUE, (LPARAM)L"例: 入出荷 / 工場 / 2026");


        // Results
        g_listResults = CreateWindowW(WC_LISTVIEWW, L"", WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT, 0, 0, 0, 0, hwnd, (HMENU)IDC_LIST_RESULTS, g_hInst, nullptr);
        InitListViewColumns(g_listResults);

        // Status
        g_status = CreateWindowW(STATUSCLASSNAMEW, L"", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_STATUS, g_hInst, nullptr);

        HWND controls[] = {
            g_staticRoot, g_staticRootsHint, g_staticMode, g_staticDays, g_staticTimeBase, g_staticFrom, g_staticTo, g_staticFilter,
            g_staticExclFolder, g_staticExclPattern, g_staticExclName,
            g_editRoot, g_btnBrowseRoot, g_listRoots, g_btnRootRemove, g_btnRootUp, g_btnRootDown,
            g_cmbMode, g_editDays, g_cmbTimeBase, g_dtpFrom, g_dtpTo, g_tabLeft,
            g_frameFolderExcl, g_frameNameExcl,
            g_chkEnableFolderExcl, g_listExcludes, g_btnAddExclFolder, g_btnRemoveExcl, g_btnExclUp, g_btnExclDown, g_btnLoadExcl, g_btnSaveExcl,
            g_editExclPattern, g_btnAddPattern,
            g_chkEnableNameExcl, g_chkNameIncludeExt, g_editFNamePattern, g_btnAddFName, g_btnRemoveFName, g_btnFNameUp, g_btnFNameDown, g_listFName, g_btnLoadFNameExcl, g_btnSaveFNameExcl,
            GetDlgItem(hwnd, IDC_GRP_EXT), g_chkXls, g_chkXlsx, g_chkXlsm, g_chkXlsb, g_chkXltx, g_chkXltm,
            g_btnSearch, g_btnStop, g_btnExportCsv, g_progress, g_editFilter, g_listResults, g_status
        };
        for (HWND h : controls) {
            if (h) SendMessageW(h, WM_SETFONT, (WPARAM)hUiFont, TRUE);
        }
        SendMessageW(g_btnSearch, WM_SETFONT, (WPARAM)hUiFontBold, TRUE);

        // Paths: settings.ini / exclude.txt / results.csv are stored under LocalAppData to avoid server permission issues
        InitPaths();
        EnsureUnicodeIniWithMigration();

        // Load settings
        LoadSettings();

        // legacy single-root edit (not used in multi-root UI)
        ShowWindow(g_editRoot, SW_HIDE);

        EnableWindow(g_btnStop, FALSE);
        EnableWindow(g_btnExportCsv, FALSE);

        SetSearchingUi(false);
        SetStatus(L"待機中（デフォルト: 今日 / 更新日時）");

        // Ensure initial layout is applied even before the first WM_SIZE.
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

    case WM_COMMAND:
    {
        int id = LOWORD(wParam);
        int code = HIWORD(wParam);

        // 更新ボタン廃止: 編集欄から直接更新（Enter or フォーカスアウト）
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
                SyncLegacyRootEditFromList();
            }
            return 0;
        }

        if (id == IDC_BTN_ROOT_REMOVE) {
            int sel = (int)SendMessageW(g_listRoots, LB_GETCURSEL, 0, 0);
            if (sel != LB_ERR) {
                SendMessageW(g_listRoots, LB_DELETESTRING, (WPARAM)sel, 0);
                SyncLegacyRootEditFromList();
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
            SyncLegacyRootEditFromList();
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

        if (id == IDC_CHK_ENABLE_FOLDER_EXCL || id == IDC_CHK_ENABLE_NAME_EXCL) {
            UpdateUiEnableStates();
            return 0;
        }

        if (id == IDC_BTN_ADD_EXCL_FOLDER) {
            std::wstring p;
            if (PickFolder(hwnd, p)) AddExcludeDirPrefix(p);
            return 0;
        }

        if (id == IDC_LIST_EXCLUDES && (code == LBN_SELCHANGE || code == LBN_DBLCLK)) {
            int sel = (int)SendMessageW(g_listExcludes, LB_GETCURSEL, 0, 0);
            if (sel != LB_ERR && sel >= 0 && sel < (int)g_excludeRules.size()) {
                auto& r = g_excludeRules[(size_t)sel];
                // どの種類でも編集欄に出す（[DIR] も直接更新できるように）
                SetWindowTextWStr(g_editExclPattern, r.raw);
                if (code == LBN_DBLCLK) {
                    SetFocus(g_editExclPattern);
                    SendMessageW(g_editExclPattern, EM_SETSEL, 0, -1);
                }
            }
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

        // file name selection -> edit（ダブルクリックで編集欄へフォーカス）
        if (id == IDC_LIST_FNAME && (code == LBN_SELCHANGE || code == LBN_DBLCLK)) {
            int sel = (int)SendMessageW(g_listFName, LB_GETCURSEL, 0, 0);
            if (sel != LB_ERR && sel >= 0 && sel < (int)g_fileNamePatterns.size()) {
                SetWindowTextWStr(g_editFNamePattern, g_fileNamePatterns[(size_t)sel]);
                if (code == LBN_DBLCLK) {
                    SetFocus(g_editFNamePattern);
                    SendMessageW(g_editFNamePattern, EM_SETSEL, 0, -1);
                }
            }
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

        if (id == IDC_BTN_SEARCH) { StartSearch(); return 0; }
        if (id == IDC_BTN_STOP) { StopSearch();  return 0; }

        return 0;
    }

    case WM_NOTIFY:
    {
        LPNMHDR hdr = (LPNMHDR)lParam;
        if (!hdr) break;

        if (hdr->hwndFrom == g_tabLeft) {
            if (hdr->code == TCN_SELCHANGE) {
                int sel = TabCtrl_GetCurSel(g_tabLeft);
                SetLeftTab(sel);
                return 0;
            }
        }

        if (hdr->hwndFrom == g_listResults) {
            if (hdr->code == NM_DBLCLK) {
                int sel = ListView_GetNextItem(g_listResults, -1, LVNI_SELECTED);
                if (sel >= 0 && sel < (int)g_results.size()) {
                    ShellExecuteW(hwnd, L"open", g_results[(size_t)sel]->path.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
                }
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

    case WM_APP_PROGRESS:
    {
        unsigned long long scanned = (unsigned long long)wParam;
        unsigned long long hits = (unsigned long long)lParam;

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

        std::wstring s = L"完了：条件一致 " + std::to_wstring((int)g_results.size()) + L" 件";
        if (g_stopRequested) s += L"（途中停止）";
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
        SaveSettings();
        if (g_hFontUi) { DeleteObject(g_hFontUi); g_hFontUi = nullptr; }
        if (g_hFontUiBold) { DeleteObject(g_hFontUiBold); g_hFontUiBold = nullptr; }
        if (GetParent(hwnd) == nullptr) {
            PostQuitMessage(0);
        }
        return 0;
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}


// -------------------- Page exports --------------------
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
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);

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
        WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN,
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
