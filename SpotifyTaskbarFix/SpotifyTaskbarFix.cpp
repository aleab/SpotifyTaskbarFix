#include <chrono>
#include <csignal>
#include <iostream>
#include <iomanip>
#include <string>

#include <Windows.h>
#include <WbemIdl.h>
#include <comutil.h>
#include <tlhelp32.h>

#include "EventSink.h"

using namespace std;
using namespace std::chrono;


typedef struct {
    HWND hPWnd;
    HWND hWnd;
    bool isMainProcess;
} FindSpotifyMainWindowResult;


constexpr auto WAIT_FOR_INPUT_IDLE_TIMEOUT = 1 * 1000;
constexpr auto FIND_SPOTIFY_WINDOW_RETRY_TIMEOUT = 5 * 1000;
constexpr auto WINDOW_VISIBLE_TIMEOUT = 10 * 1000;


// ====================
//     DECLARATIONS
// ====================

// GLOBAL VARIABLES
HANDLE hMutex;
HWND hConsoleWindow;
bool showConsole = false;

// FUNCTIONS
void FixSpotifyTaskbarIssue(DWORD);
void FindSpotifyMainWindow(const DWORD, FindSpotifyMainWindowResult* const);

void GetLocalTime(const system_clock::time_point* const, char* const, size_t);
RECT GetWindowPosition(const HWND);
void PrintWindowStyles(const LONG, const LONG, const LONG, const LONG);

// INLINE FUNCTIONS
inline void HideConsole()
{
    if (hConsoleWindow != nullptr)
        ShowWindow(hConsoleWindow, SW_HIDE);
}

inline void ShowConsole()
{
    if (hConsoleWindow != nullptr)
        ShowWindow(hConsoleWindow, SW_SHOW);
}

inline void Abort(const int signum)
{
    if (hMutex != nullptr && hMutex != INVALID_HANDLE_VALUE)
    {
        ReleaseMutex(hMutex);
        CloseHandle(hMutex);
    }

    exit(signum);
}

inline void ReadLine()
{
    ShowConsole();
    cout << "\nPress enter to continue...";
    getchar();
}

inline bool HasFlag(LONG bitfield, LONG flag)
{
    return (bitfield & flag) != 0;
}


// ============
//     MAIN
// ============

int main(const int argc, char* argv[]) // NOLINT
{
    // Parse arguments
    for (int i = 0; i < argc; i++)
    {
        if (strcmp(argv[i], "-w") == 0)
            showConsole = true;
    }

    hConsoleWindow = GetConsoleWindow();
    if (!showConsole)
        HideConsole();

    // Signal handling
    signal(SIGABRT, Abort);
    signal(SIGINT, Abort);
    signal(SIGTERM, Abort);

    // Mutex
    hMutex = CreateMutex(nullptr, TRUE, L"SpotifyTaskbarFix-{150F0728-6840-4C9D-B2EE-DE289EAFE29F}");
    if (GetLastError() == ERROR_ALREADY_EXISTS)
    {
        cout << "ERROR: Program is already running!" << endl;
        ReadLine();
        return 1;
    }

    try
    {
        // Initialize COM
        HRESULT hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
        if (FAILED(hres))
        {
            cout << "Failed to initialize COM library. Error code = 0x" << hex << hres << endl;
            ReadLine();
            return 1;
        }

        // Set general COM security levels
        hres = CoInitializeSecurity(
            nullptr,
            -1,
            nullptr,
            nullptr,
            RPC_C_AUTHN_LEVEL_DEFAULT,
            RPC_C_IMP_LEVEL_IMPERSONATE,
            nullptr,
            EOAC_NONE,
            nullptr);
        if (FAILED(hres))
        {
            CoUninitialize();

            cout << "Failed to initialize security. Error code = 0x" << hex << hres << endl;
            ReadLine();
            return 1;
        }

        // Obtain the initial locator to WMI
        IWbemLocator* pLoc = nullptr;
        hres = CoCreateInstance(CLSID_WbemLocator, nullptr, CLSCTX_INPROC_SERVER, IID_IWbemLocator,
                                reinterpret_cast<LPVOID*>(&pLoc));
        if (FAILED(hres))
        {
            CoUninitialize();

            cout << "Failed to create IWbemLocator object. Error code = 0x" << hex << hres << endl;
            ReadLine();
            return 1;
        }

        // Connect to WMI through the IWbemLocator::ConnectServer method
        IWbemServices* pSvc = nullptr;
        hres = pLoc->ConnectServer(_bstr_t(L"ROOT\\CIMV2"), nullptr, nullptr, nullptr, NULL, nullptr, nullptr, &pSvc);
        if (FAILED(hres))
        {
            pLoc->Release();
            CoUninitialize();

            cout << "Could not connect. Error code = 0x" << hex << hres << endl;
            ReadLine();
            return 1;
        }

        // Set security levels on the proxy
        hres = CoSetProxyBlanket(
            pSvc,
            RPC_C_AUTHN_WINNT,
            RPC_C_AUTHZ_NONE,
            nullptr,
            RPC_C_AUTHN_LEVEL_CALL,
            RPC_C_IMP_LEVEL_IMPERSONATE,
            nullptr,
            EOAC_NONE);
        if (FAILED(hres))
        {
            pSvc->Release();
            pLoc->Release();
            CoUninitialize();

            cout << "Could not set proxy blanket. Error code = 0x" << hex << hres << endl;
            ReadLine();
            return 1;
        }

        // Receive event notifications
        IUnsecuredApartment* pUnsecApp = nullptr;
        hres = CoCreateInstance(CLSID_UnsecuredApartment, nullptr, CLSCTX_LOCAL_SERVER, IID_IUnsecuredApartment,
                                reinterpret_cast<void**>(&pUnsecApp));
        if (FAILED(hres))
        {
            pSvc->Release();
            pLoc->Release();
            pUnsecApp->Release();
            CoUninitialize();

            cout << "CoCreateInstance failed with error: 0x" << hex << hres << endl;
            ReadLine();
            return 1;
        }

        EventSink* pSink = new EventSink(FixSpotifyTaskbarIssue);
        pSink->AddRef();

        IUnknown* pStubUnk = nullptr;
        pUnsecApp->CreateObjectStub(pSink, &pStubUnk);

        IWbemObjectSink* pStubSink = nullptr;
        pStubUnk->QueryInterface(IID_IWbemObjectSink, reinterpret_cast<void**>(&pStubSink));

        // The ExecNotificationQueryAsync method will call the EventQuery::Indicate method when an event occurs
        hres = pSvc->ExecNotificationQueryAsync(
            _bstr_t("WQL"),
            _bstr_t("SELECT * FROM Win32_ProcessStartTrace WHERE ProcessName = \"Spotify.exe\""),
            WBEM_FLAG_SEND_STATUS, nullptr, pStubSink);
        if (FAILED(hres))
        {
            pSvc->Release();
            pLoc->Release();
            pUnsecApp->Release();
            pStubUnk->Release();
            pSink->Release();
            pStubSink->Release();
            CoUninitialize();

            cout << "ExecNotificationQueryAsync failed with error: 0x" << hex << hres << endl;
            ReadLine();
            return 1;
        }

        cout << "READY!" << endl;
        while (true)
        {
            Sleep(1500);
        }
    }
    catch (const std::exception& e)
    {
        cout << "ERROR: Unhandled exception\n    " << e.what() << endl;
        ReadLine();
        Abort(1);
    }
}

void FixSpotifyTaskbarIssue(const DWORD processId)
{
    const auto startedTime = system_clock::now();

    const HANDLE hProcess = OpenProcess(PROCESS_ALL_ACCESS, FALSE, processId);
    WaitForInputIdle(hProcess, WAIT_FOR_INPUT_IDLE_TIMEOUT);
    CloseHandle(hProcess);

    const auto idleTime = system_clock::now();
    const long long msToIdle = duration_cast<milliseconds>(idleTime.time_since_epoch() - startedTime.time_since_epoch()).count();

    if (msToIdle < 500)
        Sleep(500);

    FindSpotifyMainWindowResult spResult;
    FindSpotifyMainWindow(processId, &spResult);

    if (spResult.isMainProcess)
    {
        char localTime[32];
        GetLocalTime(&startedTime, localTime, sizeof localTime);
        printf("[%s] Spotify process started.\n", localTime);
        printf("  | Process ID:    0x%08lX\n", processId);

        // The main window might not be immediately available after the process starts;
        // in case the main process is correctly identified, but it doesn't YET contain
        // the main window, we wait for a bit and look for it again.

        long ms = 0;
        while ((spResult.hPWnd == nullptr || spResult.hWnd == nullptr) && ms < FIND_SPOTIFY_WINDOW_RETRY_TIMEOUT)
        {
            Sleep(250);
            FindSpotifyMainWindow(processId, &spResult);
            ms += 250;
        }

        if (spResult.hPWnd != nullptr && spResult.hWnd != nullptr)
        {
            const RECT wPos = GetWindowPosition(spResult.hPWnd);
            const SIZE wSz = { wPos.right - wPos.left, wPos.bottom - wPos.top };
            const RECT zero = { 0, 0, wSz.cx, wSz.cy };

            printf("  | Window handle: 0x%08lX\n", reinterpret_cast<unsigned long>(spResult.hWnd));
            printf("  | Window position: (%d, %d)\n", wPos.left, wPos.top);

            long ms = 0;
            //LONG prevStyles = 0;
            //LONG prevExStyles = 0;
            //while (ms < 5 * 1000)
            //{
            //    const LONG styles = GetWindowLongA(spResult.hWnd, GWL_STYLE);
            //    const LONG exStyles = GetWindowLongA(spResult.hWnd, GWL_EXSTYLE);
            //    PrintWindowStyles(styles, prevStyles, exStyles, prevExStyles);
            //    prevStyles = styles;
            //    prevExStyles = exStyles;

            //    Sleep(250); 
            //    ms += 250;
            //}
            bool validWindowStyle = false;
            while (!validWindowStyle && ms < WINDOW_VISIBLE_TIMEOUT)
            {
                Sleep(100);
                ms += 100;
                validWindowStyle = HasFlag(GetWindowLongA(spResult.hWnd, GWL_STYLE), WS_VISIBLE) && HasFlag(GetWindowLongA(spResult.hPWnd, GWL_STYLE), WS_VISIBLE | WS_SYSMENU);
            }

            if (validWindowStyle)
            {
                char localTime[32];

                const auto visibleTime = system_clock::now();
                GetLocalTime(&visibleTime, localTime, sizeof localTime);
                printf("[%s] The window is visible.\n", localTime);

                MoveWindow(spResult.hPWnd, zero.left, zero.top, wSz.cx, wSz.cy, TRUE);
                MoveWindow(spResult.hPWnd, wPos.left, wPos.top, wSz.cx, wSz.cy, TRUE);

                const auto doneTime = system_clock::now();
                GetLocalTime(&doneTime, localTime, sizeof localTime);
                printf("[%s] The window has been moved.\n\n", localTime);
            }
            else
            {
                cout << "No visible window found!" << endl;
            }
        }
        else
        {
            cout << "Spotify main window not found!" << endl;
        }
    }
}

BOOL FindSpotifyMainWindow_EnumChildWindows(const HWND hWnd, const LPARAM lparam)
{
    if (hWnd == INVALID_HANDLE_VALUE)
        return TRUE;

    auto result = (FindSpotifyMainWindowResult* const)lparam;

    WCHAR title[256], className[256];
    GetWindowText(hWnd, title, 256);
    GetClassName(hWnd, className, 256);

    //printf("##       0x%08lX: \"%ls\" \"%ls\"\n", hWnd, className, title);

    if (_wcsicmp(className, L"Chrome_RenderWidgetHostHWND") == 0)
    {
        result->hWnd = hWnd;
        return FALSE;
    }
    return TRUE;
}

BOOL FindSpotifyMainWindow_EnumThreadWindows(const HWND hWnd, const LPARAM lparam)
{
    if (hWnd == INVALID_HANDLE_VALUE)
        return TRUE;

    auto result = (FindSpotifyMainWindowResult* const)lparam;

    WCHAR title[256], className[256];
    GetWindowText(hWnd, title, 256);
    GetClassName(hWnd, className, 256);

    //printf("##    0x%08lX: \"%ls\" \"%ls\"\n", hWnd, className, title);

    if (_wcsicmp(className, L"GDI+ Hook Window Class") == 0 && _wcsicmp(title, L"G") == 0)
    {
        result->isMainProcess = true;
    }
    else if (_wcsicmp(className, L"Chrome_WidgetWin_0") == 0)
    {
        result->isMainProcess = true;
        if (_wcsicmp(title, L"") != 0)
        {
            // Enumerate child windows
            EnumChildWindows(hWnd, FindSpotifyMainWindow_EnumChildWindows, lparam);
            if (result->hWnd != nullptr)
            {
                result->hPWnd = hWnd;
                return FALSE;
            }
        }
    }
    return TRUE;
}

void FindSpotifyMainWindow(const DWORD processId, FindSpotifyMainWindowResult *const result)
{
    //printf("## Process ID: 0x%08lX\n", processId);
    result->hPWnd = nullptr;
    result->hWnd = nullptr;
    result->isMainProcess = false;

    const HANDLE hSnapshotThread = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, NULL);
    if (hSnapshotThread != INVALID_HANDLE_VALUE)
    {
        THREADENTRY32 thread;
        thread.dwSize = sizeof thread;

        // Enumerate this process's threads
        if (Thread32First(hSnapshotThread, &thread))
        {
            do
            {
                if (thread.th32OwnerProcessID != processId)
                    continue;
                if (thread.dwSize < FIELD_OFFSET(THREADENTRY32, th32OwnerProcessID) + sizeof thread.th32OwnerProcessID)
                    continue;

                // Enumerate this thread's windows
                EnumThreadWindows(thread.th32ThreadID, FindSpotifyMainWindow_EnumThreadWindows, (LPARAM)result);
            } while (Thread32Next(hSnapshotThread, &thread));
        }
        CloseHandle(hSnapshotThread);
    }
}


// =================
//     FUNCTIONS
// =================

void GetLocalTime(const system_clock::time_point *const tp, char *const buffer, size_t bufferLength)
{
    const auto time = system_clock::to_time_t(*tp);
    struct tm ptm;
    localtime_s(&ptm, &time);

    std::strftime(buffer, bufferLength, "%H:%M:%S", &ptm);
    const auto ms = static_cast<unsigned>((tp->time_since_epoch() - duration_cast<seconds>(tp->time_since_epoch())) / milliseconds(1));
    snprintf(buffer, bufferLength, "%s.%d", buffer, ms);
}

RECT GetWindowPosition(const HWND hWnd)
{
    RECT rect;
    GetWindowRect(hWnd, &rect);
    MapWindowPoints(HWND_DESKTOP, GetParent(hWnd), reinterpret_cast<LPPOINT>(&rect), 2);
    return rect;
}

void PrintWindowStyle(const LONG current, const LONG previous, LONG flag, const char* name)
{
    bool has = HasFlag(current, flag);
    bool had = HasFlag(previous, flag);

    if (has && !had)
        printf("  + %s\n", name);
    else if (!has && had)
        printf("  - %s\n", name);
}

void PrintWindowStyles(const LONG styles, const LONG prevStyles, const LONG exStyles, const LONG prevExStyles)
{
    printf("Window styles:\n");

    PrintWindowStyle(styles, prevStyles, WS_BORDER, "WS_BORDER");
    PrintWindowStyle(styles, prevStyles, WS_CAPTION, "WS_CAPTION");
    PrintWindowStyle(styles, prevStyles, WS_CHILD, "WS_CHILD");
    PrintWindowStyle(styles, prevStyles, WS_CHILDWINDOW, "WS_CHILDWINDOW");
    PrintWindowStyle(styles, prevStyles, WS_CLIPCHILDREN, "WS_CLIPCHILDREN");
    PrintWindowStyle(styles, prevStyles, WS_CLIPSIBLINGS, "WS_CLIPSIBLINGS");
    PrintWindowStyle(styles, prevStyles, WS_DISABLED, "WS_DISABLED");
    PrintWindowStyle(styles, prevStyles, WS_DLGFRAME, "WS_DLGFRAME");
    PrintWindowStyle(styles, prevStyles, WS_GROUP, "WS_GROUP");
    PrintWindowStyle(styles, prevStyles, WS_HSCROLL, "WS_HSCROLL");
    PrintWindowStyle(styles, prevStyles, WS_ICONIC, "WS_ICONIC");
    PrintWindowStyle(styles, prevStyles, WS_MAXIMIZE, "WS_MAXIMIZE");
    PrintWindowStyle(styles, prevStyles, WS_MAXIMIZEBOX, "WS_MAXIMIZEBOX");
    PrintWindowStyle(styles, prevStyles, WS_MINIMIZE, "WS_MINIMIZE");
    PrintWindowStyle(styles, prevStyles, WS_MINIMIZEBOX, "WS_MINIMIZEBOX");
    PrintWindowStyle(styles, prevStyles, WS_OVERLAPPED, "WS_OVERLAPPED");
    PrintWindowStyle(styles, prevStyles, WS_OVERLAPPEDWINDOW, "WS_OVERLAPPEDWINDOW");
    PrintWindowStyle(styles, prevStyles, WS_POPUP, "WS_POPUP");
    PrintWindowStyle(styles, prevStyles, WS_POPUPWINDOW, "WS_POPUPWINDOW");
    PrintWindowStyle(styles, prevStyles, WS_SIZEBOX, "WS_SIZEBOX");
    PrintWindowStyle(styles, prevStyles, WS_SYSMENU, "WS_SYSMENU");
    PrintWindowStyle(styles, prevStyles, WS_TABSTOP, "WS_TABSTOP");
    PrintWindowStyle(styles, prevStyles, WS_THICKFRAME, "WS_THICKFRAME");
    PrintWindowStyle(styles, prevStyles, WS_TILED, "WS_TILED");
    PrintWindowStyle(styles, prevStyles, WS_TABSTOP, "WS_TABSTOP");
    PrintWindowStyle(styles, prevStyles, WS_TILEDWINDOW, "WS_TILEDWINDOW");
    PrintWindowStyle(styles, prevStyles, WS_VISIBLE, "WS_VISIBLE");
    PrintWindowStyle(styles, prevStyles, WS_VSCROLL, "WS_VSCROLL");

    PrintWindowStyle(exStyles, prevExStyles, WS_EX_ACCEPTFILES, "WS_EX_ACCEPTFILES");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_APPWINDOW, "WS_EX_APPWINDOW");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_CLIENTEDGE, "WS_EX_CLIENTEDGE");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_COMPOSITED, "WS_EX_COMPOSITED");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_CONTEXTHELP, "WS_EX_CONTEXTHELP");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_CONTROLPARENT, "WS_EX_CONTROLPARENT");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_DLGMODALFRAME, "WS_EX_DLGMODALFRAME");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_LAYERED, "WS_EX_LAYERED");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_LAYOUTRTL, "WS_EX_LAYOUTRTL");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_LEFT, "WS_EX_LEFT");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_LEFTSCROLLBAR, "WS_EX_LEFTSCROLLBAR");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_LTRREADING, "WS_EX_LTRREADING");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_MDICHILD, "WS_EX_MDICHILD");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_NOACTIVATE, "WS_EX_NOACTIVATE");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_NOINHERITLAYOUT, "WS_EX_NOINHERITLAYOUT");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_NOPARENTNOTIFY, "WS_EX_NOPARENTNOTIFY");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_NOREDIRECTIONBITMAP, "WS_EX_NOREDIRECTIONBITMAP");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_OVERLAPPEDWINDOW, "WS_EX_OVERLAPPEDWINDOW");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_PALETTEWINDOW, "WS_EX_PALETTEWINDOW");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_RIGHT, "WS_EX_RIGHT");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_RIGHTSCROLLBAR, "WS_EX_RIGHTSCROLLBAR");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_RTLREADING, "WS_EX_RTLREADING");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_STATICEDGE, "WS_EX_STATICEDGE");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_TOOLWINDOW, "WS_EX_TOOLWINDOW");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_TOPMOST, "WS_EX_TOPMOST");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_TRANSPARENT, "WS_EX_TRANSPARENT");
    PrintWindowStyle(exStyles, prevExStyles, WS_EX_WINDOWEDGE, "WS_EX_WINDOWEDGE");
}
