#include <iostream>
#include <csignal>
#include <comutil.h>
#include <Windows.h>
#include <WbemIdl.h>
#include <tlhelp32.h>

#include "EventSink.h"

using namespace std;


// ====================
//     DECLARATIONS
// ====================

HANDLE hMutex;

HWND hSpotifyMainWindow;


void FixSpotifyTaskbarIssue(DWORD);
RECT GetWindowPosition(HWND);
BOOL IsSpotifyMainProcess(DWORD);


inline void HideConsole() { ShowWindow(GetConsoleWindow(), SW_HIDE); }
inline void ShowConsole() { ShowWindow(GetConsoleWindow(), SW_SHOW); }
inline bool IsConsoleVisible() { return IsWindowVisible(GetConsoleWindow()) != FALSE; }

void Abort(const int signum)
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
    cout << "\nPress enter to continue...";
    getchar();
}


// ============
//     MAIN
// ============

int main()  // NOLINT
{
    HideConsole();

    // Signal handling
    signal(SIGABRT, Abort);
    signal(SIGINT, Abort);
    signal(SIGTERM, Abort);

    // Mutex
    hMutex = CreateMutex(nullptr, TRUE, L"SpotifyTaskbarFix-{150F0728-6840-4C9D-B2EE-DE289EAFE29F}");
    if (GetLastError() == ERROR_ALREADY_EXISTS)
    {
        ShowConsole();
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
            ShowConsole();
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

            ShowConsole();
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

            ShowConsole();
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

            ShowConsole();
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

            ShowConsole();
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

            ShowConsole();
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

            ShowConsole();
            cout << "ExecNotificationQueryAsync failed with error: 0x" << hex << hres << endl;
            ReadLine();
            return 1;
        }

        while (true)
        {
            Sleep(1500);
        }
    }
    catch (const std::exception& e)
    {
        ShowConsole();
        cout << "ERROR: Unhandled exception\n    " << e.what() << endl;
        ReadLine();
        Abort(1);
    }
}


// =================
//     FUNCTIONS
// =================

void FixSpotifyTaskbarIssue(const DWORD processId)
{
    Sleep(120);
    hSpotifyMainWindow = nullptr;
    if (IsSpotifyMainProcess(processId) == TRUE)
    {
        const HWND hWnd = hSpotifyMainWindow;
        if (hWnd != nullptr)
        {
            const HANDLE hProcess = OpenProcess(PROCESS_ALL_ACCESS, FALSE, processId);
            WaitForInputIdle(hProcess, 1000);
            CloseHandle(hProcess);

            const RECT wPos = GetWindowPosition(hWnd);
            const SIZE wSz = {wPos.right - wPos.left, wPos.bottom - wPos.top};
            const RECT zero = {0, 0, wSz.cx, wSz.cy};

            Sleep(200);
            MoveWindow(hWnd, zero.left, zero.top, wSz.cx, wSz.cy, TRUE);
            Sleep(50);
            MoveWindow(hWnd, wPos.left, wPos.top, wSz.cx, wSz.cy, TRUE);
        }
    }
}

RECT GetWindowPosition(const HWND hWnd)
{
    RECT rect;
    GetWindowRect(hWnd, &rect);
    MapWindowPoints(HWND_DESKTOP, GetParent(hWnd), reinterpret_cast<LPPOINT>(&rect), 2);
    return rect;
}

BOOL IsSpotifyMainProcess(const DWORD processId)
{
    BOOL hasMainWindow = FALSE;

    // Enumerate threads
    const HANDLE hSnapshotThread = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, NULL);
    if (hSnapshotThread != INVALID_HANDLE_VALUE)
    {
        THREADENTRY32 thread;
        thread.dwSize = sizeof thread;

        if (Thread32First(hSnapshotThread, &thread))
        {
            do
            {
                if (thread.th32OwnerProcessID != processId)
                    continue;
                if (thread.dwSize < FIELD_OFFSET(THREADENTRY32, th32OwnerProcessID) + sizeof thread.th32OwnerProcessID)
                    continue;

                hasMainWindow = !EnumThreadWindows(thread.th32ThreadID, [](const HWND hWnd, const LPARAM) -> BOOL
                {
                    if (hWnd == INVALID_HANDLE_VALUE)
                        return true;

                    WCHAR title[256], className[256];
                    GetWindowText(hWnd, title, 256);
                    GetClassName(hWnd, className, 256);

                    if (_wcsicmp(className, L"Chrome_WidgetWin_0") == 0 && _wcsicmp(title, L"") != 0)
                    {
                        // Main window found!
                        hSpotifyMainWindow = hWnd;
                        return false;
                    }
                    return true;
                }, NULL);
            }
            while (hasMainWindow == FALSE && Thread32Next(hSnapshotThread, &thread));
        }
        CloseHandle(hSnapshotThread);
    }

    return hasMainWindow;
}
