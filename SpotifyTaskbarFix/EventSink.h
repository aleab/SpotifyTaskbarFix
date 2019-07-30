#pragma once

#define _WIN32_DCOM
#include <iostream>
using namespace std;

#include <comdef.h>
#include <Wbemidl.h>

#pragma comment(lib, "wbemuuid.lib")

class EventSink final : public IWbemObjectSink
{
    LONG m_lRef;
    bool bDone;

    void (*callback)(DWORD);

protected:
    ~EventSink() { bDone = true; }

public:
    explicit EventSink(void (*fn)(DWORD))
    {
        m_lRef = 0;
        bDone = false;
        callback = fn;
    }

    ULONG STDMETHODCALLTYPE AddRef() override;
    ULONG STDMETHODCALLTYPE Release() override;
    HRESULT
        STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override;

    HRESULT STDMETHODCALLTYPE Indicate(
        LONG lObjectCount,
        IWbemClassObject __RPC_FAR* __RPC_FAR* apObjArray
    ) override;

    HRESULT STDMETHODCALLTYPE SetStatus(
        _In_ LONG lFlags,
        _In_ HRESULT hResult,
        _In_ BSTR strParam,
        _In_ IWbemClassObject __RPC_FAR* pObjParam
    ) override;
};
