#include "EventSink.h"

ULONG EventSink::AddRef()
{
    return InterlockedIncrement(&m_lRef);
}

ULONG EventSink::Release()
{
    const LONG lRef = InterlockedDecrement(&m_lRef);
    if (lRef == 0)
        delete this;
    return lRef;
}

HRESULT EventSink::QueryInterface(REFIID riid, void** ppv)
{
    if (riid == IID_IUnknown || riid == IID_IWbemObjectSink)
    {
        *ppv = static_cast<IWbemObjectSink*>(this);
        AddRef();
        return WBEM_S_NO_ERROR;
    }
    return E_NOINTERFACE;
}


HRESULT EventSink::Indicate(const long lObjectCount, IWbemClassObject** apObjArray)
{
    HRESULT hres = S_OK;

    for (int i = 0; i < lObjectCount; i++)
    {
        VARIANT v;
        if (SUCCEEDED(apObjArray[i]->Get(L"ProcessID", NULL, &v, nullptr, nullptr)))
            callback(static_cast<DWORD>(V_UINT(&v)));
    }

    return WBEM_S_NO_ERROR;
}

HRESULT EventSink::SetStatus(const LONG lFlags, const HRESULT hResult, BSTR strParam, IWbemClassObject __RPC_FAR* pObjParam)
{
    return WBEM_S_NO_ERROR;
}
