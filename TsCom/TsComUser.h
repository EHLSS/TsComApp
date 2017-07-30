// TsComUser.h : Declaration of the CTsComUser

#pragma once
#include "resource.h"       // main symbols



#include "TsCom_i.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;


// CTsComUser
#define DllExport   __declspec( dllexport )

class DllExport ATL_NO_VTABLE CTsComUser :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CTsComUser, &CLSID_TsComUser>,
	public IDispatchImpl<ITsComUser, &IID_ITsComUser, &LIBID_TsComLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CTsComUser()
	{
        _pTaskService = NULL;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_TSCOMUSER)

BEGIN_COM_MAP(CTsComUser)
	COM_INTERFACE_ENTRY(ITsComUser)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
        Cleanup();
	}

public:

    STDMETHOD(ConnectTS)(void);
    STDMETHOD(CreateTask)(BSTR taskName, BSTR startTime);
    STDMETHOD(IsTaskScheduled)(BSTR taskName, VARIANT_BOOL* isScheduled);

private:
    ITaskService *_pTaskService;

    void Cleanup()
    {
        if (_pTaskService != NULL)
        {
            _pTaskService->Release();
            _pTaskService = NULL;
        }
    }

    void LogComError(HRESULT hr, _bstr_t message);
};

OBJECT_ENTRY_AUTO(__uuidof(TsComUser), CTsComUser)
