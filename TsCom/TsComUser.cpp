// TsComUser.cpp : Implementation of CTsComUser

#include "stdafx.h"
#include "TsComUser.h"

//const IID LIBID_TsComLib;
#ifndef MIDL_DEFINE_GUID
#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}
#endif

//__declspec( dllexport )
//MIDL_DEFINE_GUID(IID, LIBID_TsComLib,0x532C9128,0x3727,0x4A43,0x96,0xCC,0xA6,0x0C,0x51,0x2E,0xB3,0x62);

// CTsComUser

STDMETHODIMP CTsComUser::ConnectTS(void)
{
    HRESULT hr = CoCreateInstance( CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER, IID_ITaskService, (void**)&_pTaskService );  

    if (FAILED(hr))
    {
        LogComError(hr, _bstr_t("failed to create Task Service instance."));
        return hr;
    }

    hr = _pTaskService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t());

    if (FAILED(hr))
    {
        Cleanup();
        LogComError(hr, _bstr_t("failed to connect to Task Service."));
        return hr;
    }

    return hr;  // S_OK;
}


STDMETHODIMP CTsComUser::CreateTask(BSTR taskName, BSTR startTime)
{
    if (_pTaskService == NULL)
        return E_POINTER;

    ITaskFolder* pRootFolder = NULL;
    HRESULT hr = _pTaskService->GetFolder( _bstr_t( L"\\") , &pRootFolder );
    if( FAILED(hr) )
    {
        Cleanup();
        LogComError(hr, _bstr_t("failed to get root folder from the task service."));
        return hr;
    }

    ITaskDefinition* pTaskDef = NULL;
    hr = _pTaskService->NewTask(0, &pTaskDef);
    if (FAILED(hr))
    {
        Cleanup();
        LogComError(hr, _bstr_t("failed to get the Task Definition."));
        return hr;
    }

    //  Get the trigger collection to insert the time trigger.
    ITriggerCollection *pTriggerCollection = NULL;
    hr = pTaskDef->get_Triggers( &pTriggerCollection );
    if( FAILED(hr) )
    {
        printf("\nCannot get trigger collection: %x", hr );
        pRootFolder->Release();
        pTaskDef->Release();
        return hr;
    }

    //  Add the time trigger to the task.
    ITrigger *pTrigger = NULL;    
    hr = pTriggerCollection->Create( TASK_TRIGGER_TIME, &pTrigger );     
    pTriggerCollection->Release();
    if( FAILED(hr) )
    {
        printf("\nCannot create trigger: %x", hr );
        pRootFolder->Release();
        pTaskDef->Release();
        return hr;
    }

    ITimeTrigger *pTimeTrigger = NULL;
    hr = pTrigger->QueryInterface( 
        IID_ITimeTrigger, (void**) &pTimeTrigger );
    pTrigger->Release();
    if( FAILED(hr) )
    {
        printf("\nQueryInterface call failed for ITimeTrigger: %x", hr );
        pRootFolder->Release();
        pTaskDef->Release();
        return hr;
    }

    hr = pTimeTrigger->put_Id( _bstr_t( L"Trigger1" ) );
    if( FAILED(hr) )
        printf("\nCannot put trigger ID: %x", hr);

    //  Set the task to start at a certain time. The time 
    //  format should be YYYY-MM-DDTHH:MM:SS(+-)(timezone).
    //  For example, the start boundary below
    //  is January 1st 2005 at 12:05
    hr = pTimeTrigger->put_StartBoundary( _bstr_t(startTime) );
    pTimeTrigger->Release();    
    if( FAILED(hr) )
    {
        printf("\nCannot add start boundary to trigger: %x", hr );
        pRootFolder->Release();
        pTaskDef->Release();
        return hr;
    }

    //  Add an Action to the task. This task will execute notepad.exe.     
    IActionCollection *pActionCollection = NULL;

    //  Get the task action collection pointer.
    hr = pTaskDef->get_Actions( &pActionCollection );
    if( FAILED(hr) )
    {
        printf("\nCannot get Task collection pointer: %x", hr );
        pRootFolder->Release();
        pTaskDef->Release();
        return hr;
    }
    
    //  Create the action, specifying that it is an executable action.
    IAction *pAction = NULL;
    hr = pActionCollection->Create( TASK_ACTION_EXEC, &pAction );
    pActionCollection->Release();
    if( FAILED(hr) )
    {
        printf("\nCannot create action: %x", hr );
        pRootFolder->Release();
        pTaskDef->Release();
        return hr;
    }

    IExecAction *pExecAction = NULL;

    //  QI for the executable task pointer.
    hr = pAction->QueryInterface( 
        IID_IExecAction, (void**) &pExecAction );
    pAction->Release();
    if( FAILED(hr) )
    {
        printf("\nQueryInterface call failed for IExecAction: %x", hr );
        pRootFolder->Release();
        pTaskDef->Release();
        return hr;
    }

    //  Get the windows directory and set the path to notepad.exe.
    _bstr_t wstrExecutablePath = getenv("WINDIR");
    wstrExecutablePath += "\\SYSTEM32\\NOTEPAD.EXE";

    //  Set the path of the executable to notepad.exe.
    hr = pExecAction->put_Path(wstrExecutablePath); 
    pExecAction->Release(); 
    if( FAILED(hr) )
    {
        printf("\nCannot put the action executable path: %x", hr );
        pRootFolder->Release();
        pTaskDef->Release();
        return hr;
    }
    
    //  ------------------------------------------------------
    //  Save the task in the root folder.
    IRegisteredTask *pRegisteredTask = NULL;
    hr = pRootFolder->RegisterTaskDefinition(
            taskName,
            pTaskDef,
            TASK_CREATE_OR_UPDATE, 
            _variant_t(), 
            _variant_t(), 
            TASK_LOGON_INTERACTIVE_TOKEN,
            _variant_t(L""),
            &pRegisteredTask);
    if( FAILED(hr) )
    {
        printf("\nError saving the Task : %x", hr );
        pRootFolder->Release();
        pTaskDef->Release();
        return hr;
    }
    return S_OK;
}


STDMETHODIMP CTsComUser::IsTaskScheduled(BSTR taskName, VARIANT_BOOL* isScheduled)
{
    if (_pTaskService == NULL)
        return E_POINTER;

    if (isScheduled == NULL)
        return E_INVALIDARG;

    *isScheduled = VARIANT_FALSE;
    _bstr_t expectedName(taskName, false);

    ITaskFolder *pRootFolder = NULL;
    HRESULT hr = _pTaskService->GetFolder( _bstr_t( L"\\") , &pRootFolder );
    
    _pTaskService->Release();
    if( FAILED(hr) )
    {
        printf("Cannot get Root Folder pointer: %x", hr );
        return hr;
    }
    
    IRegisteredTaskCollection* pTaskCollection = NULL;
    hr = pRootFolder->GetTasks( NULL, &pTaskCollection );

    pRootFolder->Release();
    if( FAILED(hr) )
    {
        printf("Cannot get the registered tasks.: %x", hr);
        return hr;
    }

    LONG numTasks = 0;
    hr = pTaskCollection->get_Count(&numTasks);
    if( numTasks == 0 )
    {
        printf("\nNo Tasks are currently running" );
        pTaskCollection->Release();
        return hr;
    }

    for (LONG i=0; i < numTasks; i++)
    {
        IRegisteredTask* pRegisteredTask = NULL;
        hr = pTaskCollection->get_Item( _variant_t(i+1), &pRegisteredTask );
        
        if (SUCCEEDED(hr))
        {
            BSTR tName = NULL;
            hr = pRegisteredTask->get_Name(&tName);
            if (SUCCEEDED(hr))
            {
                _bstr_t foundName = tName;
                SysFreeString(tName);

                if (foundName == expectedName)
                {
                    *isScheduled = VARIANT_TRUE;
                    break;
                }
            }
            pRegisteredTask->Release();
        }
    }

    pTaskCollection->Release();
    return S_OK;
}

void CTsComUser::LogComError(HRESULT hr, _bstr_t message)
{
    // log the error to a file, and/or output to the console if there is one.
    printf("\nHRESULT: %d, %s", hr, (const char*)message);
}

