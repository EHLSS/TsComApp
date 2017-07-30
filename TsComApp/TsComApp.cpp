// TsComApp.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#import "TsCom.tlb"
#include "..\TsCom\TsComUser.h"

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

MIDL_DEFINE_GUID(IID, IID_ITsComUser,0xEC7182B0,0xBB40,0x4D55,0x9F,0x42,0x98,0x0A,0x94,0x91,0xE9,0x7A);
MIDL_DEFINE_GUID(IID, LIBID_TsComLib,0x532C9128,0x3727,0x4A43,0x96,0xCC,0xA6,0x0C,0x51,0x2E,0xB3,0x62);
MIDL_DEFINE_GUID(CLSID, CLSID_TsComUser,0x46A42EBD,0xD7A6,0x459E,0xAE,0x61,0x05,0x7B,0x8A,0x2F,0x49,0x30);

HRESULT CreateScheduledTask()
{
    ITsComUser *pTsCom = NULL;
    HRESULT hr = CoCreateInstance( CLSID_TsComUser,
                           NULL,
                           CLSCTX_INPROC_SERVER,
                           IID_ITsComUser,
                           (void**)&pTsCom );  
    if (FAILED(hr))
    {
        printf("\nFailed to CoCreate an instance of the TsComUser class: %x", hr);
        return hr;
    }

    hr = pTsCom->ConnectTS();
    if (FAILED(hr))
    {
        printf("\nFailed to connect to the Task Scheduler: %x", hr);
        pTsCom->Release();
        return hr;
    }

    BSTR taskName = SysAllocString(L"PhishMe Scheduled Task Test");
    BSTR startTime = SysAllocString(L"2017-07-30T19:00:00");
    hr = pTsCom->CreateTask(taskName, startTime);
    if (FAILED(hr))
    {
          SysFreeString(taskName);
          SysFreeString(startTime);
        printf("\nFailed to create the scheduled task: %x", hr);
        pTsCom->Release();
        return hr;
    }

    SysFreeString(taskName);
    SysFreeString(startTime);
    pTsCom->Release();
    printf("\nthe scheduled task is created.\n");
    return hr;
}

HRESULT TestIfTaskScheduled()
{
    ITsComUser *pTsCom = NULL;
    HRESULT hr = CoCreateInstance( CLSID_TsComUser,
                           NULL,
                           CLSCTX_INPROC_SERVER,
                           IID_ITsComUser,
                           (void**)&pTsCom );  
    if (FAILED(hr))
    {
        printf("\nFailed to CoCreate an instance of the TsComUser class: %x", hr);
        return hr;
    }

    hr = pTsCom->ConnectTS();
    if (FAILED(hr))
    {
        printf("\nFailed to connect to the Task Scheduler: %x", hr);
        pTsCom->Release();
        return hr;
    }

    BSTR taskName = SysAllocString(L"PhishMe Scheduled Task Test");
    VARIANT_BOOL isCreated;
    hr = pTsCom->IsTaskScheduled(taskName, &isCreated);
    if (FAILED(hr))
    {
        printf("\nFailed to check if the task is scheduled or not: %x", hr);
        SysFreeString(taskName);
        pTsCom->Release();
        return hr;
    }
    
    SysFreeString(taskName);
    printf("\nthe task is %s.\n", (isCreated == VARIANT_TRUE) ? "scheduled" : "not scheduled");
    pTsCom->Release();
    return hr;
}

int main(int argc, char* argv[])
{
	HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	if (FAILED(hr))
	{
		printf("\nCoInitializeEx failed: %x", hr);
		return 1;
	}

    if (argc < 2)
    {
        printf("TsComApp usuage:\n"
            "\t'TsComApp -s (for scheduling a task with Task Scheduler)\n"
            "\t'TsComApp -t (for checking if a task is scheduled successfully with Task Scheduler)\n");
        return 0;
    }

    if (strncmp(argv[1], "-s", 2) == 0)
    {
        CreateScheduledTask();
    }
    else if (strncmp(argv[1], "-t", 2) == 0)
    {
        TestIfTaskScheduled();
    }
    else
    {
        printf("\nUnknown option. Use -s or -t.\n");
    }

    CoUninitialize();
	return 0;
}

