

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 7.00.0555 */
/* at Sun Jul 30 17:13:52 2017
 */
/* Compiler settings for TsCom.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 7.00.0555 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __TsCom_i_h__
#define __TsCom_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __ITsComUser_FWD_DEFINED__
#define __ITsComUser_FWD_DEFINED__
typedef interface ITsComUser ITsComUser;
#endif 	/* __ITsComUser_FWD_DEFINED__ */


#ifndef __TsComUser_FWD_DEFINED__
#define __TsComUser_FWD_DEFINED__

#ifdef __cplusplus
typedef class TsComUser TsComUser;
#else
typedef struct TsComUser TsComUser;
#endif /* __cplusplus */

#endif 	/* __TsComUser_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"
#include "shobjidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __ITsComUser_INTERFACE_DEFINED__
#define __ITsComUser_INTERFACE_DEFINED__

/* interface ITsComUser */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ITsComUser;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("EC7182B0-BB40-4D55-9F42-980A9491E97A")
    ITsComUser : public IDispatch
    {
    public:
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE ConnectTS( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE CreateTask( 
            /* [in] */ BSTR taskName,
            BSTR startTime) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE IsTaskScheduled( 
            /* [in] */ BSTR taskName,
            /* [out] */ VARIANT_BOOL *isScheduled) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ITsComUserVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ITsComUser * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ITsComUser * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ITsComUser * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ITsComUser * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ITsComUser * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ITsComUser * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ITsComUser * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *ConnectTS )( 
            ITsComUser * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *CreateTask )( 
            ITsComUser * This,
            /* [in] */ BSTR taskName,
            BSTR startTime);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *IsTaskScheduled )( 
            ITsComUser * This,
            /* [in] */ BSTR taskName,
            /* [out] */ VARIANT_BOOL *isScheduled);
        
        END_INTERFACE
    } ITsComUserVtbl;

    interface ITsComUser
    {
        CONST_VTBL struct ITsComUserVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ITsComUser_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ITsComUser_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ITsComUser_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ITsComUser_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ITsComUser_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ITsComUser_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ITsComUser_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ITsComUser_ConnectTS(This)	\
    ( (This)->lpVtbl -> ConnectTS(This) ) 

#define ITsComUser_CreateTask(This,taskName,startTime)	\
    ( (This)->lpVtbl -> CreateTask(This,taskName,startTime) ) 

#define ITsComUser_IsTaskScheduled(This,taskName,isScheduled)	\
    ( (This)->lpVtbl -> IsTaskScheduled(This,taskName,isScheduled) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ITsComUser_INTERFACE_DEFINED__ */



#ifndef __TsComLib_LIBRARY_DEFINED__
#define __TsComLib_LIBRARY_DEFINED__

/* library TsComLib */
/* [version][uuid] */ 

EXTERN_C const IID LIBID_TsComLib;

EXTERN_C const CLSID CLSID_TsComUser;

#ifdef __cplusplus

class DECLSPEC_UUID("46A42EBD-D7A6-459E-AE61-057B8A2F4930")
TsComUser;
#endif
#endif /* __TsComLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long *, BSTR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


