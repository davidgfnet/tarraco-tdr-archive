/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Sat Jun 09 23:03:46 2007
 */
/* Compiler settings for C:\Documents and Settings\David\Mis documentos\gameproject\cal3dvb\cal3dvb.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
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

#ifndef __cal3dvb_h__
#define __cal3dvb_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __ICal3DObject_FWD_DEFINED__
#define __ICal3DObject_FWD_DEFINED__
typedef interface ICal3DObject ICal3DObject;
#endif 	/* __ICal3DObject_FWD_DEFINED__ */


#ifndef __Cal3DObject_FWD_DEFINED__
#define __Cal3DObject_FWD_DEFINED__

#ifdef __cplusplus
typedef class Cal3DObject Cal3DObject;
#else
typedef struct Cal3DObject Cal3DObject;
#endif /* __cplusplus */

#endif 	/* __Cal3DObject_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

#ifndef __ICal3DObject_INTERFACE_DEFINED__
#define __ICal3DObject_INTERFACE_DEFINED__

/* interface ICal3DObject */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ICal3DObject;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("01E7F1E4-D966-46A2-9D48-78273FA6B86F")
    ICal3DObject : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE loadmesh( 
            BSTR __RPC_FAR *file,
            int __RPC_FAR *meshid) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE loadskeleton( 
            BSTR __RPC_FAR *file,
            int __RPC_FAR *result) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE loadanimation( 
            BSTR __RPC_FAR *file,
            int __RPC_FAR *animid) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE loadmaterial( 
            BSTR __RPC_FAR *file,
            int __RPC_FAR *matid) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE createmodel( 
            int __RPC_FAR *modelid) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE render( 
            int modelid,
            int __RPC_FAR *numvertices,
            int __RPC_FAR *numfaces,
            int __RPC_FAR *nummaterials,
            float __RPC_FAR *vertices,
            float __RPC_FAR *uvcoords,
            int __RPC_FAR *indices,
            int __RPC_FAR *textures,
            int __RPC_FAR *atributes,
            int using_vs) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE setlod( 
            int modelid,
            float lod) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE update( 
            int modelid,
            float eseconds) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE blendcycle( 
            int modelid,
            int animid,
            float weight,
            float delay) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE clearcycle( 
            int modelid,
            int animid,
            float delay) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE setanimationtime( 
            int modelid,
            float time) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ready( 
            int __RPC_FAR *texturelist,
            int __RPC_FAR *numtextures) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE getanimationduration( 
            int animid,
            float __RPC_FAR *animationduration) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE executeaction( 
            int modelid,
            int animid,
            float delayin,
            float delayout,
            float weight) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ICal3DObjectVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ICal3DObject __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ICal3DObject __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ICal3DObject __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ICal3DObject __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ICal3DObject __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ICal3DObject __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ICal3DObject __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *loadmesh )( 
            ICal3DObject __RPC_FAR * This,
            BSTR __RPC_FAR *file,
            int __RPC_FAR *meshid);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *loadskeleton )( 
            ICal3DObject __RPC_FAR * This,
            BSTR __RPC_FAR *file,
            int __RPC_FAR *result);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *loadanimation )( 
            ICal3DObject __RPC_FAR * This,
            BSTR __RPC_FAR *file,
            int __RPC_FAR *animid);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *loadmaterial )( 
            ICal3DObject __RPC_FAR * This,
            BSTR __RPC_FAR *file,
            int __RPC_FAR *matid);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *createmodel )( 
            ICal3DObject __RPC_FAR * This,
            int __RPC_FAR *modelid);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *render )( 
            ICal3DObject __RPC_FAR * This,
            int modelid,
            int __RPC_FAR *numvertices,
            int __RPC_FAR *numfaces,
            int __RPC_FAR *nummaterials,
            float __RPC_FAR *vertices,
            float __RPC_FAR *uvcoords,
            int __RPC_FAR *indices,
            int __RPC_FAR *textures,
            int __RPC_FAR *atributes,
            int using_vs);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *setlod )( 
            ICal3DObject __RPC_FAR * This,
            int modelid,
            float lod);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *update )( 
            ICal3DObject __RPC_FAR * This,
            int modelid,
            float eseconds);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *blendcycle )( 
            ICal3DObject __RPC_FAR * This,
            int modelid,
            int animid,
            float weight,
            float delay);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *clearcycle )( 
            ICal3DObject __RPC_FAR * This,
            int modelid,
            int animid,
            float delay);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *setanimationtime )( 
            ICal3DObject __RPC_FAR * This,
            int modelid,
            float time);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ready )( 
            ICal3DObject __RPC_FAR * This,
            int __RPC_FAR *texturelist,
            int __RPC_FAR *numtextures);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *getanimationduration )( 
            ICal3DObject __RPC_FAR * This,
            int animid,
            float __RPC_FAR *animationduration);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *executeaction )( 
            ICal3DObject __RPC_FAR * This,
            int modelid,
            int animid,
            float delayin,
            float delayout,
            float weight);
        
        END_INTERFACE
    } ICal3DObjectVtbl;

    interface ICal3DObject
    {
        CONST_VTBL struct ICal3DObjectVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICal3DObject_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ICal3DObject_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ICal3DObject_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ICal3DObject_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ICal3DObject_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ICal3DObject_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ICal3DObject_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define ICal3DObject_loadmesh(This,file,meshid)	\
    (This)->lpVtbl -> loadmesh(This,file,meshid)

#define ICal3DObject_loadskeleton(This,file,result)	\
    (This)->lpVtbl -> loadskeleton(This,file,result)

#define ICal3DObject_loadanimation(This,file,animid)	\
    (This)->lpVtbl -> loadanimation(This,file,animid)

#define ICal3DObject_loadmaterial(This,file,matid)	\
    (This)->lpVtbl -> loadmaterial(This,file,matid)

#define ICal3DObject_createmodel(This,modelid)	\
    (This)->lpVtbl -> createmodel(This,modelid)

#define ICal3DObject_render(This,modelid,numvertices,numfaces,nummaterials,vertices,uvcoords,indices,textures,atributes,using_vs)	\
    (This)->lpVtbl -> render(This,modelid,numvertices,numfaces,nummaterials,vertices,uvcoords,indices,textures,atributes,using_vs)

#define ICal3DObject_setlod(This,modelid,lod)	\
    (This)->lpVtbl -> setlod(This,modelid,lod)

#define ICal3DObject_update(This,modelid,eseconds)	\
    (This)->lpVtbl -> update(This,modelid,eseconds)

#define ICal3DObject_blendcycle(This,modelid,animid,weight,delay)	\
    (This)->lpVtbl -> blendcycle(This,modelid,animid,weight,delay)

#define ICal3DObject_clearcycle(This,modelid,animid,delay)	\
    (This)->lpVtbl -> clearcycle(This,modelid,animid,delay)

#define ICal3DObject_setanimationtime(This,modelid,time)	\
    (This)->lpVtbl -> setanimationtime(This,modelid,time)

#define ICal3DObject_ready(This,texturelist,numtextures)	\
    (This)->lpVtbl -> ready(This,texturelist,numtextures)

#define ICal3DObject_getanimationduration(This,animid,animationduration)	\
    (This)->lpVtbl -> getanimationduration(This,animid,animationduration)

#define ICal3DObject_executeaction(This,modelid,animid,delayin,delayout,weight)	\
    (This)->lpVtbl -> executeaction(This,modelid,animid,delayin,delayout,weight)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICal3DObject_loadmesh_Proxy( 
    ICal3DObject __RPC_FAR * This,
    BSTR __RPC_FAR *file,
    int __RPC_FAR *meshid);


void __RPC_STUB ICal3DObject_loadmesh_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICal3DObject_loadskeleton_Proxy( 
    ICal3DObject __RPC_FAR * This,
    BSTR __RPC_FAR *file,
    int __RPC_FAR *result);


void __RPC_STUB ICal3DObject_loadskeleton_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICal3DObject_loadanimation_Proxy( 
    ICal3DObject __RPC_FAR * This,
    BSTR __RPC_FAR *file,
    int __RPC_FAR *animid);


void __RPC_STUB ICal3DObject_loadanimation_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICal3DObject_loadmaterial_Proxy( 
    ICal3DObject __RPC_FAR * This,
    BSTR __RPC_FAR *file,
    int __RPC_FAR *matid);


void __RPC_STUB ICal3DObject_loadmaterial_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICal3DObject_createmodel_Proxy( 
    ICal3DObject __RPC_FAR * This,
    int __RPC_FAR *modelid);


void __RPC_STUB ICal3DObject_createmodel_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICal3DObject_render_Proxy( 
    ICal3DObject __RPC_FAR * This,
    int modelid,
    int __RPC_FAR *numvertices,
    int __RPC_FAR *numfaces,
    int __RPC_FAR *nummaterials,
    float __RPC_FAR *vertices,
    float __RPC_FAR *uvcoords,
    int __RPC_FAR *indices,
    int __RPC_FAR *textures,
    int __RPC_FAR *atributes,
    int using_vs);


void __RPC_STUB ICal3DObject_render_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICal3DObject_setlod_Proxy( 
    ICal3DObject __RPC_FAR * This,
    int modelid,
    float lod);


void __RPC_STUB ICal3DObject_setlod_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICal3DObject_update_Proxy( 
    ICal3DObject __RPC_FAR * This,
    int modelid,
    float eseconds);


void __RPC_STUB ICal3DObject_update_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICal3DObject_blendcycle_Proxy( 
    ICal3DObject __RPC_FAR * This,
    int modelid,
    int animid,
    float weight,
    float delay);


void __RPC_STUB ICal3DObject_blendcycle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICal3DObject_clearcycle_Proxy( 
    ICal3DObject __RPC_FAR * This,
    int modelid,
    int animid,
    float delay);


void __RPC_STUB ICal3DObject_clearcycle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICal3DObject_setanimationtime_Proxy( 
    ICal3DObject __RPC_FAR * This,
    int modelid,
    float time);


void __RPC_STUB ICal3DObject_setanimationtime_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICal3DObject_ready_Proxy( 
    ICal3DObject __RPC_FAR * This,
    int __RPC_FAR *texturelist,
    int __RPC_FAR *numtextures);


void __RPC_STUB ICal3DObject_ready_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICal3DObject_getanimationduration_Proxy( 
    ICal3DObject __RPC_FAR * This,
    int animid,
    float __RPC_FAR *animationduration);


void __RPC_STUB ICal3DObject_getanimationduration_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ICal3DObject_executeaction_Proxy( 
    ICal3DObject __RPC_FAR * This,
    int modelid,
    int animid,
    float delayin,
    float delayout,
    float weight);


void __RPC_STUB ICal3DObject_executeaction_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ICal3DObject_INTERFACE_DEFINED__ */



#ifndef __CAL3DVBLib_LIBRARY_DEFINED__
#define __CAL3DVBLib_LIBRARY_DEFINED__

/* library CAL3DVBLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_CAL3DVBLib;

EXTERN_C const CLSID CLSID_Cal3DObject;

#ifdef __cplusplus

class DECLSPEC_UUID("6F1BCB9E-077E-4F5C-8732-9B361F7763E9")
Cal3DObject;
#endif
#endif /* __CAL3DVBLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long __RPC_FAR *, unsigned long            , BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserMarshal(  unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserUnmarshal(unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long __RPC_FAR *, BSTR __RPC_FAR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
