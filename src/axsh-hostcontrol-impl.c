/* ActiveX Script Host
 * (C) 2012 Alexander F. Mayer
 */
#include "axsh-include-all.h"

/* This contains:
 *     GUIDs for object and its vtable
 *     static method implementations for interfaces
 *     static vtables for each interface
 *     helper functions for initialization
 */

const CLSID CLSID_ITclHostControl = {0x45397c60, 0xa814, 0x4c6d, {0xa1, 0x55,
    0x1f, 0x2d, 0x87, 0x2b, 0x1e, 0x83}};

const IID IID_ITclHostControl = {0xde08c005, 0xcadc, 0x4444, {0x90, 0xdd,
    0xe0, 0x44, 0x25, 0x1c, 0xb7, 0xe8}};

/* -------------------------------------------------------------------------
   -------------------------- ITclHostControl ------------------------------
   ------------------------------------------------------------------------- */
// TODO ITclHostControl implememtations
static STDMETHODIMP_(ULONG) AddRef(AXSH_TclHostControl *this)
{
    this->referenceCount++;
    return this->referenceCount;
}

static STDMETHODIMP_(ULONG) Release(AXSH_TclHostControl *this)
{
    this->referenceCount--;
    if (this->referenceCount == 0)
    {
        /* Release() all COM objects we're holding pointers to */
        if (this->pObjectTypeInfo != NULL)
            this->pObjectTypeInfo->lpVtbl->Release(this->pObjectTypeInfo);
        if (this->pVtableTypeInfo != NULL)
            this->pVtableTypeInfo->lpVtbl->Release(this->pVtableTypeInfo);

        /* free() the object itself */
        free(this);
        return 0;
    }

    return this->referenceCount;
}

static STDMETHODIMP QueryInterface(AXSH_TclHostControl *this, REFIID riid,
                                   void **ppv)
{
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IDispatch))
        *ppv = this;
    else if (IsEqualIID(riid, &IID_IProvideMultipleClassInfo) ||
        IsEqualIID(riid, &IID_IProvideClassInfo2) ||
        IsEqualIID(riid, &IID_IProvideClassInfo))
        ((unsigned char *)this +
        offsetof(AXSH_TclHostControl, multiClassInfo));
    else
    {
        *ppv = 0;
        return E_NOINTERFACE;
    }
    AddRef(this);
    return S_OK;
}

static STDMETHODIMP GetTypeInfoCount(AXSH_TclHostControl *this, UINT *pctinfo)
{
    *pctinfo = 1;
    return S_OK;
}

static STDMETHODIMP GetTypeInfo(AXSH_TclHostControl *this, UINT ctinfo,
                                LCID lcid, ITypeInfo **typeInfo)
{
    *typeInfo = this->pVtableTypeInfo;
    this->pVtableTypeInfo->lpVtbl->AddRef(this->pVtableTypeInfo);
    return S_OK;
}

static STDMETHODIMP GetIDsOfNames(AXSH_TclHostControl *this, REFIID riid,
                                  OLECHAR **rgszNames, UINT cNames, LCID lcid,
                                  DISPID *rgdispid)
{
    HRESULT hr;

    /* delegate to type info */
    hr = this->pVtableTypeInfo->lpVtbl->GetIDsOfNames(this->pVtableTypeInfo,
        rgszNames, cNames, rgdispid);
    return hr;
}

static STDMETHODIMP Invoke(AXSH_TclHostControl *this, DISPID id, REFIID riid,
                           LCID lcid, WORD flag, DISPPARAMS *params,
                           VARIANT *ret, EXCEPINFO *pei, UINT *pu)
{
    HRESULT hr;

    /* delegate to type info */
    hr = this->pVtableTypeInfo->lpVtbl->Invoke(this->pVtableTypeInfo,
        this, id, flag, params, ret, pei, pu);
    return hr;
}
static STDMETHODIMP GetStringVar(AXSH_TclHostControl *this, BSTR variable,
                                 BSTR *value)
{
    Tcl_Obj     *pVariableNameObj;
    Tcl_Obj     *pVariableObj;
    Tcl_UniChar *pStringValueUTF16;
    int         stringValueUTF16Length;

    pVariableNameObj = Tcl_NewUnicodeObj(variable, -1);
    Tcl_IncrRefCount(pVariableNameObj); /* so it can be decreased afterwards */
    pVariableObj = Tcl_ObjGetVar2(this->pEngineState->pTclInterp,
        pVariableNameObj, NULL, TCL_LEAVE_ERR_MSG);
    Tcl_DecrRefCount(pVariableNameObj); /* now it is free()'d */
    if (pVariableObj == NULL)
    {
        /* an error occurred while retrieving the variable's value */
        // TODO: somehow pass the error string (stored in interpreter's result) to the script engine...
        return E_FAIL;
    }
    pStringValueUTF16 = Tcl_GetUnicodeFromObj(pVariableObj,
        &stringValueUTF16Length);
    *value = SysAllocStringLen(pStringValueUTF16, stringValueUTF16Length);
    return S_OK;
}

static STDMETHODIMP GetIntVar(AXSH_TclHostControl *this, BSTR variable,
                              INT *value)
{
    Tcl_Obj *pVariableNameObj;
    Tcl_Obj *pVariableObj;

    pVariableNameObj = Tcl_NewUnicodeObj(variable, -1);
    Tcl_IncrRefCount(pVariableNameObj); /* so it can be decreased afterwards */
    pVariableObj = Tcl_ObjGetVar2(this->pEngineState->pTclInterp,
        pVariableNameObj, NULL, TCL_LEAVE_ERR_MSG);
    Tcl_DecrRefCount(pVariableNameObj); /* now it is free()'d */
    if (pVariableObj == NULL)
    {
        /* an error occurred while retrieving the variable's value */
        // TODO: somehow pass the error string (stored in interpreter's result) to the script engine...
        return E_FAIL;
    }
    Tcl_GetIntFromObj(this->pEngineState->pTclInterp, pVariableObj, value);
    return S_OK;
}

/* -------------------------------------------------------------------------
   ---------------------- IProvideMultipleClassInfo ------------------------
   ------------------------------------------------------------------------- */
// TODO IProvideMultipleClassInfo implementations
// TODO other names for 'secondary' interface implementations
static STDMETHODIMP QueryInterface_CInfo(IProvideMultipleClassInfo *this,
                                         REFIID riid, void **ppv)
{
    /* delegate to base object */
    AXSH_TclHostControl *pBaseObj =
        (AXSH_TclHostControl *)(((unsigned char *)this -
        offsetof(AXSH_TclHostControl, multiClassInfo)));
    return QueryInterface(pBaseObj, riid, ppv);
}

static STDMETHODIMP_(ULONG) AddRef_CInfo(IProvideMultipleClassInfo *this)
{
    /* delegate to base object */
    AXSH_TclHostControl *pBaseObj =
        (AXSH_TclHostControl *)(((unsigned char *)this -
        offsetof(AXSH_TclHostControl, multiClassInfo)));
    return AddRef(pBaseObj);
}

static STDMETHODIMP_(ULONG) Release_CInfo(IProvideMultipleClassInfo *this)
{
    /* delegate to base object */
    AXSH_TclHostControl *pBaseObj =
        (AXSH_TclHostControl *)(((unsigned char *)this -
        offsetof(AXSH_TclHostControl, multiClassInfo)));
    return Release(pBaseObj);
}

static STDMETHODIMP GetClassInfo_CInfo(IProvideMultipleClassInfo *this,
                                       ITypeInfo **classITypeInfo)
{
    AXSH_TclHostControl *pBaseObj =
        (AXSH_TclHostControl *)(((unsigned char *)this -
        offsetof(AXSH_TclHostControl, multiClassInfo)));
    *classITypeInfo = pBaseObj->pObjectTypeInfo;
    pBaseObj->pObjectTypeInfo->lpVtbl->AddRef(pBaseObj->pObjectTypeInfo);
    return S_OK;
}

static STDMETHODIMP GetGUID_CInfo(IProvideMultipleClassInfo *this,
                                  DWORD guidType, GUID *guid)
{
    if (guidType == GUIDKIND_DEFAULT_SOURCE_DISP_IID)
        return E_NOTIMPL;

    return E_INVALIDARG;
}
static STDMETHODIMP GetMultiTypeInfoCount_CInfo(
                                IProvideMultipleClassInfo *this, ULONG *count)
{
    *count = 1;
    return S_OK;
}
static STDMETHODIMP GetInfoOfIndex_CInfo(IProvideMultipleClassInfo *this,
                                         ULONG objNum, DWORD flags,
                                         ITypeInfo **classITypeInfo,
                                         DWORD *retFlags, ULONG *reservedIds,
                                         GUID *defVTableGuid,
                                         GUID *defSrcVTableGuid)
{
    AXSH_TclHostControl *pBaseObj =
        (AXSH_TclHostControl *)(((unsigned char *)this -
        offsetof(AXSH_TclHostControl, multiClassInfo)));

    *retFlags = 0; /* reset flags that are returned */

    /* objNum is ignored - only 1 implemented in this IDispatch interface */

    if (flags & MULTICLASSINFO_GETNUMRESERVEDDISPIDS)
    {
        /* caller asks for number of DISPIDs */
        *reservedIds = 1; // TODO number of DISPIDs in ITclHostControl likely to change, don't forget it here...
        *retFlags |= MULTICLASSINFO_GETNUMRESERVEDDISPIDS;
    }

    if (flags & MULTICLASSINFO_GETIIDPRIMARY)
    {
        /* caller asks for IID of the default interface (for this IDispatch) */
        memcpy(defVTableGuid, &IID_ITclHostControl,
            sizeof(IID_ITclHostControl));
        *retFlags |= MULTICLASSINFO_GETIIDPRIMARY;
    }

    if (flags & MULTICLASSINFO_GETIIDSOURCE)
    {
        /* NOT IMPLEMENTED */
    }

    if (flags & MULTICLASSINFO_GETTYPEINFO)
    {
        /* caller asks for object type info */
        *classITypeInfo = pBaseObj->pObjectTypeInfo;
        pBaseObj->pObjectTypeInfo->lpVtbl->AddRef(pBaseObj->pObjectTypeInfo);
        *retFlags |= MULTICLASSINFO_GETTYPEINFO;
    }

    return S_OK;
}

/* -------------------------------------------------------------------------
   ----------------------------- vtables -----------------------------------
   ------------------------------------------------------------------------- */

/* vtable for ITclHostControl object */
#pragma warning( push )
#pragma warning( disable : 4028 )
static ITclHostControlVtbl g_TclHostControlVTable = {
    QueryInterface,
    AddRef,
    Release,
    GetTypeInfoCount,
    GetTypeInfo,
    GetIDsOfNames,
    Invoke,
    GetStringVar,
    GetIntVar
};
#pragma warning( pop )

/* vtable for IProvideMultipleClassInfo object */
static IProvideMultipleClassInfoVtbl g_MultiClassInfoVTable = {
    QueryInterface_CInfo,
    AddRef_CInfo,
    Release_CInfo,
    GetClassInfo_CInfo,
    GetGUID_CInfo,
    GetMultiTypeInfoCount_CInfo,
    GetInfoOfIndex_CInfo
};

/* -------------------------------------------------------------------------
   ------------------------- initialization function -----------------------
   ------------------------------------------------------------------------- */
char * AXSH_InitHostControl(AXSH_TclHostControl *this,
                            AXSH_EngineState *pEngineState)
{
    HRESULT   hr;
    ITypeLib  *pTempTypeLib;
    ITypeInfo *pTempObjTypeInfo;
    ITypeInfo *pTempVtableTypeInfo;

    /* vtables */
    this->hostCtl.lpVtbl = &g_TclHostControlVTable;
    this->multiClassInfo.lpVtbl = &g_MultiClassInfoVTable;

    /* initialize other data */
    this->referenceCount = 0;
    this->pEngineState = pEngineState;
    this->pObjectTypeInfo = NULL;
    this->pVtableTypeInfo = NULL;

    hr = LoadTypeLib(_wpgmptr, &pTempTypeLib); /* get type info from DLL */
    if (FAILED(hr))
        return "Could not load type library for ITclHostControl";

    hr = pTempTypeLib->lpVtbl->GetTypeInfoOfGuid(pTempTypeLib,
        &CLSID_ITclHostControl, &pTempObjTypeInfo);
    if (FAILED(hr))
        return "Could not get ITclHostControl object type info "
            "from type library";

    hr = pTempTypeLib->lpVtbl->GetTypeInfoOfGuid(pTempTypeLib,
        &IID_ITclHostControl, &pTempVtableTypeInfo);
    if (FAILED(hr))
        return "Could not get ITclHostControl VTable type info "
            "from type library";

    /* release TypeLib after TypeInfo extraction */
    pTempTypeLib->lpVtbl->Release(pTempTypeLib);

    /* TypeInfo extraction successful */
    pTempObjTypeInfo->lpVtbl->AddRef(pTempObjTypeInfo);
    this->pObjectTypeInfo = pTempObjTypeInfo;
    pTempVtableTypeInfo->lpVtbl->AddRef(pTempVtableTypeInfo);
    this->pVtableTypeInfo = pTempVtableTypeInfo;

    /* no error */
    return NULL;
}
