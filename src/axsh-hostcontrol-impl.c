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

// TODO remove definitions of CLSID_TclHostControl and IID_ITclHostControl (include generated .h by MIDL)
const CLSID CLSID_TclHostControl = {0x45397c60, 0xa814, 0x4c6d, {0xa1, 0x55,
    0x1f, 0x2d, 0x87, 0x2b, 0x1e, 0x83}};

const IID IID_ITclHostControl = {0xde08c005, 0xcadc, 0x4444, {0x90, 0xdd,
    0xe0, 0x44, 0x25, 0x1c, 0xb7, 0xe8}};

/* -------------------------------------------------------------------------
   -------------------------- ITclHostControl ------------------------------
   ------------------------------------------------------------------------- */
// TODO ITclHostControl implememtations
static STDMETHODIMP_(ULONG) ITclHostControl_AddRef(AXSH_TclHostControl *this)
{
    this->referenceCount++;
    return this->referenceCount;
}

static STDMETHODIMP_(ULONG) ITclHostControl_Release(AXSH_TclHostControl *this)
{
    this->referenceCount--;
    if (this->referenceCount == 0)
    {
        /* Release() all COM objects we're holding pointers to */
        if (this->pObjectTypeInfo != NULL)
            this->pObjectTypeInfo->lpVtbl->Release(this->pObjectTypeInfo);
        if (this->pInterfaceTypeInfo != NULL)
            this->pInterfaceTypeInfo->lpVtbl->Release(this->pInterfaceTypeInfo);

        /* release memory for the object */
        CoTaskMemFree(this);
        return 0;
    }

    return this->referenceCount;
}

static STDMETHODIMP ITclHostControl_QueryInterface(
            AXSH_TclHostControl *this,
            REFIID riid,
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
    ITclHostControl_AddRef(this);
    return S_OK;
}

static STDMETHODIMP ITclHostControl_GetTypeInfoCount(
            AXSH_TclHostControl *this,
            UINT *pctinfo)
{
    *pctinfo = 1;
    return S_OK;
}

static STDMETHODIMP ITclHostControl_GetTypeInfo(
            AXSH_TclHostControl *this,
            UINT ctinfo,
            LCID lcid,
            ITypeInfo **typeInfo)
{
    *typeInfo = this->pInterfaceTypeInfo;
    this->pInterfaceTypeInfo->lpVtbl->AddRef(this->pInterfaceTypeInfo);
    return S_OK;
}

static STDMETHODIMP ITclHostControl_GetIDsOfNames(
            AXSH_TclHostControl *this,
            REFIID riid,
            OLECHAR **rgszNames,
            UINT cNames,
            LCID lcid,
            DISPID *rgdispid)
{
    HRESULT hr;

    /* delegate to type info */
    hr = this->pInterfaceTypeInfo->lpVtbl->
        GetIDsOfNames(this->pInterfaceTypeInfo, rgszNames, cNames, rgdispid);
    return hr;
}

static STDMETHODIMP ITclHostControl_Invoke(
            AXSH_TclHostControl *this,
            DISPID id,
            REFIID riid,
            LCID lcid,
            WORD flag,
            DISPPARAMS *params,
            VARIANT *ret,
            EXCEPINFO *pei,
            UINT *pu)
{
    HRESULT hr;

    /* delegate to type info */
    hr = this->pInterfaceTypeInfo->lpVtbl->Invoke(this->pInterfaceTypeInfo,
        this, id, flag, params, ret, pei, pu);
    return hr;
}
static STDMETHODIMP ITclHostControl_GetStringVar(
            AXSH_TclHostControl *this,
            BSTR variable,
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

static STDMETHODIMP ITclHostControl_GetIntVar(
            AXSH_TclHostControl *this,
            BSTR variable,
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
static STDMETHODIMP IProvideMultipleClassInfo_QueryInterface(
            IProvideMultipleClassInfo *this,
            REFIID riid,
            void **ppv)
{
    /* delegate to base object */
    AXSH_TclHostControl *pBaseObj =
        (AXSH_TclHostControl *)(((unsigned char *)this -
        offsetof(AXSH_TclHostControl, multiClassInfo)));
    return ITclHostControl_QueryInterface(pBaseObj, riid, ppv);
}

static STDMETHODIMP_(ULONG) IProvideMultipleClassInfo_AddRef(
            IProvideMultipleClassInfo *this)
{
    /* delegate to base object */
    AXSH_TclHostControl *pBaseObj =
        (AXSH_TclHostControl *)(((unsigned char *)this -
        offsetof(AXSH_TclHostControl, multiClassInfo)));
    return ITclHostControl_AddRef(pBaseObj);
}

static STDMETHODIMP_(ULONG) IProvideMultipleClassInfo_Release(
            IProvideMultipleClassInfo *this)
{
    /* delegate to base object */
    AXSH_TclHostControl *pBaseObj =
        (AXSH_TclHostControl *)(((unsigned char *)this -
        offsetof(AXSH_TclHostControl, multiClassInfo)));
    return ITclHostControl_Release(pBaseObj);
}

static STDMETHODIMP IProvideMultipleClassInfo_GetClassInfo(
            IProvideMultipleClassInfo *this,
            ITypeInfo **classITypeInfo)
{
    AXSH_TclHostControl *pBaseObj =
        (AXSH_TclHostControl *)(((unsigned char *)this -
        offsetof(AXSH_TclHostControl, multiClassInfo)));
    *classITypeInfo = pBaseObj->pObjectTypeInfo;
    pBaseObj->pObjectTypeInfo->lpVtbl->AddRef(pBaseObj->pObjectTypeInfo);
    return S_OK;
}

static STDMETHODIMP IProvideMultipleClassInfo_GetGUID(
            IProvideMultipleClassInfo *this,
            DWORD guidType,
            GUID *guid)
{
    if (guidType == GUIDKIND_DEFAULT_SOURCE_DISP_IID)
        return E_NOTIMPL;

    return E_INVALIDARG;
}
static STDMETHODIMP IProvideMultipleClassInfo_GetMultiTypeInfoCount(
            IProvideMultipleClassInfo *this,
            ULONG *count)
{
    *count = 1;
    return S_OK;
}
static STDMETHODIMP IProvideMultipleClassInfo_GetInfoOfIndex(
            IProvideMultipleClassInfo *this,
            ULONG objNum,
            DWORD flags,
            ITypeInfo **classITypeInfo,
            DWORD *retFlags,
            ULONG *reservedIds,
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
    ITclHostControl_QueryInterface,
    ITclHostControl_AddRef,
    ITclHostControl_Release,
    ITclHostControl_GetTypeInfoCount,
    ITclHostControl_GetTypeInfo,
    ITclHostControl_GetIDsOfNames,
    ITclHostControl_Invoke,
    ITclHostControl_GetStringVar,
    ITclHostControl_GetIntVar
};
#pragma warning( pop )

/* vtable for IProvideMultipleClassInfo object */
static IProvideMultipleClassInfoVtbl g_MultiClassInfoVTable = {
    IProvideMultipleClassInfo_QueryInterface,
    IProvideMultipleClassInfo_AddRef,
    IProvideMultipleClassInfo_Release,
    IProvideMultipleClassInfo_GetClassInfo,
    IProvideMultipleClassInfo_GetGUID,
    IProvideMultipleClassInfo_GetMultiTypeInfoCount,
    IProvideMultipleClassInfo_GetInfoOfIndex
};

/* -------------------------------------------------------------------------
   ------------------------- initialization function -----------------------
   ------------------------------------------------------------------------- */
AXSH_TclHostControl * AXSH_CreateTclHostControl(AXSH_EngineState *pEngineState)
{
    HRESULT hr;
    AXSH_TclHostControl *pTemp;

    if (g_pTypeLibrary == NULL)
    {
        // TODO write error message "type library not loaded"
        return NULL;
    }
    pTemp = CoTaskMemAlloc(sizeof(*pTemp));
    if (pTemp == NULL)
    {
        // TODO write error message "out of memory"
        return NULL;
    }
    pTemp->hostCtl.lpVtbl = &g_TclHostControlVTable;
    pTemp->multiClassInfo.lpVtbl = &g_MultiClassInfoVTable;
    pTemp->referenceCount = 1;
    pTemp->pObjectTypeInfo = NULL;
    pTemp->pInterfaceTypeInfo = NULL;
    pTemp->pEngineState = pEngineState;

    /* get type info from type library */
    hr = g_pTypeLibrary->lpVtbl->GetTypeInfoOfGuid(g_pTypeLibrary,
        &CLSID_TclHostControl, &pTemp->pObjectTypeInfo);
    if (FAILED(hr))
    {
        // TODO write error message "could not get TclHostControl class type info from type library"
        goto errcleanup1;
    }
    hr = g_pTypeLibrary->lpVtbl->GetTypeInfoOfGuid(g_pTypeLibrary,
        &IID_ITclHostControl, &pTemp->pInterfaceTypeInfo);
    if (FAILED(hr))
    {
        // TODO write error message "could not get ITclHostControl interface type info from type library"
        goto errcleanup2;
    }

    /* no error */
    return pTemp;

errcleanup2:
    pTemp->pInterfaceTypeInfo->lpVtbl->Release(pTemp->pInterfaceTypeInfo);
errcleanup1:
    CoTaskMemFree(pTemp);
    return NULL;
}
