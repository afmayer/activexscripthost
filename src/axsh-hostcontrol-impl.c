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
{return S_OK;}
static STDMETHODIMP GetTypeInfo(AXSH_TclHostControl *this, UINT ctinfo,
                                LCID lcid, ITypeInfo **typeInfo)
{return S_OK;}
static STDMETHODIMP GetIDsOfNames(AXSH_TclHostControl *this, REFIID riid,
                                  OLECHAR **rgszNames, UINT cNames, LCID lcid,
                                  DISPID *rgdispid)
{return S_OK;}
static STDMETHODIMP Invoke(AXSH_TclHostControl *this, DISPID id, REFIID riid,
                           LCID lcid, WORD flag, DISPPARAMS *params,
                           VARIANT *ret, EXCEPINFO *pei, UINT *pu)
{return S_OK;}
static STDMETHODIMP GetStringVar(AXSH_TclHostControl *this, BSTR variable,
                                 BSTR *value)
{return S_OK;}

/* -------------------------------------------------------------------------
   ---------------------- IProvideMultipleClassInfo ------------------------
   ------------------------------------------------------------------------- */
// TODO IProvideMultipleClassInfo implementations
// TODO other names for 'secondary' interface implementations
static STDMETHODIMP QueryInterface_CInfo(IProvideMultipleClassInfo *this,
                                         REFIID riid, void **ppv)
{
    /* delegate to base object */
    this = (IProvideMultipleClassInfo *)(((unsigned char *)this -
        offsetof(AXSH_TclHostControl, multiClassInfo)));
    return QueryInterface((AXSH_TclHostControl *)this, riid, ppv);
}

static STDMETHODIMP_(ULONG) AddRef_CInfo(IProvideMultipleClassInfo *this)
{
    /* delegate to base object */
    this = (IProvideMultipleClassInfo *)(((unsigned char *)this -
        offsetof(AXSH_TclHostControl, multiClassInfo)));
    return AddRef((AXSH_TclHostControl *)this);
}

static STDMETHODIMP_(ULONG) Release_CInfo(IProvideMultipleClassInfo *this)
{
    /* delegate to base object */
    this = (IProvideMultipleClassInfo *)(((unsigned char *)this -
        offsetof(AXSH_TclHostControl, multiClassInfo)));
    return Release((AXSH_TclHostControl *)this);
}

static STDMETHODIMP GetClassInfo_CInfo(IProvideMultipleClassInfo *this,
                                       ITypeInfo **classITypeInfo)
{return S_OK;}
static STDMETHODIMP GetGUID_CInfo(IProvideMultipleClassInfo *this,
                                  DWORD guidType, GUID *guid)
{return S_OK;}
static STDMETHODIMP GetMultiTypeInfoCount_CInfo(
                                IProvideMultipleClassInfo *this, ULONG *count)
{return S_OK;}
static STDMETHODIMP GetInfoOfIndex_CInfo(IProvideMultipleClassInfo *this,
                                         ULONG objNum, DWORD flags,
                                         ITypeInfo **classITypeInfo,
                                         DWORD *retFlags, ULONG *reservedIds,
                                         GUID *defVTableGuid,
                                         GUID *defSrcVTableGuid)
{return S_OK;}

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
    GetStringVar
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
char * AXSH_InitHostControl(AXSH_TclHostControl *this)
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
    this->pObjectTypeInfo = NULL;
    this->pVtableTypeInfo = NULL;

    hr = LoadTypeLib(_wpgmptr, &pTempTypeLib); /* get type info from DLL */
    if (hr != S_OK)
        return "Could not load type library for ITclHostControl";

    hr = pTempTypeLib->lpVtbl->GetTypeInfoOfGuid(pTempTypeLib,
        &CLSID_ITclHostControl, &pTempObjTypeInfo);
    if (hr != S_OK)
        return "Could not get ITclHostControl object type info "
            "from type library";

    hr = pTempTypeLib->lpVtbl->GetTypeInfoOfGuid(pTempTypeLib,
        &IID_ITclHostControl, &pTempVtableTypeInfo);
    if (hr != S_OK)
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
