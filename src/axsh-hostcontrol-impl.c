#include "axsh-include-all.h"

/* This contains:
 *     static method implementations for interfaces
 *     static vtables for each interface
 *     helper functions for initialization
 */

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
void AXSH_InitHostControl(AXSH_TclHostControl *this)
{
    /* vtables */
    this->hostCtl.lpVtbl = &g_TclHostControlVTable;
    this->multiClassInfo.lpVtbl = &g_MultiClassInfoVTable;

    /* data */
    this->referenceCount = 0;
}
