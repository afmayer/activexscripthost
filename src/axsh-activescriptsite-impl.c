/* ActiveX Script Host
 * (C) 2012 Alexander F. Mayer
 */

#include "axsh-include-all.h"

/* This contains:
 *     static method implementations for interfaces
 *     static vtables for each interface
 *     helper functions for initialization
 */

/* -------------------------------------------------------------------------
   ------------------------- IActiveScriptSite -----------------------------
   ------------------------------------------------------------------------- */
static STDMETHODIMP_(ULONG) IActiveScriptSite_AddRef(
            AXSH_TclActiveScriptSite *this)
{
    this->referenceCount++;
    return this->referenceCount;
}

static STDMETHODIMP_(ULONG) IActiveScriptSite_Release(
            AXSH_TclActiveScriptSite *this)
{
    this->referenceCount--;
    if (this->referenceCount == 0)
    {
        CoTaskMemFree(this);
        return 0;
    }
    return this->referenceCount;
}

static STDMETHODIMP IActiveScriptSite_QueryInterface(
            AXSH_TclActiveScriptSite *this,
            REFIID riid,
            void **ppv)
{
    // An ActiveX Script Host is supposed to provide an object
    // with multiple interfaces, where the base object is an
    // IActiveScriptSite, and there is an IActiveScriptSiteWindow
    // sub-object. Therefore, a caller can pass an IUnknown or
    // IActiveScript VTable GUID if he wants our IActiveScriptSite.
    // Or, the caller can pass a IActiveScriptSiteWindow VTable
    // GUID if he wants our IActiveScriptSiteWindow sub-object
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IActiveScriptSite))
        *ppv = this;
    else if (IsEqualIID(riid, &IID_IActiveScriptSiteWindow))
        *ppv = ((unsigned char *)this + offsetof(AXSH_TclActiveScriptSite, siteWnd)); 
    else
    {
        *ppv = 0;
        return E_NOINTERFACE;
    }
    IActiveScriptSite_AddRef(this);
    return S_OK;
}

static STDMETHODIMP IActiveScriptSite_GetItemInfo(
            AXSH_TclActiveScriptSite *this,
            LPCOLESTR objectName,
            DWORD dwReturnMask,
            IUnknown **objPtr,
            ITypeInfo **typeInfo)
{
    // TODO return E_POINTER when one of the [out] pointers is NULL (COM rule violation)
    //      but allow workaround for buggy engines (like PerlScript) as long as
    //      the [out] pointer is not actually written to
    if (objPtr != NULL) *objPtr = 0;
    if (typeInfo != NULL) *typeInfo = 0;

    if (!wcscmp(objectName, L"tcl"))
    {
        if (dwReturnMask & SCRIPTINFO_IUNKNOWN)
        {
            *objPtr =
                (IUnknown *)(&this->pEngineState->pTclHostControl->hostCtl);
            this->pEngineState->pTclHostControl->hostCtl.lpVtbl->
                AddRef(&this->pEngineState->pTclHostControl->hostCtl);
        }

        if (dwReturnMask & SCRIPTINFO_ITYPEINFO)
        {
            *typeInfo = this->pEngineState->pTclHostControl->pObjectTypeInfo;
            this->pEngineState->pTclHostControl->pObjectTypeInfo->lpVtbl->
                AddRef(this->pEngineState->pTclHostControl->pObjectTypeInfo);
        }

        return S_OK;
    }

    /* unknown object name */
    return TYPE_E_ELEMENTNOTFOUND;
}

// Called by the script engine when there is an error running/parsing a script.
// The script engine passes us its IActiveScriptError object whose functions
// we can call to get information about the error, such as the line/character
// position (in the script) where the error occurred, an error message we can
// display to the user, etc. Typically, our OnScriptError will get the error
// message and display it to the user.
static STDMETHODIMP IActiveScriptSite_OnScriptError(
            AXSH_TclActiveScriptSite *this,
            IActiveScriptError *scriptError)
{
    ULONG       lineNumber;
    BSTR        desc;
    EXCEPINFO   ei;

    Tcl_Interp *pTclInterp = this->pEngineState->pTclInterp;
    Tcl_Obj **ppErrorResult = &this->pEngineState->pErrorResult;

    scriptError->lpVtbl->GetSourcePosition(scriptError, 0, &lineNumber, 0);

    // Note: The IActiveScriptError is supposed to clear our BSTR pointer
    // if it can't return the line. So we shouldn't have to do
    // that first. But you may need to uncomment the following line
    // if a script engine isn't properly written. Unfortunately, too many
    // engines are not written properly, so we uncomment it
    desc = 0;
    scriptError->lpVtbl->GetSourceLineText(scriptError, &desc);

    // Note: The IActiveScriptError is supposed to zero out any fields of
    // our EXCEPINFO that it doesn't fill in. So we shouldn't have to do
    // that first. But you may need to uncomment the following line
    // if a script engine isn't properly written
    memset(&ei, 0, sizeof(ei));
    scriptError->lpVtbl->GetExceptionInfo(scriptError, &ei);

    /* assemble an error message and store it in the engine state
       no reference counting used for this Tcl object because the object
       is used nowhere else within the interpreter - only here and when
       the error is set as result */
    *ppErrorResult = Tcl_NewObj();
    if (ei.bstrSource != NULL)
    {
        Tcl_AppendUnicodeToObj(*ppErrorResult, ei.bstrSource,
            SysStringLen(ei.bstrSource));
        Tcl_AppendToObj(*ppErrorResult, ":", -1);
    }
    if (ei.bstrDescription != NULL)
    {
        Tcl_AppendToObj(*ppErrorResult, " ", -1);
        Tcl_AppendUnicodeToObj(*ppErrorResult, ei.bstrDescription,
            SysStringLen(ei.bstrDescription));
    }
    {
        char lineAsText[32];
        _snprintf(lineAsText, sizeof(lineAsText), " at line %u",
            lineNumber + 1);
        lineAsText[sizeof(lineAsText) - 1] = 0;
        Tcl_AppendToObj(*ppErrorResult, lineAsText, -1);
    }
    if (desc != NULL)
    {
        Tcl_AppendToObj(*ppErrorResult, " (\"", -1);
        Tcl_AppendUnicodeToObj(*ppErrorResult, desc, -1);
        Tcl_AppendToObj(*ppErrorResult, "\")", -1);
    }

    // Free what we got from the IActiveScriptError functions
    SysFreeString(desc);
    SysFreeString(ei.bstrSource);
    SysFreeString(ei.bstrDescription);
    SysFreeString(ei.bstrHelpFile);
 
    return S_OK;
}

// Called when the script engine wants to know what language ID our
// program is using.
static STDMETHODIMP IActiveScriptSite_GetLCID(
            AXSH_TclActiveScriptSite *this,
            LCID *lcid)
{
    *lcid = LOCALE_USER_DEFAULT;
    return S_OK;
}

// Called when the script engine wishes to retrieve the version number
// (as a string) for the current document. We are expected to return a
// SysAllocString()'ed BSTR copy of this version string.
//
// Here's what I believe this is for. An engine may implement some
// IPersist object. When we AddScriptlet or ParseScriptText some
// script with the SCRIPTTEXT_ISPERSISTENT flag, then the script
// engine will save that script to disk for us, when we call the
// engine IPersist's Save() function. But maybe the engine wants to
// minimize unnecessary saving to disk. So, it fetches this version
// string and caches it when it does the first save to disk. Then,
// when we subsequently call IPersist Save again, the engine fetches
// this version string again, and compares it to the previous version
// string. If the same string, then the engine assumes the script doesn't
// need to be resaved. If different, then the engine assumes our
// document has changed and therefore it should resave the scripts.
// In other words, this is used simply as an "IsDirty" flag by IPersist
// Save to see if a persistant script needs to be resaved. I don't
// believe that engines without any IPersist object bother with this,
// and I have seen example engines _with_ an IPersist that don't
// bother with it.
static STDMETHODIMP IActiveScriptSite_GetDocVersionString(
            AXSH_TclActiveScriptSite *this,
            BSTR *version)
{
    // We have no document versions
    *version = 0;

    // If an engine chokes on the above, try this instead:
    //if (!(*version = SysAllocString(&EmptyStr[0])))
    //  return E_OUTOFMEMORY;

    return S_OK;
}

// Called when the engine has completely finished running scripts and
// has returned to INITIALIZED state. In many engines, this is not
// called because an engine alone can't determine when we are going
// to stop running/adding scripts to it.
static STDMETHODIMP IActiveScriptSite_OnScriptTerminate(
            AXSH_TclActiveScriptSite *this,
            const VARIANT *pvr,
            const EXCEPINFO *pei)
{
    return S_OK;
}

// Called when the script engine's state is changed (for example, by
// us calling the engine IActiveScript->SetScriptState). We're passed
// the new state.
static STDMETHODIMP IActiveScriptSite_OnStateChange(
            AXSH_TclActiveScriptSite *this,
            SCRIPTSTATE state)
{
    return S_OK;
}

// Called right before the script engine executes/interprets each
// script added via the engine IActiveScriptParse->ParseScriptText()
// or AddScriptlet(). This is also called when our IApp object
// calls some function in the script.
static STDMETHODIMP IActiveScriptSite_OnEnterScript(
            AXSH_TclActiveScriptSite *this)
{
    return S_OK;
}

// Called immediately after the script engine executes/interprets
// each script.
static STDMETHODIMP IActiveScriptSite_OnLeaveScript(
            AXSH_TclActiveScriptSite *this)
{
    return S_OK;
}

/* -------------------------------------------------------------------------
   ---------------------- IActiveScriptSiteWindow --------------------------
   ------------------------------------------------------------------------- */
static STDMETHODIMP IActiveScriptSiteWindow_QueryInterface(
            IActiveScriptSiteWindow *this,
            REFIID riid, void **ppv)
{
    /* delegate to base object */
    AXSH_TclActiveScriptSite *pBaseObj =
        (AXSH_TclActiveScriptSite *)(((unsigned char *)this -
        offsetof(AXSH_TclActiveScriptSite, siteWnd)));
    return IActiveScriptSite_QueryInterface(pBaseObj, riid, ppv);
}

static STDMETHODIMP_(ULONG) IActiveScriptSiteWindow_AddRef(
            IActiveScriptSiteWindow *this)
{
    /* delegate to base object */
    AXSH_TclActiveScriptSite *pBaseObj =
        (AXSH_TclActiveScriptSite *)(((unsigned char *)this -
        offsetof(AXSH_TclActiveScriptSite, siteWnd)));
    return IActiveScriptSite_AddRef(pBaseObj);
}

static STDMETHODIMP_(ULONG) IActiveScriptSiteWindow_Release(
            IActiveScriptSiteWindow *this)
{
    /* delegate to base object */
    AXSH_TclActiveScriptSite *pBaseObj =
        (AXSH_TclActiveScriptSite *)(((unsigned char *)this -
        offsetof(AXSH_TclActiveScriptSite, siteWnd)));
    return IActiveScriptSite_Release(pBaseObj);
}

// Called by the script engine when it wants to know what window it should
// use as the owner of any dialog box the engine presents.
static STDMETHODIMP IActiveScriptSiteWindow_GetSiteWindow(
            IActiveScriptSiteWindow *this,
            HWND *phwnd)
{
    // We have no app window
    *phwnd = 0;
    return S_OK;
}

// Called when the script engine wants us to enable/disable all of our open
// windows.
static STDMETHODIMP IActiveScriptSiteWindow_EnableModeless(
            IActiveScriptSiteWindow *this,
            BOOL enable)
{
    // We have no open windows
    return S_OK;
}

/* -------------------------------------------------------------------------
   ----------------------------- vtables -----------------------------------
   ------------------------------------------------------------------------- */

/* vtable for ActiveScriptSite object */
#pragma warning( push )
#pragma warning( disable : 4028 )
static IActiveScriptSiteVtbl g_SiteVTable = {
    IActiveScriptSite_QueryInterface,
    IActiveScriptSite_AddRef,
    IActiveScriptSite_Release,
    IActiveScriptSite_GetLCID,
    IActiveScriptSite_GetItemInfo,
    IActiveScriptSite_GetDocVersionString,
    IActiveScriptSite_OnScriptTerminate,
    IActiveScriptSite_OnStateChange,
    IActiveScriptSite_OnScriptError,
    IActiveScriptSite_OnEnterScript,
    IActiveScriptSite_OnLeaveScript
};
#pragma warning( pop )

/* vtable for ActiveScriptSiteWindow object */
static IActiveScriptSiteWindowVtbl g_SiteWindowVTable = {
    IActiveScriptSiteWindow_QueryInterface,
    IActiveScriptSiteWindow_AddRef,
    IActiveScriptSiteWindow_Release,
    IActiveScriptSiteWindow_GetSiteWindow,
    IActiveScriptSiteWindow_EnableModeless
};

/* -------------------------------------------------------------------------
   ------------------------- initialization function -----------------------
   ------------------------------------------------------------------------- */
AXSH_TclActiveScriptSite * AXSH_CreateTclActiveScriptSite(
            AXSH_EngineState *pEngineState)
{
    AXSH_TclActiveScriptSite *pTemp = CoTaskMemAlloc(sizeof(*pTemp));
    if (pTemp == NULL)
        return NULL;
    pTemp->site.lpVtbl = &g_SiteVTable;
    pTemp->siteWnd.lpVtbl = &g_SiteWindowVTable;
    pTemp->referenceCount = 1;
    pTemp->pEngineState = pEngineState;
    return pTemp;
}
