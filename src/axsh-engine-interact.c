/* ActiveX Script Host
 * (C) 2012 Alexander F. Mayer
 */
#include "axsh-include-all.h"

char * AXSH_InitEngineState(
            AXSH_EngineState *pEngineState,
            GUID *pEngineGuid,
            Tcl_Interp *pTclInterp)
{
    HRESULT hr;
    char *pRetString;

    /* create script engine instance */
    hr = CoCreateInstance(pEngineGuid, 0, CLSCTX_ALL, &IID_IActiveScript,
        (void **)&(pEngineState->pActiveScript));
    if (FAILED(hr))
        return "CoCreateInstance() on engine failed";

    /* get IActiveScriptParse interface */
    hr = pEngineState->pActiveScript->lpVtbl->QueryInterface(pEngineState->
            pActiveScript, &IID_IActiveScriptParse,
            (void **)&(pEngineState->pActiveScriptParse));
    if (FAILED(hr))
    {
        pRetString =
            "QueryInterface() for getting ActiveScriptParse object failed";
        goto errcleanup1;
    }

    /* allocate space for ActiveScriptSite object and init vtables */
    pEngineState->pTclScriptSite = AXSH_CreateTclActiveScriptSite(pEngineState);
    if (pEngineState->pTclScriptSite == NULL)
    {
        pRetString = "out of memory";
        goto errcleanup2;
    }

    /* initialize engine */
    hr = pEngineState->pActiveScriptParse->lpVtbl->
            InitNew(pEngineState->pActiveScriptParse);
    if (FAILED(hr))
    {
        pRetString = "InitNew() on ActiveScriptParse object failed";
        goto errcleanup3;
    }

    hr = pEngineState->pActiveScript->lpVtbl->
            SetScriptSite(pEngineState->pActiveScript,
            &pEngineState->pTclScriptSite->site);
    if (FAILED(hr))
    {
        pRetString = "SetScriptSite() on ActiveScript object failed";
        goto errcleanup3;
    }

    /* allocate space for TclHostControl object and initialize it */
    pEngineState->pTclHostControl = AXSH_CreateTclHostControl(pEngineState);
    if (pEngineState->pTclHostControl == NULL)
    {
        pRetString = "error creating TclHostControl object";
        goto errcleanup3;
    }

    /* store a pointer to the Tcl interpreter in the engine state
       this is referenced by callback functions of the
       ActiveScriptSite interface */
    pEngineState->pTclInterp = pTclInterp;

    /* no error */
    return NULL;

errcleanup3:
    pEngineState->pTclScriptSite->site.lpVtbl->
        Release(&pEngineState->pTclScriptSite->site);
errcleanup2:
    pEngineState->pActiveScriptParse->lpVtbl->
        Release(pEngineState->pActiveScriptParse);
errcleanup1:
    pEngineState->pActiveScript->lpVtbl->Release(pEngineState->pActiveScript);
    return pRetString;
}

void AXSH_CleanupEngineState(AXSH_EngineState *pEngineState)
{
    /* we ignore the return code of IActiveScript::Close()
        --> E_UNEXPECTED means we can't recover anyways (already closed, etc.)
        --> OLESCRIPT_S_PENDING is not defined in any header files (unused) */
    pEngineState->pActiveScript->lpVtbl->Close(pEngineState->pActiveScript);

    pEngineState->pTclHostControl->hostCtl.lpVtbl->
        Release(&pEngineState->pTclHostControl->hostCtl);
    pEngineState->pTclScriptSite->site.lpVtbl->
        Release(&pEngineState->pTclScriptSite->site);

    /* release ActiveScript and ActiveScriptParse objects */
    pEngineState->pActiveScript->lpVtbl->
        Release(pEngineState->pActiveScript);
    pEngineState->pActiveScriptParse->lpVtbl->
        Release(pEngineState->pActiveScriptParse);
}
