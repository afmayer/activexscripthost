/* ActiveX Script Host
 * (C) 2012 Alexander F. Mayer
 */
#include "axsh-include-all.h"

char * AXSH_InitEngineState(AXSH_EngineState *pEngineState, GUID *pEngineGuid,
                             Tcl_Interp *pTclInterp)
{
    HRESULT hr;
    char *pRetString;
    AXSH_TclHostControl *pTempHostCtl;

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
        goto cleanup1;
    }

    /* allocate space for ActiveScriptSite object and init vtables */
    pEngineState->pTclScriptSite = AXSH_CreateTclActiveScriptSite(pEngineState);
    if (pEngineState->pTclScriptSite == NULL)
    {
        pRetString = "out of memory";
        goto cleanup2;
    }

    /* initialize engine */
    hr = pEngineState->pActiveScriptParse->lpVtbl->
            InitNew(pEngineState->pActiveScriptParse);
    if (FAILED(hr))
    {
        pRetString = "InitNew() on ActiveScriptParse object failed";
        goto cleanup3;
    }

    hr = pEngineState->pActiveScript->lpVtbl->
            SetScriptSite(pEngineState->pActiveScript,
            &pEngineState->pTclScriptSite->site);
    if (FAILED(hr))
    {
        pRetString = "SetScriptSite() on ActiveScript object failed";
        goto cleanup3;
    }

    /* allocate space for TclHostControl object and initialize it */
    pTempHostCtl = malloc(sizeof(*pTempHostCtl));
    if (pTempHostCtl == NULL)
    {
        pRetString = "out of memory";
        goto cleanup3;
    }
    pRetString = AXSH_InitHostControl(pTempHostCtl, pEngineState);
    if (pRetString != NULL)
    {
        goto cleanup4;
    }

    /* store pointer to TclHostControl object in engine state and AddRef() it
       manually - otherwise the language engine could inadequately delete it */
    pEngineState->pTclHostControl = pTempHostCtl;
    pTempHostCtl->hostCtl.lpVtbl->
        AddRef((ITclHostControl *)pTempHostCtl);

    /* store a pointer to the Tcl interpreter in the engine state
       this is referenced by callback functions of the
       ActiveScriptSite interface */
    pEngineState->pTclInterp = pTclInterp;

    /* no error */
    return NULL;

cleanup4:
    free(pTempHostCtl);
cleanup3:
    pEngineState->pTclScriptSite->site.lpVtbl->
        Release(&pEngineState->pTclScriptSite->site);
cleanup2:
    pEngineState->pActiveScriptParse->lpVtbl->
        Release(pEngineState->pActiveScriptParse);
cleanup1:
    pEngineState->pActiveScript->lpVtbl->Release(pEngineState->pActiveScript);
    return pRetString;
}

char * AXSH_CleanupEngineState(AXSH_EngineState *pEngineState)
{
    HRESULT hr;

    hr = pEngineState->pActiveScript->lpVtbl-> // TODO check for hr == OLESCRIPT_S_PENDING or similar
        Close(pEngineState->pActiveScript);
    if (FAILED(hr))
        return "IActiveScript::Close failed";

    // TODO reactivate this better error message...
    //if (FAILED(hr))
    //{
    //    AXSH_SetTclResultToHRESULTErrString(interp, buffer, sizeof(buffer), hr,
    //        "IActiveScript::Close");
    //    return TCL_ERROR;
    //}

    pEngineState->pTclScriptSite->site.lpVtbl->
        Release(&pEngineState->pTclScriptSite->site);

    /* release ActiveScript and ActiveScriptParse objects */
    pEngineState->pActiveScript->lpVtbl->
        Release(pEngineState->pActiveScript);
    pEngineState->pActiveScriptParse->lpVtbl->
        Release(pEngineState->pActiveScriptParse);

    /* we don't need to Release() our TclHostControl object for some reason
       although it is manually AddRef()'d at engine state initialization */

    /* no error */
    return NULL;
}
