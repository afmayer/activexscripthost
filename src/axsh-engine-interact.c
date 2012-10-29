/* ActiveX Script Host
 * (C) 2012 Alexander F. Mayer
 */
#include "axsh-include-all.h"

char * AXSH_InitEngineState(AXSH_EngineState *pEngineState, GUID *pEngineGuid,
                             Tcl_Interp *pTclInterp)
{
    // TODO split into static subfunctions - each responsible for 1 element in the structure
    //      so memory cleanup is easier
    HRESULT hr;
    char *pRetString;
    AXSH_TclActiveScriptSite *pTemp;
    AXSH_TclHostControl *pTempHostCtl;

    /* create script engine instance */
    hr = CoCreateInstance(pEngineGuid, 0, CLSCTX_ALL, &IID_IActiveScript,
        (void **)&(pEngineState->pActiveScript)); // TODO store in temp var before assigning to struct...
    if (FAILED(hr))
        return "CoCreateInstance() on engine failed";

    /* get IActiveScriptParse interface */
    hr = pEngineState->pActiveScript->lpVtbl->QueryInterface(pEngineState->
            pActiveScript, &IID_IActiveScriptParse,
            (void **)&(pEngineState->pActiveScriptParse));
    if (FAILED(hr))
        return "QueryInterface() for getting ActiveScriptParse object failed";

    /* allocate space for ActiveScriptSite object and init vtables */
    pTemp = malloc(sizeof(*pTemp));
    if (pTemp == NULL)
        return "out of memory";
    AXSH_InitActiveScriptSite(pTemp, pEngineState);
    pEngineState->pTclScriptSite = pTemp;

    /* initialize engine */
    hr = pEngineState->pActiveScriptParse->lpVtbl->
            InitNew(pEngineState->pActiveScriptParse);
    if (FAILED(hr))
    {
        free(pEngineState->pTclScriptSite);
        return "InitNew() on ActiveScriptParse object failed";
    }

    hr = pEngineState->pActiveScript->lpVtbl->
            SetScriptSite(pEngineState->pActiveScript,
            &pEngineState->pTclScriptSite->site);
    if (FAILED(hr))
    {
        free(pEngineState->pTclScriptSite);
        return "SetScriptSite() on ActiveScript object failed";
    }

    /* allocate space for TclHostControl object and initialize it */
    pTempHostCtl = malloc(sizeof(*pTempHostCtl));
    if (pTempHostCtl == NULL)
    {
        free(pEngineState->pTclScriptSite);
        return "out of memory";
    }
    pRetString = AXSH_InitHostControl(pTempHostCtl, pEngineState);
    if (pRetString != NULL)
    {
        free(pTempHostCtl);
        free(pEngineState->pTclScriptSite);
        return pRetString;
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

    /* we don't need to Release() our AXSH_TclActiveScriptSite or
       AXSH_TclHostControl objects because we never AddRef()'d it - the engine
       will do this */

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
