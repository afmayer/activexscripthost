#include "axsh-include-all.h"

char * AXSH_OpenScriptEngine(AXSH_EngineState *pEngineState, GUID *pEngineGuid,
                             Tcl_Interp *pTclInterp)
{
    // TODO split into static subfunctions - each responsible for 1 element in the structure
    //      so memory cleanup is easier
    HRESULT hr;
    AXSH_TclActiveScriptSite *pTemp;

    /* create script engine instance */
    hr = CoCreateInstance(pEngineGuid, 0, CLSCTX_ALL, &IID_IActiveScript,
        (void **)&(pEngineState->pActiveScript)); // TODO store in temp var before assigning to struct...
    if (hr != S_OK)
        return "CoCreateInstance() on engine failed";

    /* get IActiveScriptParse interface */
    hr = pEngineState->pActiveScript->lpVtbl->QueryInterface(pEngineState->
            pActiveScript, &IID_IActiveScriptParse,
            (void **)&(pEngineState->pActiveScriptParse));
    if (hr != S_OK)
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
    if (hr != S_OK)
    {
        free(pEngineState->pTclScriptSite);
        return "InitNew() on ActiveScriptParse object failed";
    }

    hr = pEngineState->pActiveScript->lpVtbl->
            SetScriptSite(pEngineState->pActiveScript,
            &pEngineState->pTclScriptSite->site);
    if (hr != S_OK)
    {
        free(pEngineState->pTclScriptSite);
        return "SetScriptSite() on ActiveScript object failed";
    }

    // TODO initialize AXSH_TclHostControl object

    /* store a pointer to the Tcl interpreter in the engine state
       this is referenced by callback functions of the
       ActiveScriptSite interface */
    pEngineState->pTclInterp = pTclInterp;

    /* no error */
    return NULL;
}

char * AXSH_CloseScriptEngine(AXSH_EngineState *pEngineState)
{
    HRESULT hr;

    hr = pEngineState->pActiveScript->lpVtbl-> // TODO check for hr == OLESCRIPT_S_PENDING or similar
        Close(pEngineState->pActiveScript);
    if (hr != S_OK)
        return "IActiveScript::Close failed";

    // TODO REACTIVATE THIS BETTER ERROR MESSAGE...
    //if (hr != S_OK)
    //{
    //    AXSH_SetTclResultToHRESULTErrString(interp, buffer, sizeof(buffer), hr,
    //        "IActiveScript::Close");
    //    return TCL_ERROR;
    //}

    /* we don't need to Release() our AXSH_TclActiveScriptSite object because
       we never AddRef()'d it - the engine will do this */

    /* release ActiveScript and ActiveScriptParse objects */
    pEngineState->pActiveScript->lpVtbl->
        Release(pEngineState->pActiveScript);
    pEngineState->pActiveScriptParse->lpVtbl->
        Release(pEngineState->pActiveScriptParse);

    /* no error */
    return NULL;
}
