#include "axsh-include-all.h"

static void AXSH_SetTclResultToHRESULTErrString(Tcl_Interp *interp,
                                char *pBuffer, size_t bufferSize, HRESULT hr,
                                char *pFunctionName)
{
    _snprintf(pBuffer, bufferSize, "%s returned %s", pFunctionName,
        AXSH_HRESULT2String(hr));
    pBuffer[bufferSize-1] = 0;
    Tcl_SetResult(interp, pBuffer, TCL_VOLATILE);
}

int AXSH_Tcl_CloseScriptEngine(ClientData clientData, Tcl_Interp *interp,
                               int objc, Tcl_Obj *CONST objv[])
{
    AXSH_EngineState *pEngineState = (AXSH_EngineState *)clientData;
    HRESULT hr;
    char *pRetString;
    char buffer[128];

    pRetString = AXSH_CloseScriptEngine(pEngineState);
    if (pRetString != NULL)
    {
        _snprintf(buffer, sizeof(buffer), "AXSH_CloseScriptEngine() says '%s'",
            pRetString);
        buffer[sizeof(buffer)-1] = 0;
        Tcl_SetResult(interp, buffer, TCL_VOLATILE);
    }

    /* delete Tcl command and free space @ clientData (engine state) */
    Tcl_DeleteCommandFromToken(interp, pEngineState->pTclCommandToken);
    free(pEngineState);

    return TCL_OK;
}

int AXSH_Tcl_ParseText(ClientData clientData, Tcl_Interp *interp, int objc,
                       Tcl_Obj *CONST objv[])
{
    AXSH_EngineState *pEngineState = (AXSH_EngineState *)clientData;
    char buffer[128];
    Tcl_UniChar *pScriptUTF16;
    int scriptUTF16Length;
    HRESULT hr;

    if (objc != 3)
    {
        Tcl_WrongNumArgs(interp, 2, objv, "script");
        return TCL_ERROR;
    }

    /* convert string to unicode format */
    pScriptUTF16 = Tcl_GetUnicodeFromObj(objv[2], &scriptUTF16Length);
    hr = pEngineState->pActiveScriptParse->lpVtbl->
            ParseScriptText(pEngineState->pActiveScriptParse,
            (wchar_t *)pScriptUTF16,
            0, 0, 0, 0, 0, 0, 0, 0);
    if (hr != S_OK)
    {
        AXSH_SetTclResultToHRESULTErrString(interp, buffer, sizeof(buffer), hr,
            "IActiveScriptParse::ParseScriptText");
        return TCL_ERROR;
    }

    /* no error */
    return TCL_OK;
}

int AXSH_Tcl_SetScriptState(ClientData clientData, Tcl_Interp *interp,
                            int objc, Tcl_Obj *CONST objv[])
{
    AXSH_EngineState *pEngineState = (AXSH_EngineState *)clientData;
    HRESULT hr;
    int ret;
    int scriptStateAsInt;
    SCRIPTSTATE scriptState;
    char *scriptStateTable[] = {"uninitialized", "started",
        "connected", "disconnected", "closed", "initialized", NULL};
    char buffer[128];

    if (objc != 3)
    {
        Tcl_WrongNumArgs(interp, 2, objv, "scriptstate");
        return TCL_ERROR;
    }

    /* determine scriptstate */
    ret = Tcl_GetIndexFromObj(interp, objv[2],
        scriptStateTable, "scriptstate", 0, &scriptStateAsInt);
    scriptState = (SCRIPTSTATE)scriptStateAsInt;
    if (ret != TCL_OK)
        return TCL_ERROR;

    /* set script state */
    hr = pEngineState->pActiveScript->lpVtbl->
        SetScriptState(pEngineState->pActiveScript, scriptState);
    if (hr != S_OK)
    {
        AXSH_SetTclResultToHRESULTErrString(interp, buffer, sizeof(buffer), hr,
            "IActiveScript::SetScriptState");
        return TCL_ERROR;
    }

    /* check for errors */
    // TODO in what situations do we need to check for errors from OnScriptError()? is it only after SetScriptState()? -> events?
    if (pEngineState->pErrorResult != NULL)
    {
        /* no reference counting used for this Tcl object because the object
           is used nowhere else within the interpreter - only here and when
           the error is detected */
        Tcl_SetObjResult(interp, pEngineState->pErrorResult);
        pEngineState->pErrorResult = NULL;
        return TCL_ERROR;
    }

    /* no error */
    return TCL_OK;
}
