/* ActiveX Script Host
 * (C) 2012 Alexander F. Mayer
 */
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
    char *pRetString;
    char buffer[128];

    pRetString = AXSH_CleanupEngineState(pEngineState);
    if (pRetString != NULL)
    {
        _snprintf(buffer, sizeof(buffer),
            "AXSH_CleanupEngineState() says '%s'", pRetString);
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
    unsigned int parseScriptTextFlags = 0;
    VARIANT parseResultVariant;
    int i;
    Tcl_Obj *pScript = NULL;

    if (objc < 3)
    {
        Tcl_WrongNumArgs(interp, 2, objv, "?-visible? ?-expression? "
            "?-persistent? script");
        return TCL_ERROR;
    }

    /* evaluate arguments up to the previous to last and set flags
     * for ParseScriptText():
     *     -visible        ->    SCRIPTTEXT_ISVISIBLE
     *     -expression     ->    SCRIPTTEXT_ISEXPRESSION
     *     -persistent     ->    SCRIPTTEXT_ISPERSISTENT */
    for (i=2; i < objc - 1; i++)
    {
        if (!strcmp(Tcl_GetStringFromObj(objv[i], NULL), "-visible"))
            parseScriptTextFlags |= SCRIPTTEXT_ISVISIBLE;
        else if (!strcmp(Tcl_GetStringFromObj(objv[i], NULL), "-expression"))
            parseScriptTextFlags |= SCRIPTTEXT_ISEXPRESSION;
        else if (!strcmp(Tcl_GetStringFromObj(objv[i], NULL), "-persistent"))
            parseScriptTextFlags |= SCRIPTTEXT_ISPERSISTENT;
    }

    /* the last argument is the script to be parsed */
    pScript = objv[objc-1];

    /* convert string to unicode format */
    pScriptUTF16 = Tcl_GetUnicodeFromObj(pScript, &scriptUTF16Length);

    /* call engine's ParseScriptText() method */
    hr = pEngineState->pActiveScriptParse->lpVtbl->
            ParseScriptText(pEngineState->pActiveScriptParse,
            (wchar_t *)pScriptUTF16, /* script to be parsed */
            NULL, /* item name */
            NULL, /* context */
            NULL, /* end-of-script delimiter */
            0, /* source context cookie */
            0, /* start line number */
            parseScriptTextFlags,
            &parseResultVariant, /* result is stored here */
            NULL);
    if (FAILED(hr))
    {
        AXSH_SetTclResultToHRESULTErrString(interp, buffer, sizeof(buffer), hr,
            "IActiveScriptParse::ParseScriptText");
        return TCL_ERROR;
    }

    /* when the SCRIPTTEXT_ISEXPRESSION flag is set, ParseScriptText()
     * stores the result in a VARIANT (pointed to by one of the arguments) */
    // TODO read result VARIANT and store into Tcl result

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
    if (FAILED(hr))
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
