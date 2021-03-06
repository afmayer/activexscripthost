/* ActiveX Script Host
 * (C) 2012 Alexander F. Mayer
 */
#include "axsh-include-all.h"

/* global variable holding a pointer to the type library - the
 * type library is never released */
ITypeLib *g_pTypeLibrary;

static void AXSH_Tcl_EngineCommandDeleteProc(ClientData clientData)
{
    AXSH_EngineState *pEngineState = (AXSH_EngineState *)clientData;
    AXSH_CleanupEngineState(pEngineState);
    free(pEngineState);
}

static int AXSH_Tcl_EngineCommandProc(
            ClientData clientData,
            Tcl_Interp *interp,
            int objc,
            Tcl_Obj *CONST objv[])
{
    int ret;
    int subcommandIndex;
    static const char *subcommandTable[] =
        {"close", "parse", "setscriptstate", NULL};

    if (objc < 2)
    {
        Tcl_WrongNumArgs(interp, 1, objv, "subcommand ...");
        return TCL_ERROR;
    }

    /* get subcommand */
    ret = Tcl_GetIndexFromObj(interp, objv[1], subcommandTable, "subcommand",
        0, &subcommandIndex);
    if (ret != TCL_OK)
        return TCL_ERROR;

    /* delegate subcommand */
    switch (subcommandIndex)
    {
    case 0: /* close */
        return AXSH_Tcl_CloseScriptEngine(clientData, interp, objc, objv);
    case 1: /* parse */
        return AXSH_Tcl_ParseText(clientData, interp, objc, objv);
    case 2: /* setscriptstate */
        return AXSH_Tcl_SetScriptState(clientData, interp, objc, objv);
    default:
        Tcl_SetResult(interp, "internal error - unhandled subcommand",
            TCL_STATIC);
        return TCL_ERROR;
    }

    /* no error - but actually we should not end up here, evereyone should
       have returned in the switch block */
    return TCL_OK;
}

static int AXSH_Tcl_OpenEngine(
            ClientData clientData,
            Tcl_Interp *interp,
            int objc,
            Tcl_Obj *CONST objv[])
{
    AXSH_EngineState *pEngineState;
    char *pRetString;
    char buffer[128];
    // VBScript: {B54F3741-5B07-11cf-A4B0-00AA004A55E8}
    char *pEngineStringUTF8;
    GUID engineGuid;
    Tcl_Command pTempCmdToken;
    HRESULT hr;

    /* check if called with exactly one argument */
    if (objc != 2)
    {
        Tcl_WrongNumArgs(interp, 1, objv, "engineGuid");
        return TCL_ERROR;
    }

    /* allocate and initialize memory for engine state */
    pEngineState = malloc(sizeof(*pEngineState));
    if (pEngineState == NULL)
    {
        Tcl_SetResult(interp, "out of memory", TCL_STATIC);
        return TCL_ERROR;
    }
    memset(pEngineState, 0, sizeof(*pEngineState));

    pEngineStringUTF8 = Tcl_GetStringFromObj(objv[1], NULL);

    /* engine can be specified as CLSID or by engine name */
    if (pEngineStringUTF8[0] == '{')
    {
        /* convert string to GUID structure
           we don't need to convert this from UTF-8 to ASCII because all
           valid characters are the same in UTF-8 and ASCII */
        pRetString = AXSH_StringToGuid(pEngineStringUTF8, &engineGuid);
        if (pRetString != NULL)
        {
            _snprintf(buffer, sizeof(buffer), "AXSH_StringToGuid() says '%s'",
                pRetString);
            buffer[sizeof(buffer)-1] = 0;
            goto errcleanup1;
        }
    }
    else
    {
        /* try to find engine by name */
        Tcl_UniChar *pEngineStringUTF16;
        pEngineStringUTF16 = Tcl_GetUnicodeFromObj(objv[1], NULL);
        pRetString = AXSH_GetEngineCLSIDFromProgID(pEngineStringUTF16,
            &engineGuid);
        if (pRetString != NULL)
        {
            _snprintf(buffer, sizeof(buffer), "AXSH_GetEngineCLSIDFromProgID()"
                " says '%s'", pRetString);
            buffer[sizeof(buffer)-1] = 0;
            goto errcleanup1;
        }
    }

    /* open engine */
    pRetString = AXSH_InitEngineState(pEngineState, &engineGuid, interp);
    if (pRetString != NULL)
    {
        _snprintf(buffer, sizeof(buffer), // TODO replace this by HRESULT2String() result
            "AXSH_InitEngineState() says '%s'", pRetString);
        buffer[sizeof(buffer)-1] = 0;
        goto errcleanup1;
    }

    /* add named item */
    // TODO allow overriding the name for the TclHostControl named item (store Tcl_Obj)
    // TODO allow prevention of AddNamedItem()
    hr = pEngineState->pActiveScript->lpVtbl->
        AddNamedItem(pEngineState->pActiveScript, L"tcl",
        SCRIPTITEM_GLOBALMEMBERS | SCRIPTITEM_ISVISIBLE);
    if (FAILED(hr))
    {
        _snprintf(buffer, sizeof(buffer), // TODO use AXSH_SetTclResultToHRESULTErrString() instead...
            "IActiveScript::AddNamedItem returned %s", AXSH_HRESULT2String(hr));
        buffer[sizeof(buffer)-1] = 0;
        goto errcleanup2;
    }

    /* create a Tcl command for our new engine - pass a
       pointer (client data) to the engine state when the command is handled */
    _snprintf(buffer, sizeof(buffer), AXSH_NAMESPACE "::handle0x%p",
        pEngineState);
    buffer[sizeof(buffer)-1] = 0;
    pTempCmdToken = Tcl_CreateObjCommand(interp, buffer,
        AXSH_Tcl_EngineCommandProc, pEngineState,
        AXSH_Tcl_EngineCommandDeleteProc);
    if (pTempCmdToken == NULL)
    {
        _snprintf(buffer, sizeof(buffer), "error creating new Tcl command");
        buffer[sizeof(buffer)-1] = 0;
        goto errcleanup2;
    }

    /* store token for created Tcl command in engine state */
    pEngineState->pTclCommandToken = pTempCmdToken;

    /* no error - return name of new command */
    Tcl_SetResult(interp, buffer, TCL_VOLATILE);
    return TCL_OK;

    /* the cleanup-stack */
errcleanup2:
    pEngineState->pActiveScript->lpVtbl->Close(pEngineState->pActiveScript);
    pEngineState->pActiveScript->lpVtbl->        /* release ActiveScript... */
        Release(pEngineState->pActiveScript);
    pEngineState->pActiveScriptParse->lpVtbl->   /* ...and ActiveScriptParse */
        Release(pEngineState->pActiveScriptParse);
errcleanup1:
    free(pEngineState);

    /* finally return error string */
    Tcl_SetResult(interp, buffer, TCL_VOLATILE);
    return TCL_ERROR;
}

int Activexscripthost_Init(Tcl_Interp *interp)
{
    int ret;
    HRESULT hr;

    /* initialize Tcl Stubs */
    if (Tcl_InitStubs(interp, "8.4", 0) == NULL)
        return TCL_ERROR;

    /* the required type library is loaded in this function call - the working
     * directory is temporarily set to the directory conatining the DLL
     * during "load", so the filename without a path is enough for
     * LoadTypeLib() --> it's otherwise hard to discover the DLL path */
    hr = LoadTypeLib(L"activexscripthost.dll", &g_pTypeLibrary);
    if (FAILED(hr))
    {
        Tcl_SetResult(interp, "cannot load type library - loaded DLL must be "
            "in current working directory", TCL_STATIC);
        return TCL_ERROR;
    }

    Tcl_CreateObjCommand(interp, AXSH_NAMESPACE "::openengine",
        AXSH_Tcl_OpenEngine, NULL, NULL);

    /* export namespace procedures */
    Tcl_Eval(interp, "namespace eval " AXSH_NAMESPACE " {namespace export *}");

    /* provide package */
    ret = Tcl_PkgProvide(interp, AXSH_NAMESPACE, QUOTESTR(AXSH_VERSIONMAJOR)
            "." QUOTESTR(AXSH_VERSIONMINOR));
    if (ret != TCL_OK)
        return TCL_ERROR;

    return TCL_OK;
}
