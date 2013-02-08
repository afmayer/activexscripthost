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

static void AXSH_SetTclResultToVARIANT(Tcl_Interp *interp, VARIANT *pVariant)
{
    Tcl_Obj *pObject = NULL;

    if (pVariant->vt & VT_ARRAY)
        return; // TODO handle VT_ARRAY in AXSH_SetTclResultToVARIANT()

    if (pVariant->vt & VT_BYREF)
        return; // TODO handle VT_BYREF in AXSH_SetTclResultToVARIANT()

    /* try to convert the VARIANT to a reasonable Tcl object */
    switch (pVariant->vt & VT_TYPEMASK)
    {
    case VT_EMPTY:     /* nothing */
        break;

    case VT_NULL:      /* SQL style Null */
        break; // TODO handle VT_NULL in AXSH_SetTclResultToVARIANT()

    case VT_I2:        /* 2 byte signed int */
        pObject = Tcl_NewIntObj(V_I2(pVariant)); break;
    case VT_I4:        /* 4 byte signed int */
        pObject = Tcl_NewIntObj(V_I4(pVariant)); break;
    case VT_R4:        /* 4 byte real */
        pObject = Tcl_NewDoubleObj(V_R4(pVariant)); break;
    case VT_R8:        /* 8 byte real */
        pObject = Tcl_NewDoubleObj(V_R8(pVariant)); break;

    case VT_CY:        /* currency */
        break; // TODO handle CURRENCY type in AXSH_SetTclResultToVARIANT()
    case VT_DATE:      /* date */
        break; // TODO handle VT_DATE --> http://msdn.microsoft.com/en-us/library/vstudio/82ab7w69(v=vs.90).aspx

    case VT_BSTR:      /* OLE Automation string */
        pObject = Tcl_NewUnicodeObj(V_BSTR(pVariant),
            SysStringLen(V_BSTR(pVariant)));
        break;

    case VT_DISPATCH:  /* (IDispatch *) */
    case VT_ERROR:     /* SCODE */
        break; // TODO handle VT_DISPATCH and VT_ERROR in AXSH_SetTclResultToVARIANT()

    case VT_BOOL:      /* True=-1, False=0 */
        pObject = Tcl_NewBooleanObj(V_BOOL(pVariant) == 0 ? 0 : 1); break;

    case VT_VARIANT:   /* (VARIANT *) */
    case VT_UNKNOWN:   /* (IUnknown *) */
    case VT_DECIMAL:   /* 16 byte fixed point */
    case VT_RECORD:    /* user defined type */
        break; // TODO handle VT_VARIANT, VT_UNKNOWN, VT_DECIMAL and VT_RECORD in AXSH_SetTclResultToVARIANT()

    case VT_I1:        /* signed char */
        pObject = Tcl_NewIntObj(V_I1(pVariant)); break;
    case VT_UI1:       /* unsigned char */
        pObject = Tcl_NewIntObj(V_UI1(pVariant)); break;
    case VT_UI2:       /* unsigned short */
        pObject = Tcl_NewIntObj(V_UI2(pVariant)); break;
    case VT_UI4:       /* unsigned long */
        pObject = Tcl_NewLongObj(V_UI4(pVariant)); break;
    case VT_INT:       /* signed machine int */
        pObject = Tcl_NewIntObj(V_INT(pVariant)); break;
    case VT_UINT:      /* unsigned machine int */
        pObject = Tcl_NewLongObj(V_UINT(pVariant)); break;
    }

    if (pObject != NULL)
        Tcl_SetObjResult(interp, pObject);
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
    VARIANT *pResultVariantPtr = NULL;
    int i;
    Tcl_Obj *pScript = NULL;
    char *optionTable[] = {"-visible", "-expression", "-persistent", NULL};
    EXCEPINFO exceptionInfo = {0, 0, NULL, NULL, NULL, 0, NULL, 0, 0};

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
        int ret;
        int optionAsInt;

        ret = Tcl_GetIndexFromObj(interp, objv[i], optionTable, "option", 0,
            &optionAsInt);
        if (ret != TCL_OK)
            return TCL_ERROR;
        switch (optionAsInt)
        {
        case 0: /* -visible */
            parseScriptTextFlags |= SCRIPTTEXT_ISVISIBLE;
            break;

        case 1: /* -expression */
            parseScriptTextFlags |= SCRIPTTEXT_ISEXPRESSION;

            /* at least for the VBScript engine the result pointer must be
             * NULL when SCRIPTTEXT_ISEXPRESSION is not set */
            pResultVariantPtr = &parseResultVariant;
            break;

        case 2: /* -persistent */
            parseScriptTextFlags |= SCRIPTTEXT_ISPERSISTENT;
            break;
        }
    }

    /* the last argument is the script to be parsed */
    pScript = objv[objc-1];

    /* convert string to unicode format */
    pScriptUTF16 = Tcl_GetUnicodeFromObj(pScript, &scriptUTF16Length);

    /* initialize the VARIANT to hold the result */
    VariantInit(&parseResultVariant);

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
            pResultVariantPtr, /* result is stored here */
            &exceptionInfo);

    if (hr == DISP_E_EXCEPTION)
    {
        // TODO read exceptionInfo and set as Tcl result, return TCL_ERROR

        /* free BSTR variables in exception info */
        SysFreeString(exceptionInfo.bstrSource);
        SysFreeString(exceptionInfo.bstrDescription);
        SysFreeString(exceptionInfo.bstrHelpFile);
    }

    if (FAILED(hr))
    {
        AXSH_SetTclResultToHRESULTErrString(interp, buffer, sizeof(buffer), hr,
            "IActiveScriptParse::ParseScriptText");
        VariantClear(&parseResultVariant);
        return TCL_ERROR;
    }

    /* when the SCRIPTTEXT_ISEXPRESSION flag is set, ParseScriptText()
     * stores the result in a VARIANT (pointed to by one of the arguments),
     * try to convert it to a Tcl result */
    if (pResultVariantPtr != NULL)
        AXSH_SetTclResultToVARIANT(interp, pResultVariantPtr);

    /* clear the VARIANT holding the result - this automatically releases
     * all memory that is owned by the VARIANT (BSTRs, arrays, ...) */
    VariantClear(&parseResultVariant);

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
