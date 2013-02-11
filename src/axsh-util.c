/* ActiveX Script Host
 * (C) 2012 Alexander F. Mayer
 */
#include "axsh-include-all.h"

#define CASE_X_RETURN_X(str) case str: return #str

static int AXSH_GetValueFromHexCipher(const char c, unsigned int *pValue)
{
    if (c >= '0' && c <= '9')
    {
        *pValue = (unsigned int)(c) - (unsigned int)'0';
        return 0;
    }

    if (c >= 'a' && c <= 'f')
    {
        *pValue = (unsigned int)(c) - (unsigned int)'a' + 10;
        return 0;
    }

    if (c >= 'A' && c <= 'F')
    {
        *pValue = (unsigned int)(c) - (unsigned int)'A' + 10;
        return 0;
    }

    /* invalid character */
    return 1;
}

char * AXSH_StringToGuid(const char *pGuidString, GUID *pGuid)
{
    GUID tempGuid;
    int retVal;
    unsigned int i;
    unsigned int cipherValue;
    unsigned int data4ArrayIndex;

    memset(&tempGuid, 0, sizeof(tempGuid));

    if (strlen(pGuidString) != 38)
        return "invalid GUID length";

    if (pGuidString[0] != '{' || pGuidString[37] != '}')
        return "not a valid GUID";

    for (i=1; i <= 8; i++)
    {
        retVal = AXSH_GetValueFromHexCipher(pGuidString[i], &cipherValue);
        if (retVal != 0)
            return "GUID contains invalid cipher (segment 1)";
        tempGuid.Data1 = (tempGuid.Data1 << 4) + cipherValue;
    }

    for (i=10; i <= 13; i++)
    {
        retVal = AXSH_GetValueFromHexCipher(pGuidString[i], &cipherValue);
        if (retVal != 0)
            return "GUID contains invalid cipher (segment 2)";
        tempGuid.Data2 = (tempGuid.Data2 << 4) + cipherValue;
    }

    for (i=15; i <= 18; i++)
    {
        retVal = AXSH_GetValueFromHexCipher(pGuidString[i], &cipherValue);
        if (retVal != 0)
            return "GUID contains invalid cipher (segment 3)";
        tempGuid.Data3 = (tempGuid.Data3 << 4) + cipherValue;
    }

    for (i=20; i <= 23; i++)
    {
        retVal = AXSH_GetValueFromHexCipher(pGuidString[i], &cipherValue);
        data4ArrayIndex = (i < 22) ? 0 : 1;
        if (retVal != 0)
            return "GUID contains invalid cipher (segment 4)";
        tempGuid.Data4[data4ArrayIndex] =
            (tempGuid.Data4[data4ArrayIndex] << 4) + cipherValue;
    }

    data4ArrayIndex = 2;
    for (i=25; i <= 36; i++)
    {
        retVal = AXSH_GetValueFromHexCipher(pGuidString[i], &cipherValue);
        if (retVal != 0)
            return "GUID contains invalid cipher (segment 5)";
        tempGuid.Data4[data4ArrayIndex] =
            (tempGuid.Data4[data4ArrayIndex] << 4) + cipherValue;
        if (i % 2 == 0) data4ArrayIndex++;
    }

    /* assign output GUID */
    *pGuid = tempGuid;

    /* no error */
    return NULL;
}

// TODO use FormatMessage() Windows API function instead?
// TODO check https://en.wikipedia.org/wiki/HRESULT
char * AXSH_HRESULT2String(HRESULT hr)
{
    switch (hr)
    {
        CASE_X_RETURN_X(S_OK);

        CASE_X_RETURN_X(E_UNEXPECTED);
        CASE_X_RETURN_X(E_NOTIMPL);
        CASE_X_RETURN_X(E_OUTOFMEMORY);
        CASE_X_RETURN_X(E_INVALIDARG);
        CASE_X_RETURN_X(E_NOINTERFACE);
        CASE_X_RETURN_X(E_POINTER);
        CASE_X_RETURN_X(E_HANDLE);
        CASE_X_RETURN_X(E_ABORT);
        CASE_X_RETURN_X(E_FAIL);
        CASE_X_RETURN_X(E_ACCESSDENIED);
        //CASE_X_RETURN_X(OLESCRIPT_S_PENDING);
        CASE_X_RETURN_X(OLESCRIPT_E_SYNTAX);
        CASE_X_RETURN_X(S_FALSE);
        CASE_X_RETURN_X(DISP_E_EXCEPTION);
        // TODO clean up here... maybe sort by facility
    }

    return "(unknown HRESULT)";
}

char * AXSH_GetEngineCLSIDFromProgID(const wchar_t *pProgIDStringUTF16,
                                     GUID * pGuid)
{
    ICatInformation *pCatMgrObj;
    CATID requestedCategories[1];
    IEnumCLSID *pScriptEngines;
    HRESULT hr;
    ULONG numberOfReturnedEngines;
    int desiredProgIDWasFound;

    desiredProgIDWasFound = 0;
    hr = CoCreateInstance(&CLSID_StdComponentCategoriesMgr, 0, CLSCTX_ALL,
        &IID_ICatInformation, (void **)&pCatMgrObj);
    if (FAILED(hr))
        return "error creating category manager";

    /* get a list of engines */
    requestedCategories[0] = CATID_ActiveScriptParse;
    hr = pCatMgrObj->lpVtbl->EnumClassesOfCategories(pCatMgrObj, 1,
        requestedCategories, 0, 0, &pScriptEngines);
    if (FAILED(hr))
    {
        pCatMgrObj->lpVtbl->Release(pCatMgrObj);
        return "error enumerating engines";
    }

    /* step through engines one by one to find desired ProgID */
    do
    {
        CLSID currentEngine[1];

        numberOfReturnedEngines = 0;
        hr = pScriptEngines->lpVtbl->Next(pScriptEngines, 1, &currentEngine[0],
            &numberOfReturnedEngines);
        if (FAILED(hr))
        {
            pScriptEngines->lpVtbl->Release(pScriptEngines);
            pCatMgrObj->lpVtbl->Release(pCatMgrObj);
            return "error iterating through engines";
        }

        /* compare ProgID with desired ProgID */
        if (numberOfReturnedEngines != 0)
        {
            wchar_t *pCurrentProgIDString;
            hr = ProgIDFromCLSID(&currentEngine[0], &pCurrentProgIDString);
            if (SUCCEEDED(hr))
            {
                if (!wcscmp(pCurrentProgIDString, pProgIDStringUTF16))
                {
                    *pGuid = currentEngine[0];
                    desiredProgIDWasFound = 1;
                }
            }
            CoTaskMemFree(pCurrentProgIDString);

            if (desiredProgIDWasFound != 0)
                break;
        }
    } while (numberOfReturnedEngines != 0);

    pScriptEngines->lpVtbl->Release(pScriptEngines);
    pCatMgrObj->lpVtbl->Release(pCatMgrObj);

    if (desiredProgIDWasFound == 0)
        return "ProgID not found";
    
    /* no error */
    return NULL;
}

Tcl_Obj * AXSH_VariantToTclObj(VARIANT *pVariant)
{
    Tcl_Obj *pObject = NULL;

    if (pVariant->vt & VT_ARRAY)
        return NULL; // TODO handle VT_ARRAY in AXSH_VariantToTclObj()

    if (pVariant->vt & VT_BYREF)
        return NULL; // TODO handle VT_BYREF in AXSH_VariantToTclObj()

    switch (pVariant->vt & VT_TYPEMASK)
    {
    case VT_EMPTY:     /* nothing */
        break;

    case VT_NULL:      /* SQL style Null */
        break; // TODO handle VT_NULL in AXSH_VariantToTclObj()

    case VT_I2:        /* 2 byte signed int */
        pObject = Tcl_NewIntObj(V_I2(pVariant)); break;
    case VT_I4:        /* 4 byte signed int */
        pObject = Tcl_NewIntObj(V_I4(pVariant)); break;
    case VT_R4:        /* 4 byte real */
        pObject = Tcl_NewDoubleObj(V_R4(pVariant)); break;
    case VT_R8:        /* 8 byte real */
        pObject = Tcl_NewDoubleObj(V_R8(pVariant)); break;

    case VT_CY:        /* currency */
        break; // TODO handle CURRENCY type in AXSH_VariantToTclObj()
    case VT_DATE:      /* date */
        break; // TODO handle VT_DATE --> http://msdn.microsoft.com/en-us/library/vstudio/82ab7w69(v=vs.90).aspx

    case VT_BSTR:      /* OLE Automation string */
        pObject = Tcl_NewUnicodeObj(V_BSTR(pVariant),
            SysStringLen(V_BSTR(pVariant)));
        break;

    case VT_DISPATCH:  /* (IDispatch *) */
    case VT_ERROR:     /* SCODE */
        break; // TODO handle VT_DISPATCH and VT_ERROR in AXSH_VariantToTclObj()

    case VT_BOOL:      /* True=-1, False=0 */
        pObject = Tcl_NewBooleanObj(V_BOOL(pVariant) == 0 ? 0 : 1); break;

    case VT_VARIANT:   /* (VARIANT *) */
    case VT_UNKNOWN:   /* (IUnknown *) */
    case VT_DECIMAL:   /* 16 byte fixed point */
    case VT_RECORD:    /* user defined type */
        break; // TODO handle VT_VARIANT, VT_UNKNOWN, VT_DECIMAL and VT_RECORD in AXSH_VariantToTclObj()

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

    return pObject;
}
