#ifndef AXSH_HOSTCONTROL_IMPL_
#define AXSH_HOSTCONTROL_IMPL_
/* ActiveX Script Host
 * (C) 2012 Alexander F. Mayer
 */

#include "axsh-include-all.h"

#ifdef __cplusplus
extern "C" {
#endif

extern const CLSID CLSID_ITclHostControl;
extern const IID IID_ITclHostControl;

#undef  INTERFACE
#define INTERFACE ITclHostControl
DECLARE_INTERFACE_ (INTERFACE, IDispatch)
{
    /* IUnknown */
    STDMETHOD  (QueryInterface) (THIS_ REFIID, void **) PURE;
    STDMETHOD_ (ULONG, AddRef)  (THIS) PURE;
    STDMETHOD_ (ULONG, Release) (THIS) PURE;
    /* IDispatch */
    STDMETHOD_ (ULONG, GetTypeInfoCount)(THIS_ UINT *) PURE;
    STDMETHOD_ (ULONG, GetTypeInfo)     (THIS_ UINT, LCID, ITypeInfo **) PURE;
    STDMETHOD_ (ULONG, GetIDsOfNames)   (THIS_ REFIID, LPOLESTR *, UINT, LCID,
        DISPID *) PURE;
    STDMETHOD_ (ULONG, Invoke)          (THIS_ DISPID, REFIID, LCID, WORD,
        DISPPARAMS *, VARIANT *, EXCEPINFO *, UINT *) PURE;
    /* ITclHostControl */
    STDMETHOD  (GetStringVar)   (THIS_ BSTR, BSTR *) PURE;
    STDMETHOD  (GetIntVar)      (THIS_ BSTR, INT *) PURE;
};

typedef struct AXSH_TclHostControl_ {
    /* interfaces */
    ITclHostControl             hostCtl; // TODO all interfaces get Ifc suffix
    IProvideMultipleClassInfo   multiClassInfo;

    /* private data of object */
    unsigned int      referenceCount;
    ITypeInfo         *pObjectTypeInfo;
    ITypeInfo         *pInterfaceTypeInfo;
    AXSH_EngineState  *pEngineState; /* pointer back to the engine state */
} AXSH_TclHostControl;

AXSH_TclHostControl * AXSH_CreateTclHostControl(AXSH_EngineState *pEngineState);

#ifdef __cplusplus
}
#endif

#endif /* #ifndef AXSH_HOSTCONTROL_IMPL_ */
