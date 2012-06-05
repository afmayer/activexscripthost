#ifndef AXSH_HOSTCONTROL_IMPL_
#define AXSH_HOSTCONTROL_IMPL_

#include "axsh-include-all.h"

#ifdef __cplusplus
extern "C" {
#endif

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
};

typedef struct AXSH_TclHostControl_ {
    /* interfaces */
    ITclHostControl             hostCtl;
    IProvideMultipleClassInfo   multiClassInfo;

    /* private data of object */
    unsigned int      referenceCount;
    // TODO is a pointer back to the engine state needed, as in TclActiveScriptSite?
} AXSH_TclHostControl;

void AXSH_InitHostControl(AXSH_TclHostControl *this);

#ifdef __cplusplus
}
#endif

#endif /* #ifndef AXSH_HOSTCONTROL_IMPL_ */
