#ifndef AXSH_ACTIVESCRIPTSITE_IMPL_H_
#define AXSH_ACTIVESCRIPTSITE_IMPL_H_
/* ActiveX Script Host
 * (C) 2012 Alexander F. Mayer
 */

#include "axsh-include-all.h"

#ifdef __cplusplus
extern "C" {
#endif

typedef struct AXSH_TclActiveScriptSite_ {
    /* interfaces */
    IActiveScriptSite       site;    /* IActiveScriptSite base object */
    IActiveScriptSiteWindow siteWnd; /* IActiveScriptSiteWindow sub-object */

    /* private data of object */
    unsigned int      referenceCount;
    AXSH_EngineState *pEngineState;  /* pointer back to the engine state */
} AXSH_TclActiveScriptSite;

AXSH_TclActiveScriptSite * AXSH_CreateTclActiveScriptSite(
            AXSH_EngineState *pEngineState);

#ifdef __cplusplus
}
#endif

#endif /* #ifndef AXSH_ACTIVESCRIPTSITE_IMPL_H_ */
