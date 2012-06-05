#ifndef AXSH_ACTIVESCRIPTSITE_IMPL_H_
#define AXSH_ACTIVESCRIPTSITE_IMPL_H_

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

void AXSH_InitActiveScriptSite(AXSH_TclActiveScriptSite *this,
                               AXSH_EngineState *pEngineState);

#ifdef __cplusplus
}
#endif

#endif /* #ifndef AXSH_ACTIVESCRIPTSITE_IMPL_H_ */
