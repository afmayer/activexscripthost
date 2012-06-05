#ifndef AXSH_ENGINE_INTERACT_H_
#define AXSH_ENGINE_INTERACT_H_

// TODO this module should contain all functions related to the AXSH_EngineState structure
//      init = init all struct elements
//      close = free all resources pointed by engine state
#include "axsh-include-all.h"

#ifdef __cplusplus
extern "C" {
#endif

typedef struct AXSH_EngineState_ {
    struct AXSH_TclActiveScriptSite_ *pTclScriptSite;
    IActiveScript         *pActiveScript;
    IActiveScriptParse    *pActiveScriptParse;

    Tcl_Interp  *pTclInterp;
    Tcl_Command  pTclCommandToken; /* this is actually a pointer */
    Tcl_Obj     *pErrorResult;
} AXSH_EngineState;

char * AXSH_OpenScriptEngine(AXSH_EngineState *pEngineState, GUID *pEngineGuid, // TODO NAME SHOULD BE: InitScriptState()
                             Tcl_Interp *pTclInterp);
char * AXSH_CloseScriptEngine(AXSH_EngineState *pEngineState); // TODO NAME SOULD BE: CleanupScriptState()

#ifdef __cplusplus
}
#endif

#endif /* #ifndef AXSH_ENGINE_INTERACT_H_ */
