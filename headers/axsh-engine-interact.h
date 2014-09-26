#ifndef AXSH_ENGINE_INTERACT_H_
#define AXSH_ENGINE_INTERACT_H_
/* ActiveX Script Host
 * (C) 2012 Alexander F. Mayer
 */

#include "axsh-include-all.h"

#ifdef __cplusplus
extern "C" {
#endif

typedef struct AXSH_EngineState_ {
    /* pointer to our object implementing IActiveScriptSite */
    struct AXSH_TclActiveScriptSite_ *pTclScriptSite;

    /* pointers to interfaces of the created script engine */
    IActiveScript         *pActiveScript;
    IActiveScriptParse    *pActiveScriptParse;

    /* pointer to our object implementing ITclHostControl */
    struct AXSH_TclHostControl_ *pTclHostControl;

    /* pointer to the Tcl interpreter */
    Tcl_Interp  *pTclInterp;

    /* pointer to the created Tcl command token */
    Tcl_Command  pTclCommandToken;

    /* pointer to a Tcl object that is used to transport an error message
     * from AXSH_TclActiveScriptSite::OnScriptError() to the Tcl command
     * function that calls the engine's SetScriptState() method */
    Tcl_Obj     *pErrorResult;
} AXSH_EngineState;

char * AXSH_InitEngineState(
            AXSH_EngineState *pEngineState,
            GUID *pEngineGuid,
            Tcl_Interp *pTclInterp);
void AXSH_CleanupEngineState(AXSH_EngineState *pEngineState);

#ifdef __cplusplus
}
#endif

#endif /* #ifndef AXSH_ENGINE_INTERACT_H_ */
