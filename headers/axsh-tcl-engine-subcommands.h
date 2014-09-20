#ifndef AXSH_TCL_ENGINE_SUBCOMMANDS_
#define AXSH_TCL_ENGINE_SUBCOMMANDS_
/* ActiveX Script Host
 * (C) 2012 Alexander F. Mayer
 */

#include "axsh-include-all.h"

#ifdef __cplusplus
extern "C" {
#endif

int AXSH_Tcl_CloseScriptEngine(
            ClientData clientData,
            Tcl_Interp *interp,
            int objc,
            Tcl_Obj *CONST objv[]);

int AXSH_Tcl_ParseText(
            ClientData clientData,
            Tcl_Interp *interp,
            int objc,
            Tcl_Obj *CONST objv[]);

int AXSH_Tcl_SetScriptState(
            ClientData clientData,
            Tcl_Interp *interp,
            int objc,
            Tcl_Obj *CONST objv[]);

#ifdef __cplusplus
}
#endif

#endif /* #ifndef AXSH_TCL_ENGINE_SUBCOMMANDS_ */
