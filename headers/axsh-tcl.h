#ifndef AXSH_TCL_H_
#define AXSH_TCL_H_

#include "axsh-include-all.h"

#define AXSH_NAMESPACE "activexscripthost"
#define AXSH_VERSIONMAJOR  1 /* IMPORTANT: keep the version number */
#define AXSH_VERSIONMINOR  0 /*       consistent with pkgIndex.tcl */
#define QUOTESTR_(x) #x
#define QUOTESTR(x) QUOTESTR_(x)

#ifdef __cplusplus
extern "C" {
#endif

/* initialization function called by Tcl on "load" */
__declspec(dllexport)
int Activexscripthost_Init(Tcl_Interp *interp);

int AXSH_Tcl_OpenEngine(ClientData clientData, Tcl_Interp *interp, int objc,
                        Tcl_Obj *CONST objv[]);

int AXSH_Tcl_EngineCommandProc(ClientData clientData, Tcl_Interp *interp, int objc,
                        Tcl_Obj *CONST objv[]);

#ifdef __cplusplus
}
#endif

#endif /* #ifndef AXSH_TCL_H_ */
