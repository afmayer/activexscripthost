#ifndef AXSH_TCL_H_
#define AXSH_TCL_H_
/* ActiveX Script Host
 * (C) 2012 Alexander F. Mayer
 */

#include "axsh-include-all.h"

#define AXSH_NAMESPACE "activexscripthost"
#define AXSH_VERSIONMAJOR  0  /* IMPORTANT: keep the version number */
#define AXSH_VERSIONMINOR  15 /*       consistent with pkgIndex.tcl */
#define QUOTESTR_(x) #x
#define QUOTESTR(x) QUOTESTR_(x)

extern ITypeLib *g_pTypeLibrary;

#ifdef __cplusplus
extern "C" {
#endif

/* initialization function called by Tcl on "load" */
__declspec(dllexport)
int Activexscripthost_Init(Tcl_Interp *interp);

#ifdef __cplusplus
}
#endif

#endif /* #ifndef AXSH_TCL_H_ */
