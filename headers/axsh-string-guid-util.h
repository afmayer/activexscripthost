#ifndef AXSH_STRING_TO_GUID_H_
#define AXSH_STRING_TO_GUID_H_

#include "axsh-include-all.h"

#ifdef __cplusplus
extern "C" {
#endif

char * AXSH_StringToGuid(const char *pGuidString, GUID *pGuid);
char * AXSH_HRESULT2String(HRESULT hr);
char * AXSH_GetEngineCLSIDFromProgID(const wchar_t *pProgIDStringUTF16,
                                     GUID * pGuid);

#ifdef __cplusplus
}
#endif

#endif /* #ifndef AXSH_STRING_TO_GUID_H_ */
