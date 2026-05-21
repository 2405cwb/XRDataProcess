/*
 * MATLAB Compiler: 6.2 (R2016a)
 * Date: Fri Feb 18 15:08:12 2022
 * Arguments: "-B" "macro_default" "-W" "lib:ForSinsDR2" "-T" "link:lib" "-d"
 * "D:\002-TFS-SrcCodeMMS\002-HnPingZhengDuSoft\001-PingZhengDu\平整度程序\ForSi
 * nsDRCompile\ForSinsDR2\for_testing" "-v"
 * "D:\002-TFS-SrcCodeMMS\002-HnPingZhengDuSoft\001-PingZhengDu\平整度程序\平整?
 * 瘸绦騖TrackInspec\ForSinsDR2.m" "-a"
 * "D:\002-TFS-SrcCodeMMS\002-HnPingZhengDuSoft\001-PingZhengDu\平整度程序\平整?
 * 瘸绦騖base\userdef\psinstypedef.m" "-a"
 * "D:\002-TFS-SrcCodeMMS\002-HnPingZhengDuSoft\001-PingZhengDu\平整度程序\平整?
 * 瘸绦騖base\userdef\test_SINS_DR_def.m" 
 */

#ifndef __ForSinsDR2_h
#define __ForSinsDR2_h 1

#if defined(__cplusplus) && !defined(mclmcrrt_h) && defined(__linux__)
#  pragma implementation "mclmcrrt.h"
#endif
#include "mclmcrrt.h"
#ifdef __cplusplus
extern "C" {
#endif

#if defined(__SUNPRO_CC)
/* Solaris shared libraries use __global, rather than mapfiles
 * to define the API exported from a shared library. __global is
 * only necessary when building the library -- files including
 * this header file to use the library do not need the __global
 * declaration; hence the EXPORTING_<library> logic.
 */

#ifdef EXPORTING_ForSinsDR2
#define PUBLIC_ForSinsDR2_C_API __global
#else
#define PUBLIC_ForSinsDR2_C_API /* No import statement needed. */
#endif

#define LIB_ForSinsDR2_C_API PUBLIC_ForSinsDR2_C_API

#elif defined(_HPUX_SOURCE)

#ifdef EXPORTING_ForSinsDR2
#define PUBLIC_ForSinsDR2_C_API __declspec(dllexport)
#else
#define PUBLIC_ForSinsDR2_C_API __declspec(dllimport)
#endif

#define LIB_ForSinsDR2_C_API PUBLIC_ForSinsDR2_C_API


#else

#define LIB_ForSinsDR2_C_API

#endif

/* This symbol is defined in shared libraries. Define it here
 * (to nothing) in case this isn't a shared library. 
 */
#ifndef LIB_ForSinsDR2_C_API 
#define LIB_ForSinsDR2_C_API /* No special import/export declaration */
#endif

extern LIB_ForSinsDR2_C_API 
bool MW_CALL_CONV ForSinsDR2InitializeWithHandlers(
       mclOutputHandlerFcn error_handler, 
       mclOutputHandlerFcn print_handler);

extern LIB_ForSinsDR2_C_API 
bool MW_CALL_CONV ForSinsDR2Initialize(void);

extern LIB_ForSinsDR2_C_API 
void MW_CALL_CONV ForSinsDR2Terminate(void);



extern LIB_ForSinsDR2_C_API 
void MW_CALL_CONV ForSinsDR2PrintStackTrace(void);

extern LIB_ForSinsDR2_C_API 
bool MW_CALL_CONV mlxForSinsDR2(int nlhs, mxArray *plhs[], int nrhs, mxArray *prhs[]);



extern LIB_ForSinsDR2_C_API bool MW_CALL_CONV mlfForSinsDR2(int nargout, mxArray** posOut, mxArray** rowCount, mxArray* imu0, mxArray* dR, mxArray* lon, mxArray* height, mxArray* CaliType, mxArray* COR, mxArray* dcount);

#ifdef __cplusplus
}
#endif
#endif
