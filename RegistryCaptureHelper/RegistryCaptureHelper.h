// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the REGISTRYCAPTUREHELPER_EXPORTS
// symbol defined on the command line. This symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// REGISTRYCAPTUREHELPER_API functions as being imported from a DLL, whereas this DLL sees symbols
// defined with this macro as being exported.
#ifdef REGISTRYCAPTUREHELPER_EXPORTS
#define REGISTRYCAPTUREHELPER_API __declspec(dllexport)
#else
#define REGISTRYCAPTUREHELPER_API __stdcall __declspec(dllimport)
#endif

extern "C" {
	REGISTRYCAPTUREHELPER_API void __stdcall TestRegWrites(LPCWSTR keyName);
	REGISTRYCAPTUREHELPER_API BOOL __stdcall StartCapture();
	REGISTRYCAPTUREHELPER_API BOOL __stdcall StopCapture();
	REGISTRYCAPTUREHELPER_API BOOL __stdcall RegisterNativeCOMDll(LPCWSTR dllFileName);
	REGISTRYCAPTUREHELPER_API BOOL __stdcall PatchNativeCOMDllPaths(LPCWSTR DllDirectory);
	REGISTRYCAPTUREHELPER_API BOOL __stdcall UnloadHive();
}
