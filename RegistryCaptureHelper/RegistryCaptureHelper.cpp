#include "stdafx.h"
#include "RegistryCaptureHelper.h"

extern "C"
{
	typedef DWORD(__stdcall *pfnDllRegisterReserver)();

	HKEY hkHive, hkRedirMachine, hkRedirClassesRoot, hkRedirCurrentUser;
	LPCWSTR hiveName = L"wlplushive";

	BOOL SetPrivileges();
	void TraverseKey(HKEY hKey, BOOL fLookForHelpdir);

	REGISTRYCAPTUREHELPER_API void __stdcall TestRegWrites(LPCWSTR keyName)
	{
		DebugOut(L"TestRegWrites called...");

		HKEY hk;
		LSTATUS lRet = RegCreateKey(HKEY_LOCAL_MACHINE, keyName, &hk);
		if (lRet != ERROR_SUCCESS) { DebugOut(L"RegCreateKey failed = %d", lRet); goto err0; }

		RegCloseKey(hk);
	err0:
		return;
	}

	REGISTRYCAPTUREHELPER_API BOOL __stdcall StartCapture(void)
	{
		LSTATUS lRet;

		DebugOut(L"StartCapture...");
		auto fRet = SetPrivileges();
		if (!fRet) goto err0;

		CoInitialize(NULL);

		lRet = RegLoadKey(HKEY_LOCAL_MACHINE, hiveName, L"ActiveHive.dat");
		if (lRet != ERROR_SUCCESS) { DebugOut(L"RegLoadKey failed = %d", lRet); goto err0; }

		lRet = RegOpenKey(HKEY_LOCAL_MACHINE, hiveName, &hkHive);
		if (lRet != ERROR_SUCCESS) { DebugOut(L"RegOpenKey failed = %d", lRet); goto err0; }

		lRet = RegCreateKey(hkHive, L"REGISTRY\\MACHINE", &hkRedirMachine);
		if (lRet != ERROR_SUCCESS) { DebugOut(L"RegCreateKey(LM) failed = %d", lRet); goto err0; }

		lRet = RegCreateKey(hkHive, L"REGISTRY\\MACHINE\\SOFTWARE\\Classes", &hkRedirClassesRoot);
		if (lRet != ERROR_SUCCESS) { DebugOut(L"RegCreateKey(CR) failed = %d", lRet); goto err0; }

		lRet = RegCreateKey(hkHive, L"REGISTRY\\USER", &hkRedirCurrentUser);
		if (lRet != ERROR_SUCCESS) { DebugOut(L"RegCreateKey(CU) failed = %d", lRet); goto err0; }

		lRet = RegOverridePredefKey(HKEY_LOCAL_MACHINE, hkRedirMachine);
		if (lRet != ERROR_SUCCESS) { DebugOut(L"RegOverridePredefKey(HKLM) failed = %d", lRet); goto err0; }

		lRet = RegOverridePredefKey(HKEY_CLASSES_ROOT, hkRedirClassesRoot);
		if (lRet != ERROR_SUCCESS) { DebugOut(L"RegOverridePredefKey(HKCR) failed = %d", lRet); goto err0; }

		lRet = RegOverridePredefKey(HKEY_CURRENT_USER, hkRedirCurrentUser);
		if (lRet != ERROR_SUCCESS) { DebugOut(L"RegOverridePredefKey(HKCU) failed = %d", lRet); goto err0; }

		return TRUE;
	err0:
		return FALSE;
	}

	REGISTRYCAPTUREHELPER_API BOOL __stdcall StopCapture(void)
	{
		LSTATUS lRet;

		DebugOut(L"StopCapture...");
		DebugOut(L"Remove registry redirects...");

		lRet = RegOverridePredefKey(HKEY_CURRENT_USER, NULL);
		if (lRet != ERROR_SUCCESS) { DebugOut(L"RegOverridePredefKey(HKCU) failed = %d", lRet); goto err0; }

		lRet = RegOverridePredefKey(HKEY_CLASSES_ROOT, NULL);
		if (lRet != ERROR_SUCCESS) { DebugOut(L"RegOverridePredefKey(HKCR) failed = %d", lRet); goto err0; }

		lRet = RegOverridePredefKey(HKEY_LOCAL_MACHINE, NULL);
		if (lRet != ERROR_SUCCESS) { DebugOut(L"RegOverridePredefKey(HKLM) failed = %d", lRet); goto err0; }

		DebugOut(L"Close registry keys");
		lRet = RegCloseKey(hkRedirCurrentUser);
		if (lRet != ERROR_SUCCESS) { DebugOut(L"RegCloseKey(hkRedirCurrentUser) failed = %d", lRet); }

		lRet = RegCloseKey(hkRedirClassesRoot);
		if (lRet != ERROR_SUCCESS) { DebugOut(L"RegCloseKey(hkRedirClassesRoot) failed = %d", lRet); }

		lRet = RegCloseKey(hkRedirMachine);
		if (lRet != ERROR_SUCCESS) { DebugOut(L"RegCloseKey(hkRedirMachine) failed = %d", lRet); }

		lRet = RegFlushKey(hkHive);
		if (lRet != ERROR_SUCCESS) { DebugOut(L"RegFlushKey failed = %d", lRet); }

		lRet = RegCloseKey(hkHive);
		if (lRet != ERROR_SUCCESS) { DebugOut(L"RegCloseKey(hkHive) failed = %d", lRet); }

		DebugOut(L"Unload hive with redirects");
		lRet = RegUnLoadKey(HKEY_LOCAL_MACHINE, hiveName);
		if (lRet != ERROR_SUCCESS) { DebugOut(L"RegUnloadKey failed = %d", lRet); goto err0; }

		return TRUE;
	err0:
		return FALSE;
	}

	REGISTRYCAPTUREHELPER_API BOOL __stdcall RegisterNativeCOMDll(LPCWSTR dllFileName)
	{
		BOOL ret = TRUE;
		DebugOut(L"Registering COM component - %s", dllFileName);
		auto hMod = LoadLibrary(dllFileName);
		if (hMod == NULL) { DebugOut(L"LoadLibrary failed = %d", GetLastError()); ret = FALSE; goto err0; }
		DebugOut(L"\tDLL loaded");

		auto DllRegisterServer = (pfnDllRegisterReserver)GetProcAddress(hMod, "DllRegisterServer");
		if (DllRegisterServer == NULL) { DebugOut(L"GetProcAddress failed = %d", GetLastError()); ret = FALSE;  goto err1; }
		DebugOut(L"\tEntry point retrieved");

		auto dwRet = DllRegisterServer();
		if (dwRet != ERROR_SUCCESS) {
			DebugOut(L"\tDllRegisterServer failed = %X", dwRet);
		}
	err1:
		FreeLibrary(hMod);
	err0:
		return ret;
	}

	REGISTRYCAPTUREHELPER_API BOOL __stdcall UnloadHive()
	{
		// We end up here after the unload have failed during the regular capture process (due to open keys in the hive).
		// This is called from the parent instance of the app to properly unload the hive which kan only be done after all keys are closed
		// which will happen upon process termination cleanup in the child process doing the real work.

		SetPrivileges();
		LSTATUS lRet = RegUnLoadKey(HKEY_LOCAL_MACHINE, hiveName);
		if (lRet != ERROR_SUCCESS) {
			DebugOut(L"RegUnloadKey failed = %d", lRet);
		} else {
			DebugOut(L"RegUnloadKey succeeded");
		}
		return TRUE;
	}

	BOOL SetPrivilege(HANDLE hToken, LPCTSTR lpszPrivilege, BOOL bEnablePrivilege)
	{
		TOKEN_PRIVILEGES tp;
		LUID luid;

		if (!LookupPrivilegeValue(NULL, lpszPrivilege, &luid)) {
			DebugOut(L"LookupPrivilegeValue error: %u\n", GetLastError());
			return FALSE;
		}

		tp.PrivilegeCount = 1;
		tp.Privileges[0].Luid = luid;
		if (bEnablePrivilege)
			tp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED;
		else
			tp.Privileges[0].Attributes = 0;

		// Enable the privilege or disable all privileges.
		if (!AdjustTokenPrivileges(hToken, FALSE, &tp, sizeof(TOKEN_PRIVILEGES), (PTOKEN_PRIVILEGES)NULL, (PDWORD)NULL)) {
			DebugOut(L"AdjustTokenPrivileges error: %u\n", GetLastError());
			return FALSE;
		}

		if (GetLastError() == ERROR_NOT_ALL_ASSIGNED) {
			DebugOut(L"The token does not have the specified privilege. \n");
			return FALSE;
		}
		return TRUE;
	}

	BOOL SetPrivileges()
	{
		BOOL fRet;
		HANDLE hToken;
		fRet = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hToken);
		if (!fRet) { DebugOut(L"OpenProcessToken failed = %d", GetLastError()); goto err0; }

		fRet = SetPrivilege(hToken, SE_BACKUP_NAME, TRUE);
		if (!fRet) goto err0;

		fRet = SetPrivilege(hToken, SE_RESTORE_NAME, TRUE);
	err0:
		return fRet;
	}

	#define MAX_KEY_LENGTH 255
	#define MAX_VALUE_NAME 16383

	LPCWSTR g_DllDirectory;

	void PatchRegValuePath(const HKEY hKehyParent, const HKEY hKeySub, TCHAR achKey[255], LPCWSTR achValueName, BOOL fLookForHelpdir)
	{
		LSTATUS retCode;
		DWORD dwType, cbData;
		TCHAR  achValue[MAX_VALUE_NAME];

		if (fLookForHelpdir && wcscmp(achKey, L"HELPDIR") == 0) {
			retCode = RegQueryValueEx(hKeySub, NULL, NULL, &dwType, (LPBYTE)achValue, &cbData);
			if (retCode == ERROR_SUCCESS) {
				std::wstring directory(achValue);
				directory += L"\\";
				std::wstring newPath;
				if (wcscmp(achValue, g_DllDirectory) == 0) {
					newPath = std::wstring(L"[{AppVPackageRoot}]WLPlus\\");
					RegSetValueEx(hKeySub, NULL, 0, REG_SZ, (LPBYTE)(newPath.data()), newPath.size() * sizeof(wchar_t));
				} else if (wcscmp(directory.data(), g_DllDirectory) == 0) {
					newPath = std::wstring(L"[{AppVPackageRoot}]\\WLPlus");
				}
				if (newPath.size() > 0)
					RegSetValueEx(hKeySub, NULL, 0, REG_SZ, (LPBYTE)(newPath.data()), newPath.size() * sizeof(wchar_t));
			}
		} else if (wcscmp(achKey, achValueName) == 0) {
			retCode = RegQueryValueEx(hKeySub, NULL, NULL, &dwType, (LPBYTE)achValue, &cbData);
			if (retCode == ERROR_SUCCESS) {
				std::wstring dllFullPath(achValue);
				auto newPath = std::wstring(L"[{AppVPackageRoot}]\\WLPlus\\");

				auto currentPath = std::experimental::filesystem::path(dllFullPath);
				auto dllFileName = currentPath.filename();
				newPath = newPath.append(dllFileName);

				RegSetValueEx(hKeySub, NULL, 0, REG_SZ, (LPBYTE)(newPath.data()), newPath.size() * sizeof(wchar_t));
			}
		} else {
			TraverseKey(hKeySub, fLookForHelpdir);
		}
	}

	void TraverseKey(HKEY hKey, BOOL fLookForHelpdir)
	{
		TCHAR    achKey[MAX_KEY_LENGTH];   // buffer for subkey name
		DWORD    cbName;                   // size of name string 
		TCHAR    achClass[MAX_PATH] = TEXT("");  // buffer for class name 
		DWORD    cchClassName = MAX_PATH;  // size of class string 
		DWORD    cSubKeys = 0;               // number of subkeys 
		DWORD    cbMaxSubKey;              // longest subkey size 
		DWORD    cchMaxClass;              // longest class string 
		DWORD    cValues;              // number of values for key 
		DWORD    cchMaxValue;          // longest value name 
		DWORD    cbMaxValueData;       // longest value data 
		DWORD    cbSecurityDescriptor; // size of security descriptor 
		FILETIME ftLastWriteTime;      // last write time 

		DWORD i, retCode;

		// Get the class name and the value count. 
		retCode = RegQueryInfoKey(
			hKey,                    // key handle 
			achClass,                // buffer for class name 
			&cchClassName,           // size of class string 
			NULL,                    // reserved 
			&cSubKeys,               // number of subkeys 
			&cbMaxSubKey,            // longest subkey size 
			&cchMaxClass,            // longest class string 
			&cValues,                // number of values for this key 
			&cchMaxValue,            // longest value name 
			&cbMaxValueData,         // longest value data 
			&cbSecurityDescriptor,   // security descriptor 
			&ftLastWriteTime);       // last write time 

		// Enumerate the subkeys, until RegEnumKeyEx fails.

		if (cSubKeys) {
			DebugOut(L"Number of subkeys: %d", cSubKeys);

			for (i = 0; i<cSubKeys; i++) {
				cbName = MAX_KEY_LENGTH;
				retCode = RegEnumKeyEx(hKey, i, achKey, &cbName, NULL, NULL, NULL, &ftLastWriteTime);
				if (retCode == ERROR_SUCCESS) {
					HKEY hKeySub;
					DebugOut(L"(%d) %s", i + 1, achKey);
					retCode = RegOpenKey(hKey, achKey, &hKeySub);
					if (retCode != ERROR_SUCCESS) {
						DebugOut(L"RegOpenKey failed = %d", retCode);
					} else {
						PatchRegValuePath(hKey, hKeySub, achKey, L"InprocServer32", fLookForHelpdir);
						PatchRegValuePath(hKey, hKeySub, achKey, L"ToolboxBitmap32", fLookForHelpdir);
						PatchRegValuePath(hKey, hKeySub, achKey, L"win32", fLookForHelpdir);
						PatchRegValuePath(hKey, hKeySub, achKey, L"HELPDIR", fLookForHelpdir);
						RegCloseKey(hKeySub);
					}
				}
			}
		}
	}

	REGISTRYCAPTUREHELPER_API BOOL __stdcall PatchNativeCOMDllPaths(LPCWSTR DllDirectory)
	{
		HKEY hkCLSID, hkTypeLib;
		g_DllDirectory = DllDirectory;

		LSTATUS lRet = RegOpenKey(hkHive, L"REGISTRY\\MACHINE\\SOFTWARE\\Classes\\CLSID", &hkCLSID);
		if (lRet != ERROR_SUCCESS) {
			DebugOut(L"RegCreateKey(Classes\\CLSID) failed = %d", lRet);
		} else {
			TraverseKey(hkCLSID, FALSE);
			RegCloseKey(hkCLSID);
		}

		lRet = RegOpenKey(hkHive, L"REGISTRY\\MACHINE\\SOFTWARE\\Classes\\TypeLib", &hkTypeLib);
		if (lRet != ERROR_SUCCESS) {
			DebugOut(L"RegCreateKey(Classes\\TypeLib) failed = %d", lRet);
		} else {
			TraverseKey(hkTypeLib, TRUE);
			RegCloseKey(hkTypeLib);
		}
		return TRUE;
	}
}