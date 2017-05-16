#include "stdafx.h"

void DebugOut(LPTSTR pszFormat, ...)
{
	TCHAR szUser[2048], szBuf[4096];
	va_list argptr;
	va_start(argptr, pszFormat);

	StringCbVPrintf(szUser, sizeof(szUser), pszFormat, argptr);

	StringCbPrintf(szBuf, sizeof(szBuf), L"[%d:%d] - ", GetCurrentProcessId(), GetCurrentThreadId());

	StringCbCat(szBuf, sizeof(szBuf), szUser);
	StringCbCat(szBuf, sizeof(szBuf), L"\r\n");
	OutputDebugString(szBuf);
} // DebugOut
