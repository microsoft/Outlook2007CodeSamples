#include "stdafx.h"

STDMETHODIMP MergeWithMAPISVC()
{
	Log(true,"MergeWithMAPISVC adding wrapped PST\n");
	return HrSetProfileParameters(aWrapPSTServicesIni);
}

STDMETHODIMP RemoveFromMAPISVC()
{
	Log(true,"RemoveFromMAPISVC removing wrapped PST\n");
	return HrSetProfileParameters(aREMOVE_WrapPSTServicesIni);
}

// $--GetRegistryValue---------------------------------------------------------
// Get a registry value - allocating memory using new to hold it.
// -----------------------------------------------------------------------------
LONG GetRegistryValue(
						   IN HKEY hKey, // the key.
						   IN LPCTSTR lpszValue, // value name in key.
						   OUT DWORD* lpType, // where to put type info.
						   OUT LPVOID* lppData) // where to put the data.
{
	LONG lRes = ERROR_SUCCESS;

	Log(true,_T("GetRegistryValue()\n"),lpszValue);

	*lppData = NULL;
	DWORD cb = NULL;

	// Get its size
	lRes = RegQueryValueEx(
		hKey,
		lpszValue,
		NULL,
		lpType,
		NULL,
		&cb);

	// only handle types we know about - all others are bad
	if (ERROR_SUCCESS == lRes && cb &&
		(REG_SZ == *lpType || REG_DWORD == *lpType || REG_MULTI_SZ == *lpType))
	{
		*lppData = new BYTE[cb];

		if (*lppData)
		{
			// Get the current value
			lRes = RegQueryValueEx(
				hKey,
				lpszValue,
				NULL,
				lpType,
				(unsigned char*)*lppData,
				&cb);

			if (ERROR_SUCCESS != lRes)
			{
				delete[] *lppData;
				*lppData = NULL;
			}
		}
	}
	else lRes = ERROR_INVALID_DATA;

	return lRes;
}

///////////////////////////////////////////////////////////////////////////////
// Function name   : GetMAPISVCPath
// Description     : This will get the correct path to the MAPISVC.INF file.
// Return type     : void
// Argument        : LPSTR szMAPIDir - Buffer to hold the path to the MAPISVC file.
//                   ULONG cchMAPIDir - size of the buffer
void GetMAPISVCPath(LPTSTR szMAPIDir, ULONG cchMAPIDir)
{
	LONG lRes = S_OK;
	BOOL bRet = false;
	UINT uiLen = 0;
	HINSTANCE hinstStub = NULL;

	szMAPIDir[0] = '\0';	// Terminate String at pos 0 (safer if we fail below)

	// Call common code in mapistub.dll
	hinstStub = LoadLibrary(_T("mapistub.dll"));
	if (!hinstStub)
	{
		// Try stub mapi32.dll if mapistub.dll missing
		hinstStub = LoadLibrary(_T("mapi32.dll"));
	}
	
	if (hinstStub)
	{		
		LPFGETCOMPONENTPATH pfnFGetComponentPath;

		pfnFGetComponentPath = (LPFGETCOMPONENTPATH)
			GetProcAddress(hinstStub, "FGetComponentPath");

		if (pfnFGetComponentPath)
		{
			// Find some strings:
			LPTSTR szAppLCID = NULL;
			LPTSTR szOfficeLCID = NULL;

			HKEY hMicrosoftOutlook = NULL;

			lRes = RegOpenKeyEx(
				HKEY_LOCAL_MACHINE,
				_T("Software\\Clients\\Mail\\Microsoft Outlook"),
				NULL,
				KEY_READ,
				&hMicrosoftOutlook);

			if (hMicrosoftOutlook)
			{
				DWORD dwKeyType = NULL;

				lRes = GetRegistryValue(
					hMicrosoftOutlook,
					_T("MSIApplicationLCID"),
					&dwKeyType,
					(LPVOID*) &szAppLCID);

				lRes = GetRegistryValue(
					hMicrosoftOutlook,
					_T("MSIOfficeLCID"),
					&dwKeyType,
					(LPVOID*) &szOfficeLCID);

				RegCloseKey(hMicrosoftOutlook);
			}

			if (szAppLCID)
			{
				bRet = pfnFGetComponentPath(
					"{FF1D0740-D227-11D1-A4B0-006008AF820E}", // msmapi32.dll
					szAppLCID,
					szMAPIDir,
					cchMAPIDir,
					TRUE);
			}
			if ((!bRet || szMAPIDir[0] == _T('\0')) && szOfficeLCID)
			{
				bRet = pfnFGetComponentPath(
					"{FF1D0740-D227-11D1-A4B0-006008AF820E}", // msmapi32.dll
					szOfficeLCID,
					szMAPIDir,
					cchMAPIDir,
					TRUE);
			}
			if (!bRet || szMAPIDir[0] == _T('\0'))
			{
				bRet = pfnFGetComponentPath(
					"{FF1D0740-D227-11D1-A4B0-006008AF820E}", // msmapi32.dll
					NULL,
					szMAPIDir,
					cchMAPIDir,
					TRUE);
			}
			delete[] szOfficeLCID;
			delete[] szAppLCID;
		}
		FreeLibrary(hinstStub);
	}

	// We got the path to msmapi32.dll - need to strip it
	if (bRet && szMAPIDir[0] != _T('\0'))
	{
		LPTSTR lpszSlash = NULL;
		LPTSTR lpszCur = szMAPIDir;

		for (lpszSlash = lpszCur; *lpszCur; lpszCur = lpszCur++)
		{
			if (*lpszCur == _T('\\')) lpszSlash = lpszCur;
		}
		*lpszSlash = _T('\0');
	}

	if (szMAPIDir[0] == _T('\0'))
	{
		// Fall back on System32
		uiLen = GetSystemDirectory(szMAPIDir, cchMAPIDir);
	}

	if (uiLen && szMAPIDir[0] != _T('\0'))
	{
		StringCchPrintf(
			szMAPIDir,
			cchMAPIDir,
			_T("%s\\%s"),
			szMAPIDir,
			_T("MAPISVC.INF"));
	}
}

// $--HrSetProfileParameters----------------------------------------------
// Add values to MAPISVC.INF
// -----------------------------------------------------------------------------
STDMETHODIMP HrSetProfileParameters(SERVICESINIREC *lpServicesIni)
{
	HRESULT	hRes						= S_OK;
	TCHAR	szSystemDir[MAX_PATH+1]		= {0};
	TCHAR	szServicesIni[MAX_PATH+12]	= {0}; // 12 = space for "MAPISVC.INF"
	UINT	n							= 0;
	TCHAR	szPropNum[10]				= {0};

	Log(true,_T("HrSetProfileParameters()\n"));

	if (!lpServicesIni) return MAPI_E_INVALID_PARAMETER;

	GetMAPISVCPath(szServicesIni,CCH(szServicesIni));

	if (!szServicesIni[0])
	{
		UINT uiLen = 0;
		uiLen = GetSystemDirectory(szSystemDir, CCH(szSystemDir));
		if (!uiLen)
			return MAPI_E_CALL_FAILED;

		Log(true,_T("Writing to this directory: \"%s\"\n"),szSystemDir);

		hRes = StringCchPrintf(
			szServicesIni,
			CCH(szServicesIni),
			_T("%s\\%s"),
			szSystemDir,
			_T("MAPISVC.INF"));
	}

	Log(true,_T("Writing to this file: \"%s\"\n"),szServicesIni);

	//
	// Loop through and add items to MAPISVC.INF
	//

	n = 0;

	while(lpServicesIni[n].lpszSection != NULL)
	{
		LPTSTR lpszProp = lpServicesIni[n].lpszKey;
		LPTSTR lpszValue = lpServicesIni[n].lpszValue;

		// Switch the property if necessary

		if ((lpszProp == NULL) && (lpServicesIni[n].ulKey != 0))
		{

			hRes = StringCchPrintf(
				szPropNum,
				CCH(szPropNum),
				_T("%lx"),
				lpServicesIni[n].ulKey);

			if (SUCCEEDED(hRes))
				lpszProp = szPropNum;
		}

		//
		// Write the item to MAPISVC.INF
		//

		WritePrivateProfileString(
			lpServicesIni[n].lpszSection,
			lpszProp,
			lpszValue,
			szServicesIni);
		n++;
	}

	// Flush the information - ignore the return code
	WritePrivateProfileString(NULL, NULL, NULL, szServicesIni);

	return hRes;
}
