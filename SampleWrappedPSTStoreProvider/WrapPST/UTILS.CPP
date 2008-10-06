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

	Log(true,_T("GetRegistryValue(%s)\n"),lpszValue);

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

// Opens the mail key for the specified MAPI client, such as 'Microsoft Outlook' or 'ExchangeMAPI'
// Pass NULL to open the mail key for the default MAPI client
HKEY GetMailKey(LPTSTR szClient)
{
	Log(true,_T("Enter GetMailKey(%s)\n"),szClient?szClient:_T("Default"));
	HRESULT hRes = S_OK;
	LONG lRes = S_OK;
	HKEY hMailKey = NULL;
	BOOL bClientIsDefault = false;

	// If szClient is NULL, we need to read the name of the default MAPI client
	if (!szClient)
	{
		HKEY hDefaultMailKey = NULL;
		lRes = RegOpenKeyEx(
			HKEY_LOCAL_MACHINE,
			_T("Software\\Clients\\Mail"),
			NULL,
			KEY_READ,
			&hDefaultMailKey);
		if (hDefaultMailKey)
		{
			DWORD dwKeyType = NULL;
			lRes = GetRegistryValue(
				hDefaultMailKey,
				_T(""), // get the default value
				&dwKeyType,
				(LPVOID*) &szClient);
			Log(true,_T("Default MAPI = %s\n"),szClient?szClient:_T("Default"));
			bClientIsDefault = true;
			RegCloseKey(hDefaultMailKey);
		}
	}

	if (szClient)
	{
		TCHAR szMailKey[256];
		hRes = StringCchPrintf(
			szMailKey,
			CCH(szMailKey),
			_T("Software\\Clients\\Mail\\%s"),
			szClient);

		if (SUCCEEDED(hRes))
		{
			lRes = RegOpenKeyEx(
				HKEY_LOCAL_MACHINE,
				szMailKey,
				NULL,
				KEY_READ,
				&hMailKey);
		}
	}
	if (bClientIsDefault) delete[] szClient;

	return hMailKey;
} // GetMailKey

// Gets MSI IDs for the specified MAPI client, such as 'Microsoft Outlook' or 'ExchangeMAPI'
// Pass NULL to get the IDs for the default MAPI client
// Allocates with new, delete with delete[]
void GetMAPIMSIIDs(LPTSTR szClient, LPTSTR* lpszComponentID, LPTSTR* lpszAppLCID, LPTSTR* lpszOfficeLCID)
{
	Log(true,_T("GetMAPIMSIIDs()\n"));
	LONG lRes = S_OK;

	HKEY hKey = GetMailKey(szClient);

	if (hKey)
	{
		DWORD dwKeyType = NULL;

		if (lpszComponentID)
		{
			lRes = GetRegistryValue(
				hKey,
				_T("MSIComponentID"),
				&dwKeyType,
				(LPVOID*) lpszComponentID);
			Log(true,_T("MSIComponentID = %s\n"),*lpszComponentID?*lpszComponentID:_T("<not found>"));
		}

		if (lpszAppLCID)
		{
			lRes = GetRegistryValue(
				hKey,
				_T("MSIApplicationLCID"),
				&dwKeyType,
				(LPVOID*) lpszAppLCID);
			Log(true,_T("MSIApplicationLCID = %s\n"),*lpszAppLCID?*lpszAppLCID:_T("<not found>"));
		}

		if (lpszOfficeLCID)
		{
			lRes = GetRegistryValue(
				hKey,
				_T("MSIOfficeLCID"),
				&dwKeyType,
				(LPVOID*) lpszOfficeLCID);
			Log(true,_T("MSIOfficeLCID = %s\n"),*lpszOfficeLCID?*lpszOfficeLCID:_T("<not found>"));
		}

		RegCloseKey(hKey);
	}
} // GetMAPIComponentID

///////////////////////////////////////////////////////////////////////////////
// Function name   : GetMAPISVCPath
// Description     : This will get the correct path to the MAPISVC.INF file.
// Return type     : void
// Argument        : LPSTR szMAPIDir - Buffer to hold the path to the MAPISVC file.
//                   ULONG cchMAPIDir - size of the buffer
void GetMAPISVCPath(LPTSTR szMAPIDir, ULONG cchMAPIDir)
{
	Log(true,_T("Enter GetMAPISVCPath\n"));
	BOOL bRet = false;
	HINSTANCE hinstStub = NULL;

	szMAPIDir[0] = '\0';	// Terminate String at pos 0 (safer if we fail below)

	// Find some strings:
	LPTSTR szComponentID = NULL;
	LPTSTR szAppLCID = NULL;
	LPTSTR szOfficeLCID = NULL;
	
	GetMAPIMSIIDs("Microsoft Outlook",&szComponentID,&szAppLCID,&szOfficeLCID);

	// Call common code in mapistub.dll
	hinstStub = LoadLibrary(_T("mapistub.dll"));
	if (!hinstStub)
	{
		Log(true,_T("Did not find stub, trying mapi32.dll\n"));
		// Try stub mapi32.dll if mapistub.dll missing
		hinstStub = LoadLibrary(_T("mapi32.dll"));
	}
	
	if (hinstStub)
	{		
		Log(true,_T("Found MAPI\n"));
		LPFGETCOMPONENTPATH pfnFGetComponentPath;

		pfnFGetComponentPath = (LPFGETCOMPONENTPATH)
			GetProcAddress(hinstStub, "FGetComponentPath");

		if (pfnFGetComponentPath)
		{
			Log(true,_T("Got FGetComponentPath\n"));

			if (szAppLCID)
			{
				bRet = pfnFGetComponentPath(
					szComponentID,
					szAppLCID,
					szMAPIDir,
					cchMAPIDir,
					TRUE);
			}
			if ((!bRet || szMAPIDir[0] == _T('\0')) && szOfficeLCID)
			{
				bRet = pfnFGetComponentPath(
					szComponentID,
					szOfficeLCID,
					szMAPIDir,
					cchMAPIDir,
					TRUE);
			}
			if (!bRet || szMAPIDir[0] == _T('\0'))
			{
				bRet = pfnFGetComponentPath(
					szComponentID,
					NULL,
					szMAPIDir,
					cchMAPIDir,
					TRUE);
			}
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
		Log(true,_T("FGetComponentPath failed, loading system directory\n"));
		// Fall back on System32
		UINT uiLen = 0;
		uiLen = GetSystemDirectory(szMAPIDir, cchMAPIDir);
	}

	if (szMAPIDir[0] != _T('\0'))
	{
		Log(true,_T("Using directory: %s\n"),szMAPIDir);
		StringCchPrintf(
			szMAPIDir,
			cchMAPIDir,
			_T("%s\\%s"),
			szMAPIDir,
			_T("MAPISVC.INF"));
	}
	delete[] szComponentID;
	delete[] szOfficeLCID;
	delete[] szAppLCID;
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
