#include "stdafx.h"

// instance handle of DLL
extern HINSTANCE ghInstDLL = NULL;

///////////////////////////////////////////////////////////////////////////////
//    FUNCTION: DLLMain()
//
//    Purpose
//    Do initilization processesing
//
//    Return Value
//    TRUE - DLL successfully loads and LoadLibrary will succeed.
//    FALSE - will cause an Exchange error message saying it cannot locate
//            the extension DLL.
//
//    Comments
//    Need to get a copy of the DLL's HINSTANCE.
//
BOOL WINAPI DllMain(
    HINSTANCE  hinstDLL,
    DWORD  fdwReason,
    LPVOID  /*lpvReserved*/)
{
	Log(true,"wrppst32.dll in DLLMain, fdwReason = 0x%08X\n",fdwReason);
	if (DLL_PROCESS_ATTACH == fdwReason)
	{
		ghInstDLL = hinstDLL;
	}
	else if (DLL_PROCESS_DETACH == fdwReason)
	{
		DeInitLogging();
	}
	return TRUE;
}

///////////////////////////////////////////////////////////////////////////////
//    FUNCTION: ExchEntryPoint
//
//    Parameters - none
//
//    Purpose
//    The entry point called by Exchange.
//
//    Return Value
//    Pointer to Exchange Extension (IExchExt) interface
//
//    Comments
//    Exchange Client calls this for each context entry.
//

LPEXCHEXT CALLBACK ExchEntryPoint(void)
{
	Log(true,"Calling ExchEntryPoint for wrppst32.dll\n");
	return new MyExtension;
}


//////////////////////////////////////////////////////////////
// LogREFIID
// logs an REFIID
static WCHAR oleszString[512];
void LogREFIID(REFIID riid)
{
	::StringFromGUID2(riid,oleszString,255);
	Log(false,"%S",oleszString);
}
