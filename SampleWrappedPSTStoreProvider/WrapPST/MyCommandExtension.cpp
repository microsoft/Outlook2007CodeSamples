#include "stdafx.h"

extern HINSTANCE ghInstDLL; // For the toolbar bitmap

MyCommandExtension::MyCommandExtension(LPUNKNOWN pParentInterface)
{
	Log(true,"new MyCommandExtension\n");
	m_pExchExt = pParentInterface;
	m_cmdid = NULL;
	m_hWnd = NULL;
	m_cRef = 1;

	m_cmdidMyToolbarButton		= 0;
	m_itbbMyButton				= 0;
	m_itbmMyButtonBitmap		= 0;
	m_iMyString					= 0;
}

MyCommandExtension::~MyCommandExtension()
{
	Log(true,"deleting MyCommandExtension\n");
}

STDMETHODIMP MyCommandExtension::QueryInterface
(REFIID                     riid,
 LPVOID *                   ppvObj)
{
    *ppvObj = NULL;
    if (riid == IID_IExchExtCommands)
    {
        *ppvObj = (LPVOID)this;
        // Increase usage count of this object
        AddRef();
        return S_OK;
    }
    if (riid == IID_IUnknown)
    {
     // return MyExchExt as a rule of OLE COM since MyCommandExtension
     // was obtained via MyExchExt::QueryInterface
        *ppvObj = (LPVOID)m_pExchExt;
        m_pExchExt->AddRef();
        return S_OK;
    }
    return E_NOINTERFACE;
};

HRESULT MyCommandExtension::DoCommand(LPEXCHEXTCALLBACK lpeecb,UINT cmdid)
{
	Log(true,"DoCommand %d\n",cmdid);
	// Handle added commands here

	if (cmdid == m_cmdidMyToolbarButton)
	{
		HRESULT hRes = S_OK;

		LPMDB lpCurrentMDB = NULL;
		hRes = lpeecb->GetObject(&lpCurrentMDB,NULL);
		Log(true,"lpeecb->GetObject(&lpCurrentMDB,NULL) returned 0x%08X\n",hRes);

		if (lpCurrentMDB)
		{
			::MessageBox(GetActiveWindow(), "Got MDB - syncing", "Wrapped PST", MB_OK);
			hRes = DoSync(lpCurrentMDB);
			Log(true,"DoSync returned 0x%08X\n",hRes);
		}

		return S_OK;
	}

	// Return Values
	// S_OK -	The extension object recognized the cmdid parameter and carried
	//			out the command. Microsoft Exchange will consider the task handled.
	// S_FALSE - The extension object did not recognize the cmdid parameter.
	//			Microsoft Exchange will continue prompting extension objects to
	//			carry out the command and, if necessary, carry out the command
	//			itself. In some cases, an extension object might perform some
	//			action but still return S_FALSE to give other extensions the
	//			opportunity to handle the command.

	return S_FALSE;
};

HRESULT MyCommandExtension::Help(LPEXCHEXTCALLBACK /*lpeecb*/,UINT /*cmdid*/)
{
	HRESULT hRes = S_FALSE;
	Log(true,"MyCommandExtension::Help\n");
	return hRes;
};

VOID MyCommandExtension::InitMenu(LPEXCHEXTCALLBACK /*lpeecb*/)
{
	Log(true,"MyCommandExtension::InitMenu\n");
	// Do any enabling or disabling here as follows
};

HRESULT MyCommandExtension::InstallCommands(
											LPEXCHEXTCALLBACK /*lpeecb*/,
											HWND /*hwnd*/,
											HMENU /*hmenu*/,
											UINT FAR * lpcmdidBase,
											LPTBENTRY lptbeArray,
											UINT ctbe,
											ULONG /*ulFlags*/)
{
	Log(true,"MyCommandExtension::InstallCommands\n");

	// Add a toolbar item
	// Isolate the toolbar from the menu item, otherwise can't check/uncheck menu item
	m_cmdidMyToolbarButton = (*lpcmdidBase)++;

	// Get the standard toolbar
	HWND hwndToolbar = NULL;
	for (int tbindex = ctbe-1; tbindex > -1; --tbindex)
	{
		if (EETBID_STANDARD == lptbeArray[tbindex].tbid)
		{
			hwndToolbar = lptbeArray[tbindex].hwnd;
			m_itbbMyButton = lptbeArray[tbindex].itbbBase++;
			break;
		}
	}

	// Add button's bitmap to the toolbar's set of buttons
	if (hwndToolbar)
	{
		TBADDBITMAP tbab;
		char * foo = "Foo!\0\0";

		tbab.hInst = (HINSTANCE) ghInstDLL;
		tbab.nID = IDB_MYBITMAP;
		m_itbmMyButtonBitmap = ::SendMessage(hwndToolbar, TB_ADDBITMAP, 1, (LPARAM)&tbab);

		m_iMyString = ::SendMessage(hwndToolbar, TB_ADDSTRING, 0, (LPARAM)foo);
	}

	// Return Values
	// S_OK - Items were added.
	// S_FALSE - Nothing was added.

	return S_OK;
}


HRESULT MyCommandExtension::QueryButtonInfo(
						ULONG /*tbid*/,
						UINT itbb,
						LPTBBUTTON ptbb,
						LPTSTR lpsz,
						UINT cch,
						ULONG /*ulFlags*/)
{
	// Return Values
	// S_OK -	The extension object recognized the itbb parameter and provided
	//			information.
	// S_FALSE - The extension object did not recognize the itbb parameter.
	//			Microsoft Exchange will continue to call extension objects, and if
	//			necessary, provide information itself.
	HRESULT hRes = S_FALSE;
	Log(true,"MyCommandExtension::QueryButtonInfo\n");

	if (m_itbbMyButton == itbb)
	{
		ptbb->iBitmap = m_itbmMyButtonBitmap;
		ptbb->idCommand = m_cmdidMyToolbarButton;
		ptbb->fsState = TBSTATE_ENABLED;
		ptbb->fsStyle = BTNS_BUTTON;
		ptbb->dwData = 0;
		ptbb->iString = -1;
		hRes = StringCchCopy(lpsz,cch, "Wrapped PST Button");

		hRes = S_OK;
	}
	return hRes;
};

HRESULT MyCommandExtension::QueryHelpText(
					  UINT cmdid,
					  ULONG ulFlags,
					  LPTSTR lpsz,
					  UINT cch)
{
	HRESULT hRes = S_FALSE;
	Log(true,"MyCommandExtension::QueryHelpText\n");
	if (!lpsz) return hRes;
	lpsz[0] = NULL;

	if (cmdid == m_cmdidMyToolbarButton)
	{
		if (ulFlags == EECQHT_TOOLTIP)
		{
			hRes = StringCbCopy(lpsz,cch, "Synchronize wrapped PST");
		}
		hRes = S_OK;
	}

	// Return Values
	// S_OK -	The extension object recognized the cmdid parameter and returned
	//			text in the lpsz parameter.
	// S_FALSE - The extension object did not recognize the cmdid parameter.
	//			Microsoft Exchange will continue to call extension objects, and
	//			if necessary, provide text itself.

	return hRes;
};

STDMETHODIMP MyCommandExtension::ResetToolbar(
					 ULONG /*tbid*/,
					 ULONG /*ulFlags*/)
{
	HRESULT hRes = S_OK;
	Log(true,"MyCommandExtension::ResetToolbar\n");
	return hRes;
};