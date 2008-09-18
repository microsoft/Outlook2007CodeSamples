// MyCommandExtension.h: interface for the MyCommandExtension class.
//
//////////////////////////////////////////////////////////////////////
#pragma once

class MyCommandExtension : public IExchExtCommands
{
public:
    MyCommandExtension(LPUNKNOWN pParentInterface);
    ~MyCommandExtension();
    STDMETHODIMP QueryInterface
		(REFIID                     riid,
		LPVOID *                   ppvObj);
    inline STDMETHODIMP_(ULONG) AddRef()
	{
		++m_cRef;
		return m_cRef;
	};
	inline STDMETHODIMP_(ULONG) Release()
	{
		ULONG ulCount = --m_cRef;
		if (!ulCount) { delete this; }
		return ulCount;
	};
	STDMETHODIMP DoCommand(LPEXCHEXTCALLBACK lpeecb,UINT cmdid);
	STDMETHODIMP Help(LPEXCHEXTCALLBACK lpeecb,UINT cmdid);
	STDMETHODIMP_ (VOID) InitMenu(LPEXCHEXTCALLBACK lpeecb);
	STDMETHODIMP InstallCommands(
		LPEXCHEXTCALLBACK lpeecb,
		HWND hwnd,
		HMENU hmenu,
		UINT FAR * lpcmdidBase,
		LPTBENTRY lptbeArray,
		UINT ctbe,
		ULONG ulFlags);
	STDMETHODIMP QueryButtonInfo(
		ULONG tbid,
		UINT itbb,
		LPTBBUTTON ptbb,
		LPTSTR lpsz,
		UINT cch,
		ULONG ulFlags);
	STDMETHODIMP QueryHelpText(
		UINT cmdid,
		ULONG ulFlags,
		LPTSTR lpsz,
		UINT cch);
	STDMETHODIMP ResetToolbar(
		ULONG tbid,
		ULONG ulFlags);
private:
	HWND m_hWnd;
	UINT m_cmdid;
	ULONG m_cRef;
	LPUNKNOWN m_pExchExt;

	// Toolbar items
	UINT m_cmdidMyToolbarButton;

	UINT m_itbmMyButtonBitmap;
	UINT m_itbbMyButton;
	UINT m_iMyString;

};
