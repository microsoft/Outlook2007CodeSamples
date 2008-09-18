#include "stdafx.h"

///////////////////////////////////////////////////////////////////////////////
//    MyExtension::MyExtension()
//
//    Parameters - none
//
//    Purpose
//    Comstructor.  Called during instantiation of MyExtension object.
//
//

MyExtension::MyExtension()
{
	Log(true,"New MyExtension\n");
	m_cRef = 1;
	m_context = 0;
	m_lpMyCommandExtension = new MyCommandExtension(this);
};

MyExtension::~MyExtension()
{
	m_lpMyCommandExtension->Release();
	Log(true,"Deleting MyExtension\n");
}


///////////////////////////////////////////////////////////////////////////////
//  IExchExt virtual member functions implementation
//

///////////////////////////////////////////////////////////////////////////////
//    MyExtension::Release()
//
//    Parameters - none
//
//    Purpose
//    Frees memory when interface is not referenced any more
//
//    Return value
//    reference count of interface
//

STDMETHODIMP_(ULONG) MyExtension::Release()
{
 ULONG ulCount = --m_cRef;

 if (!ulCount)
 {
  delete this;
 }

return ulCount;

}

///////////////////////////////////////////////////////////////////////////////
//    MyExtension::QueryInterface()
//
//    Parameters
//    riid   -- Interface ID.
//    ppvObj -- address of interface object pointer.
//
//    Purpose
//    Returns a pointer to an interface object that is requested by ID.
//
//    Comments
//    The interfaces are requested everytime a new context is entered.  The
//    IID_IExchExt* interfaces are ignored if not specified in the Exchange
//    extensions registry.
//
//    If an interface pointer is returned for more than one context, that
//    interface is used by the client for each of those contexts.  Check the
//    current context to verify if it is appropriate to pass back an interface
//    pointer.
//
//    Return Value - none
//

STDMETHODIMP MyExtension::QueryInterface(REFIID riid, LPVOID FAR * ppvObj)
{
    HRESULT hResult = S_OK;

    *ppvObj = NULL;

	Log(true,"MyExtension: Context %d and ",m_context);

	LogREFIID(riid);
	Log(false," == ");
	if (riid == IID_IUnknown) Log(false,"IID_IUnknown");
	else if (riid == IID_IExchExt) Log(false,"IID_IExchExt");
	else if (riid == IID_IExchExtAdvancedCriteria) Log(false,"IID_IExchExtAdvancedCriteria");
	else if (riid == IID_IExchExtAttachedFileEvents) Log(false,"IID_IExchExtAttachedFileEvents");
	else if (riid == IID_IExchExtCommands) Log(false,"IID_IExchExtCommands");
	else if (riid == IID_IExchExtMessageEvents) Log(false,"IID_IExchExtMessageEvents");
	else if (riid == IID_IExchExtPropertySheets) Log(false,"IID_IExchExtPropertySheets");
	else if (riid == IID_IExchExtSessionEvents) Log(false,"IID_IExchExtSessionEvents");
	else if (riid == IID_IExchExtUserEvents) Log(false,"IID_IExchExtUserEvents");
	else
	{
		Log(false,"Unknown IID");
	}
	Log(false," requested\n");

    if (( IID_IUnknown == riid) || ( IID_IExchExt == riid) )
    {
		Log(true,"IID_IExchExt granted\n");
        *ppvObj = (LPUNKNOWN)this;
    }
	else if (IID_IExchExtCommands == riid)
	{
		Log(true,"IID_IExchExtCommands granted\n");
		*ppvObj = (LPUNKNOWN)m_lpMyCommandExtension;
	}
    else
        hResult = E_NOINTERFACE;

    if (NULL != *ppvObj)
        ((LPUNKNOWN)*ppvObj)->AddRef();

    return hResult;
}



///////////////////////////////////////////////////////////////////////////////
//    MyExtension::Install()
//
//    Parameters
//    peecb     -- pointer to Exchange Extension callback function
//    eecontext -- context code at time of being called.
//
//    Purpose
//    Called once for each new context that is entered.
//
//    Return Value
//    S_OK - the installation succeeded for the context
//    S_FALSE - deny the installation fo the extension for the context
//

STDMETHODIMP MyExtension::Install(LPEXCHEXTCALLBACK peecb, ULONG eecontext, ULONG /*ulFlags*/)
{
    ULONG ulBuildVersion;

    m_context = eecontext;

    // make sure this is the right version
    peecb->GetVersion(&ulBuildVersion, EECBGV_GETBUILDVERSION);

//	Log(true,"Major build: %x\n",ulBuildVersion);

	Log(true,"Installed in eecontext: %d, ",eecontext);

//    if (EECBGV_BUILDVERSION_MAJOR != (ulBuildVersion &
//		EECBGV_BUILDVERSION_MAJOR_MASK))
//        return S_FALSE;


    switch (eecontext)
    {
		// The duration of the Microsoft Exchange task, from program start to program exit. This may span several logons. This context does not correspond to a Microsoft Exchange window. In this context, an extension can abort Exchange startup by returning an error.
	case(EECONTEXT_TASK ):
		Log(false,"EECONTEXT_TASK");break;
		// The duration of a Microsoft Exchange session from logon to logoff. Multiple logons can occur during a single execution of Microsoft Exchange, but they might not overlap. This context does not correspond to a Microsoft Exchange window.
	case(EECONTEXT_SESSION ):
		Log(false,"EECONTEXT_SESSION");break;
		// The main Viewer window that displays the folder hierarchy in the left pane and folder contents in the right pane.
	case(EECONTEXT_VIEWER ):
		Log(false,"EECONTEXT_VIEWER");break;
		// The Remote Mail window that is displayed when the user chooses the Remote Mail command.
	case(EECONTEXT_REMOTEVIEWER ) :
		Log(false,"EECONTEXT_REMOTEVIEWER");break;
		// The Find window that is displayed when the user chooses the Find command.
	case(EECONTEXT_SEARCHVIEWER ):
		Log(false,"EECONTEXT_SEARCHVIEWER");break;
		// The Address Book window that is displayed when the user chooses the Address Book command. This context only applies to a modeless address book.
	case(EECONTEXT_ADDRBOOK ):
		Log(false,"EECONTEXT_ADDRBOOK");break;
		// The standard Compose, New Message window in which messages of class IPM.Note are composed.
	case(EECONTEXT_SENDNOTEMESSAGE ):
		Log(false,"EECONTEXT_SENDNOTEMESSAGE");break;
		// The standard read note window in which messages of class IPM.Note are read after they are received.
	case(EECONTEXT_READNOTEMESSAGE ):
		Log(false,"EECONTEXT_READNOTEMESSAGE");break;
		// The read report message window in which report messages (Read, Delivery, Non-Read, Non-Delivery) are read after they are received.
	case(EECONTEXT_READREPORTMESSAGE ):
		Log(false,"EECONTEXT_READREPORTMESSAGE");break;
		// The send resend message window that is displayed when the user chooses the Send Again command on the non-delivery report.
	case(EECONTEXT_SENDRESENDMESSAGE ):
		Log(false,"EECONTEXT_SENDRESENDMESSAGE");break;
		// The standard posting window in which existing posting messages are composed.
	case(EECONTEXT_SENDPOSTMESSAGE ):
		Log(false,"EECONTEXT_SENDPOSTMESSAGE");break;
		// The standard posting window in which existing posting messages are read.
	case(EECONTEXT_READPOSTMESSAGE ):
		Log(false,"EECONTEXT_READPOSTMESSAGE");break;
		// A property sheet window.
	case(EECONTEXT_PROPERTYSHEETS ):
		Log(false,"EECONTEXT_PROPERTYSHEETS");break;
		// The dialog box in which the user specifies advanced search criteria.
	case(EECONTEXT_ADVANCEDCRITERIA ):
		Log(false,"EECONTEXT_ADVANCEDCRITERIA");break;

	default:
		Log(false,"Unknown context");
		break;
    }
	Log(false,"\n");

    return S_OK;

}

