#pragma once

///////////////////////////////////////////////////////////////////////////////
//    forward declarations
//
class MyMessageEvents;
class MyAttachEvents;

class MyExtension : public IExchExt
{

 public:
    MyExtension();
    ~MyExtension();
    STDMETHODIMP QueryInterface
                    (REFIID                     riid,
                     LPVOID *                   ppvObj);
    inline STDMETHODIMP_(ULONG) AddRef
                    () { ++m_cRef; return m_cRef; };
    STDMETHODIMP_(ULONG) Release();
    STDMETHODIMP Install (LPEXCHEXTCALLBACK pmecb,
                        ULONG mecontext, ULONG ulFlags);

 private:
    ULONG m_cRef;
    UINT  m_context;
	MyCommandExtension *m_lpMyCommandExtension;
};
