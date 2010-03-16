#pragma once

// ParentWnd.h : header file
// This object exists to force MFC to keep our main thread alive until all of
// our objects have cleaned themselves up. All objects Addref this object on
// instantiation and Release it when they delete themselves. When there are no
// more outstanding objects, the destructor is called (from Release), where we
// call AfxPostQuitMessage to signal MFC we are done.

// We never call Create on this object so that all Dialogs created with it are
// ownerless. This means they show up in Task Manager, Alt + Tab, and on the taskbar.

class CParentWnd : public CWnd, IUnknown
{
public:
	CParentWnd();
	virtual ~CParentWnd();

	STDMETHODIMP         QueryInterface(REFIID riid, LPVOID * ppvObj);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

private:
	LONG m_cRef;
	HWINEVENTHOOK m_hwinEventHook; // Hook to trap header reordering
};