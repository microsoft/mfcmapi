#pragma once

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

	_Check_return_ STDMETHODIMP QueryInterface(REFIID riid, _Deref_out_opt_ LPVOID * ppvObj) override;
	STDMETHODIMP_(ULONG) AddRef() override;
	STDMETHODIMP_(ULONG) Release() override;

private:
	LONG m_cRef;
	HWINEVENTHOOK m_hwinEventHook; // Hook to trap header reordering
};