// ParentWnd.cpp : implementation file
//

#include "stdafx.h"
#include "ParentWnd.h"
#include "ImportProcs.h"
#include "MyWinApp.h"
#include "NamedPropCache.h"
#include "UIFunctions.h"

extern CMyWinApp theApp;

static TCHAR* CLASS = _T("CParentWnd");

// This appears to be the only way to catch a column drag event
// Since we end up catching for EVERY event, we have to be clever to ensure only
// the CSingleMAPIPropListCtrl that generated it catches it.
// First, we throw WM_MFCMAPI_SAVECOLUMNORDERHEADER to whoever generated the event
// Most controls will ignore this - only CSortHeader knows this message
// That ditches most false hits
// Then CSortHeader throws WM_MFCMAPI_SAVECOLUMNORDERLIST to its parent
// Only CSingleMAPIPropListCtrl handles that, so we'll only get hits when the header for
// a particular CSingleMAPIPropListCtrl gets changed.
// We can't throw to the CSingleMAPIPropListCtrl directly since we don't have a handle here for it
VOID CALLBACK MyWinEventProc(
							 HWINEVENTHOOK /*hWinEventHook*/,
							 DWORD event,
							 _In_ HWND hwnd,
							 LONG /*idObject*/,
							 LONG /*idChild*/,
							 DWORD /*dwEventThread*/,
							 DWORD /*dwmsEventTime*/)
{
	if (EVENT_OBJECT_REORDER == event)
	{
		HRESULT hRes = S_OK;
		// We don't need to wait on the results - just post the message
		WC_B(::PostMessage(
			hwnd,
			WM_MFCMAPI_SAVECOLUMNORDERHEADER,
			(WPARAM) NULL,
			(LPARAM) NULL));
	}
} // MyWinEventProc

/////////////////////////////////////////////////////////////////////////////
// CParentWnd

CParentWnd::CParentWnd()
{
	// OutputDebugStringOutput only at first
	// Get any settings from the registry
	SetDefaults();
	ReadFromRegistry();
	// After this call we may output to the debug file
	OpenDebugFile();
	DebugPrintVersion(DBGVersionBanner);
	// Force the system riched20 so we don't load office's version.
	(void) LoadFromSystemDir(_T("riched20.dll")); // STRING_OK
	// Second part is to load rundll32.exe
	// Don't plan on unloading this, so don't care about the return value
	(void) LoadFromSystemDir(_T("rundll32.exe")); // STRING_OK

	// Load DLLS and get functions from them
	ImportProcs();

	// Initialize objects for theming
	InitializeGDI();

	m_cRef = 1;

	m_hwinEventHook = SetWinEventHook(
		EVENT_OBJECT_REORDER,
		EVENT_OBJECT_REORDER,
		NULL,
		&MyWinEventProc,
		GetCurrentProcessId(),
		NULL,
		NULL);

	ForceOutlookMAPI(0 != RegKeys[regkeyFORCEOUTLOOKMAPI].ulCurDWORD);
	ForceSystemMAPI(0 != RegKeys[regkeyFORCESYSTEMMAPI].ulCurDWORD);

	LoadAddIns();

	// Notice we never create a window here!
	TRACE_CONSTRUCTOR(CLASS);
} // CParentWnd::CParentWnd

CParentWnd::~CParentWnd()
{
	// Since we didn't create a window, we can't call DestroyWindow to let MFC know we're done.
	// We call AfxPostQuitMessage instead
	TRACE_DESTRUCTOR(CLASS);
	UninitializeGDI();
	UnloadAddIns();
	if (m_hwinEventHook) UnhookWinEvent(m_hwinEventHook);
	UninitializeNamedPropCache();
	WriteToRegistry();
	CloseDebugFile();
	// Since we're killing what m_pMainWnd points to here, we need to clear it
	// Else MFC will try to route messages to it
	theApp.m_pMainWnd = NULL;
	AfxPostQuitMessage(0);
} // CParentWnd::~CParentWnd

_Check_return_ STDMETHODIMP CParentWnd::QueryInterface(REFIID riid,
													   _Deref_out_opt_ LPVOID * ppvObj)
{
	*ppvObj = 0;
	if (riid == IID_IUnknown)
	{
		*ppvObj = (LPVOID)this;
		AddRef();
		return S_OK;
	}
	return E_NOINTERFACE;
} // CParentWnd::QueryInterface

STDMETHODIMP_(ULONG) CParentWnd::AddRef()
{
	LONG lCount = InterlockedIncrement(&m_cRef);
	TRACE_ADDREF(CLASS,lCount);
	return lCount;
} // CParentWnd::AddRef

STDMETHODIMP_(ULONG) CParentWnd::Release()
{
	LONG lCount = InterlockedDecrement(&m_cRef);
	TRACE_RELEASE(CLASS,lCount);
	if (!lCount) delete this;
	return lCount;
} // CParentWnd::Release