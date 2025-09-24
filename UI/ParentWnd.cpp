#include <StdAfx.h>
#include <UI/ParentWnd.h>
#include <core/utility/import.h>
#include <UI/MyWinApp.h>
#include <UI/UIFunctions.h>
#include <mapistub/library/stubutils.h>
#include <core/addin/addin.h>
#include <core/utility/registry.h>
#include <core/utility/output.h>

extern ui::CMyWinApp theApp;

namespace ui
{
	static std::wstring CLASS = L"CParentWnd";

	static CParentWnd* s_parentWnd{};
	_Check_return_ CParentWnd* GetParentWnd()
	{
		if (!s_parentWnd)
		{
			s_parentWnd = new CParentWnd();
		}

		return s_parentWnd;
	}

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
	void CALLBACK MyWinEventProc(
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
			// We don't need to wait on the results - just post the message
			WC_B_S(PostMessage(hwnd, WM_MFCMAPI_SAVECOLUMNORDERHEADER, WPARAM(NULL), LPARAM(NULL)));
		}
	}

	CParentWnd::CParentWnd()
	{
		// OutputDebugStringOutput only at first
		// Get any settings from the registry
		registry::ReadFromRegistry();
		// After this call we may output to the debug file
		output::OpenDebugFile();
		output::outputVersion(output::dbgLevel::VersionBanner, nullptr);
		// Force the system riched20 so we don't load office's version.
		static_cast<void>(import::LoadFromSystemDir(L"riched20.dll")); // STRING_OK
		// Second part is to load rundll32.exe
		// Don't plan on unloading this, so don't care about the return value
		static_cast<void>(import::LoadFromSystemDir(L"rundll32.exe")); // STRING_OK

		// Load DLLS and get functions from them
		import::ImportProcs();

		// Initialize objects for theming
		InitializeGDI();

		m_cRef = 1;

		m_hwinEventHook = SetWinEventHook(
			EVENT_OBJECT_REORDER, EVENT_OBJECT_REORDER, nullptr, &MyWinEventProc, GetCurrentProcessId(), NULL, NULL);

		registry::PushOptionsToStub();

		addin::LoadAddIns();

		// Notice we never create a window here!
		TRACE_CONSTRUCTOR(CLASS);
	}

	CParentWnd::~CParentWnd()
	{
		// Since we didn't create a window, we can't call DestroyWindow to let MFC know we're done.
		// We call AfxPostQuitMessage instead
		TRACE_DESTRUCTOR(CLASS);
		s_parentWnd = nullptr;
		UninitializeGDI();
		addin::UnloadAddIns();
		if (m_hwinEventHook) UnhookWinEvent(m_hwinEventHook);
		registry::WriteToRegistry();
		output::CloseDebugFile();
		// Since we're killing what m_pMainWnd points to here, we need to clear it
		// Else MFC will try to route messages to it
		theApp.m_pMainWnd = nullptr;
		AfxPostQuitMessage(0);
	}

	_Check_return_ STDMETHODIMP CParentWnd::QueryInterface(REFIID riid, _Deref_out_opt_ LPVOID* ppvObj)
	{
		*ppvObj = nullptr;
		if (riid == IID_IUnknown)
		{
			*ppvObj = this;
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}

	STDMETHODIMP_(ULONG) CParentWnd::AddRef()
	{
		const auto lCount = InterlockedIncrement(&m_cRef);
		TRACE_ADDREF(CLASS, lCount);
		return lCount;
	}

	STDMETHODIMP_(ULONG) CParentWnd::Release()
	{
		const auto lCount = InterlockedDecrement(&m_cRef);
		TRACE_RELEASE(CLASS, lCount);
		if (!lCount) delete this;
		return lCount;
	}
} // namespace ui