#include <StdAfx.h>
#include <UI/RichEditOleCallback.h>
#include <UI/UIFunctions.h>
#include <core/utility/output.h>

namespace ui
{
	static std::wstring CLASS = L"CRichEditOleCallback";

	CRichEditOleCallback::CRichEditOleCallback(HWND hWnd, HWND hWndParent)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_hWnd = hWnd;
		m_hWndParent = hWndParent;
		m_cRef = 1;
	}

	STDMETHODIMP CRichEditOleCallback::QueryInterface(REFIID riid, LPVOID* ppvObj)
	{
		if (!ppvObj) return MAPI_E_INVALID_PARAMETER;
		*ppvObj = nullptr;
		if (riid == IID_IRichEditOleCallback || riid == IID_IUnknown)
		{
			*ppvObj = this;
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}

	STDMETHODIMP_(ULONG) CRichEditOleCallback::AddRef()
	{
		const auto lCount = InterlockedIncrement(&m_cRef);
		TRACE_ADDREF(CLASS, lCount);
		return lCount;
	}

	STDMETHODIMP_(ULONG) CRichEditOleCallback::Release()
	{
		const auto lCount = InterlockedDecrement(&m_cRef);
		TRACE_RELEASE(CLASS, lCount);
		if (!lCount) delete this;
		return lCount;
	}

	STDMETHODIMP CRichEditOleCallback::GetNewStorage(LPSTORAGE FAR* /*lplpstg*/) { return E_NOTIMPL; }

	STDMETHODIMP CRichEditOleCallback::GetInPlaceContext(
		LPOLEINPLACEFRAME FAR* /*lplpFrame*/,
		LPOLEINPLACEUIWINDOW FAR* /*lplpDoc*/,
		LPOLEINPLACEFRAMEINFO /*lpFrameInfo*/)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP CRichEditOleCallback::ShowContainerUI(BOOL /*fShow*/) { return E_NOTIMPL; }

	STDMETHODIMP CRichEditOleCallback::QueryInsertObject(LPCLSID /*lpclsid*/, LPSTORAGE /*lpstg*/, LONG /*cp*/)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP CRichEditOleCallback::DeleteObject(LPOLEOBJECT /*lpoleobj*/) { return E_NOTIMPL; }

	STDMETHODIMP CRichEditOleCallback::QueryAcceptData(
		LPDATAOBJECT /*lpdataobj*/,
		CLIPFORMAT FAR* lpcfFormat,
		DWORD /*reco*/,
		BOOL /*fReally*/,
		HGLOBAL /*hMetaPict*/)
	{
		if (lpcfFormat)
		{
			// Ensure that the only paste we ever get is CF_TEXT
			*lpcfFormat = CF_TEXT;
		}

		return S_OK;
	}

	STDMETHODIMP CRichEditOleCallback::ContextSensitiveHelp(BOOL /*fEnterMode*/) { return E_NOTIMPL; }

	STDMETHODIMP
	CRichEditOleCallback::GetClipboardData(CHARRANGE FAR* /*lpchrg*/, DWORD /*reco*/, LPDATAOBJECT FAR* /*lplpdataobj*/)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP CRichEditOleCallback::GetDragDropEffect(BOOL /*fDrag*/, DWORD /*grfKeyState*/, LPDWORD /*pdwEffect*/)
	{
		return E_NOTIMPL;
	}

	// We're supposed to return a menu, but then we can't control where the command messages are sent.
	// So, instead, we display and handle the menu ourselves, then return no menu.
	STDMETHODIMP CRichEditOleCallback::GetContextMenu(
		WORD /*seltype*/,
		LPOLEOBJECT /*lpoleobj*/,
		CHARRANGE FAR* /*lpchrg*/,
		HMENU FAR* lphmenu)
	{
		if (!lphmenu) return E_INVALIDARG;
		lphmenu = nullptr;
		const auto hContext = LoadMenu(nullptr, MAKEINTRESOURCE(IDR_MENU_RICHEDIT_POPUP));
		if (hContext)
		{
			const auto hPopup = GetSubMenu(hContext, 0);
			if (hPopup)
			{
				ConvertMenuOwnerDraw(hPopup, false);
				POINT pos = {0};
				CURSORINFO ci = {0};
				ci.cbSize = sizeof(CURSORINFO);
				GetCursorInfo(&ci);
				pos.x = ci.ptScreenPos.x;
				pos.y = ci.ptScreenPos.y;

				// If the mouse is outside the window, then this was likely a keyboard triggered edit.
				// So display the context menu where the cursor is.
				if (m_hWnd != WindowFromPoint(pos))
				{
					DWORD dwPos = 0;
					(void) ::SendMessage(
						m_hWnd, EM_GETSEL, reinterpret_cast<WPARAM>(&dwPos), static_cast<LPARAM>(NULL));
					(void) ::SendMessage(
						m_hWnd, EM_POSFROMCHAR, reinterpret_cast<WPARAM>(&pos), static_cast<LPARAM>(dwPos));
					ClientToScreen(m_hWnd, &pos);
				}

				const DWORD dwCommand = TrackPopupMenu(
					hPopup,
					TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_NONOTIFY | TPM_RETURNCMD,
					pos.x,
					pos.y,
					NULL,
					m_hWndParent,
					nullptr);
				DeleteMenuEntries(hPopup);
				(void) ::SendMessage(
					m_hWnd, dwCommand, static_cast<WPARAM>(0), static_cast<LPARAM>(EM_SETSEL == dwCommand) ? -1 : 0);
			}

			DestroyMenu(hContext);
		}

		return S_OK;
	}
} // namespace ui