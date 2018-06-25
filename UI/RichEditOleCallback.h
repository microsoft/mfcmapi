#include <Richole.h>

namespace ui
{
	class CRichEditOleCallback : public IRichEditOleCallback
	{
	public:
		CRichEditOleCallback(HWND hWnd, HWND hWndParent);

		STDMETHODIMP QueryInterface(REFIID riid, LPVOID* ppvObj);
		STDMETHODIMP_(ULONG) AddRef();
		STDMETHODIMP_(ULONG) Release();

		STDMETHOD(GetNewStorage)(THIS_ LPSTORAGE FAR* lplpstg);
		STDMETHOD(GetInPlaceContext)
		(THIS_ LPOLEINPLACEFRAME FAR* lplpFrame, LPOLEINPLACEUIWINDOW FAR* lplpDoc, LPOLEINPLACEFRAMEINFO lpFrameInfo);
		STDMETHOD(ShowContainerUI)(THIS_ BOOL fShow);
		STDMETHOD(QueryInsertObject)(THIS_ LPCLSID lpclsid, LPSTORAGE lpstg, LONG cp);
		STDMETHOD(DeleteObject)(THIS_ LPOLEOBJECT lpoleobj);
		STDMETHOD(QueryAcceptData)
		(THIS_ LPDATAOBJECT lpdataobj, CLIPFORMAT FAR* lpcfFormat, DWORD reco, BOOL fReally, HGLOBAL hMetaPict);
		STDMETHOD(ContextSensitiveHelp)(THIS_ BOOL fEnterMode);
		STDMETHOD(GetClipboardData)(THIS_ CHARRANGE FAR* lpchrg, DWORD reco, LPDATAOBJECT FAR* lplpdataobj);
		STDMETHOD(GetDragDropEffect)(THIS_ BOOL fDrag, DWORD grfKeyState, LPDWORD pdwEffect);
		STDMETHOD(GetContextMenu)(THIS_ WORD seltype, LPOLEOBJECT lpoleobj, CHARRANGE FAR* lpchrg, HMENU FAR* lphmenu);

	private:
		HWND m_hWnd;
		HWND m_hWndParent;
		LONG m_cRef;
	};
}