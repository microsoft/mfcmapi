#include <Richole.h>

namespace ui
{
	class CRichEditOleCallback : public IRichEditOleCallback
	{
	public:
		CRichEditOleCallback(HWND hWnd, HWND hWndParent);

		STDMETHODIMP QueryInterface(REFIID riid, LPVOID* ppvObj) override;
		STDMETHODIMP_(ULONG) AddRef() override;
		STDMETHODIMP_(ULONG) Release() override;

		STDMETHODIMP GetNewStorage(THIS_ LPSTORAGE FAR* lplpstg) override;
		STDMETHODIMP GetInPlaceContext(
			THIS_ LPOLEINPLACEFRAME FAR* lplpFrame,
			LPOLEINPLACEUIWINDOW FAR* lplpDoc,
			LPOLEINPLACEFRAMEINFO lpFrameInfo) override;
		STDMETHODIMP ShowContainerUI(THIS_ BOOL fShow) override;
		STDMETHODIMP QueryInsertObject(THIS_ LPCLSID lpclsid, LPSTORAGE lpstg, LONG cp) override;
		STDMETHODIMP DeleteObject(THIS_ LPOLEOBJECT lpoleobj) override;
		STDMETHODIMP QueryAcceptData(
			THIS_ LPDATAOBJECT lpdataobj,
			CLIPFORMAT FAR* lpcfFormat,
			DWORD reco,
			BOOL fReally,
			HGLOBAL hMetaPict) override;
		STDMETHODIMP ContextSensitiveHelp(THIS_ BOOL fEnterMode) override;
		STDMETHODIMP GetClipboardData(THIS_ CHARRANGE FAR* lpchrg, DWORD reco, LPDATAOBJECT FAR* lplpdataobj) override;
		STDMETHODIMP GetDragDropEffect(THIS_ BOOL fDrag, DWORD grfKeyState, LPDWORD pdwEffect) override;
		STDMETHODIMP
		GetContextMenu(THIS_ WORD seltype, LPOLEOBJECT lpoleobj, CHARRANGE FAR* lpchrg, HMENU FAR* lphmenu) override;

	private:
		HWND m_hWnd;
		HWND m_hWndParent;
		LONG m_cRef;
	};
} // namespace ui