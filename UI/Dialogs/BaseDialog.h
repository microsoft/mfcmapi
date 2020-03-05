#pragma once
#include <UI/Dialogs/Dialog.h>
#include <UI/enums.h>
#include <core/sortlistdata/sortListData.h>
#include <core/addin/mfcmapi.h>
#include <UI/FakeSplitter.h>

namespace controls::sortlistctrl
{
	class CSingleMAPIPropListCtrl;
} // namespace controls::sortlistctrl

namespace cache
{
	class CMapiObjects;
} // namespace cache

namespace mapi
{
	class adviseSink;
} // namespace mapi

// Macros to assist in OnInitMenu
#define CHECK(state) ((state) ? MF_CHECKED : MF_UNCHECKED)
#define DIM(state) ((state) ? MF_ENABLED : MF_GRAYED)
#define DIMMSOK(iNumSelected) ((iNumSelected >= 1) ? MF_ENABLED : MF_GRAYED)
#define DIMMSNOK(iNumSelected) ((iNumSelected == 1) ? MF_ENABLED : MF_GRAYED)

namespace dialog
{
	class CBaseDialog : public CMyDialog
	{
	public:
		CBaseDialog(
			_In_ ui::CParentWnd* pParentWnd,
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			ULONG ulAddInContext);
		~CBaseDialog();

		STDMETHODIMP_(ULONG) AddRef();
		STDMETHODIMP_(ULONG) Release();

		void OnUpdateSingleMAPIPropListCtrl(
			_In_opt_ LPMAPIPROP lpMAPIProp,
			_In_opt_ sortlistdata::sortListData* lpListData) const;
		_Check_return_ bool HandleKeyDown(UINT nChar, bool bShift, bool bCtrl, bool bMenu);

		void UpdateTitleBarText(_In_ const std::wstring& szMsg) const;
		void UpdateTitleBarText() const;
		void UpdateStatusBarText(statusPane nPos, _In_ const std::wstring& szMsg);
		void __cdecl UpdateStatusBarText(statusPane nPos, UINT uidMsg);
		void __cdecl UpdateStatusBarText(statusPane nPos, UINT uidMsg, ULONG ulParam1);
		void __cdecl UpdateStatusBarText(
			statusPane nPos,
			UINT uidMsg,
			const std::wstring& szParam1,
			const std::wstring& szParam2,
			const std::wstring& szParam3);
		void OnOpenEntryID(_In_ const SBinary& bin);
		_Check_return_ ui::CParentWnd* GetParentWnd() const;
		_Check_return_ std::shared_ptr<cache::CMapiObjects> GetMapiObjects() const;

		static void UpdateStatus(HWND hWndHost, const statusPane pane, const std::wstring& status);

	protected:
		// Overrides called by child classes
		virtual void CreateDialogAndMenu(UINT nIDMenuResource, UINT uiClassMenuResource, UINT uidClassMenuTitle);
		virtual void EnableAddInMenus(_In_ HMENU hMenu, ULONG ulMenu, _In_ LPMENUITEM lpAddInMenu, UINT uiEnable);
		_Check_return_ virtual bool HandleMenu(WORD wMenuSelect);
		_Check_return_ virtual bool HandlePaste();
		void OnCancel() override;
		BOOL OnInitDialog() override;
		virtual void OnInitMenu(_In_opt_ CMenu* pMenu);

		ULONG m_ulAddInContext{};
		ULONG m_ulAddInMenuItems{};
		bool m_bIsAB{};
		controls::sortlistctrl::CSingleMAPIPropListCtrl* m_lpPropDisplay{};
		controls::CFakeSplitter m_lpFakeSplitter{};
		std::wstring m_szTitle{};
		std::shared_ptr<cache::CMapiObjects> m_lpMapiObjects{};
		ui::CParentWnd* m_lpParent{};

	private:
		_Check_return_ virtual bool HandleAddInMenu(WORD wMenuSelect);
		static void AddMenu(HMENU hMenuBar, UINT uiResource, UINT uidTitle, UINT uiPos);
		virtual void HandleCopy();
		virtual void OnDeleteSelectedItem();
		virtual void OnEscHit();
		virtual void OnRefreshView();

		// Overrides from base class
		LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;
		void OnActivate(UINT nState, _In_ CWnd* pWndOther, BOOL bMinimized);
		void OnMenuSelect(UINT nItemID, UINT nFlags, HMENU hSysMenu);
		void OnOK() override;
		void OnSize(UINT nType, int cx, int cy);

		// Menu items
		void OnCompareEntryIDs();
		void OnComputeStoreHash();
		void OnDispatchNotifications();
		void OnHelp();
		void OnHexEditor();
		void OnNotificationsOff();
		void OnNotificationsOn();
		void OnOpenMainWindow();
		void OnOptions();
		void OnOutlookVersion();

		void SetStatusWidths();

		// Custom messages
		_Check_return_ LRESULT msgOnUpdateStatusBar(WPARAM wParam, LPARAM lParam);
		_Check_return_ LRESULT msgOnClearSingleMAPIPropList(WPARAM wParam, LPARAM lParam);

		inline std::wstring getStatusMessage(statusPane pane) const noexcept
		{
			return m_StatusMessages[static_cast<int>(pane)];
		}
		inline void setStatusMessage(statusPane pane, std::wstring message) noexcept
		{
			m_StatusMessages[static_cast<int>(pane)] = message;
		}
		inline int getStatusWidth(statusPane pane) const noexcept { return m_StatusWidth[static_cast<int>(pane)]; }
		inline void setStatusWidth(statusPane pane, int width) noexcept
		{
			m_StatusWidth[static_cast<int>(pane)] = width;
		}

		LONG m_cRef{1};
		HICON m_hIcon{};
		std::wstring m_StatusMessages[static_cast<int>(statusPane::numPanes)];
		int m_StatusWidth[static_cast<int>(statusPane::numPanes)]{};
		bool m_bDisplayingMenuText{};
		std::wstring m_szMenuDisplacedText{};
		mapi::adviseSink* m_lpBaseAdviseSink{};
		ULONG_PTR m_ulBaseAdviseConnection{};
		ULONG m_ulBaseAdviseObjectType{};

		DECLARE_MESSAGE_MAP()
	};
} // namespace dialog