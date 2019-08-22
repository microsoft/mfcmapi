#pragma once
#include <UI/Dialogs/Dialog.h>
#include <UI/ViewPane/ViewPane.h>
// These panes aren't all used in Editor.h, but included so that including Editor.h elsewhere enables all pane types
#include <UI/ViewPane/TextPane.h>
#include <UI/ViewPane/CheckPane.h>
#include <UI/ViewPane/DropDownPane.h>
#include <UI/ViewPane/ListPane.h>
#include <core/smartview/SmartView.h>

class CParentWnd;

namespace dialog
{
	namespace editor
	{
		// Buttons for CEditor
//#define CEDITOR_BUTTON_OK 0x00000001 // Duplicated from mfcmapi.h - do not modify
#define CEDITOR_BUTTON_ACTION1 0x00000002
#define CEDITOR_BUTTON_ACTION2 0x00000004
//#define CEDITOR_BUTTON_CANCEL 0x00000008 // Duplicated from mfcmapi.h - do not modify
#define CEDITOR_BUTTON_ACTION3 0x00000010

#define NOLIST 0XFFFFFFFF

		template <typename T> viewpane::DoListEditCallback ListEditCallBack(T* editor)
		{
			return [editor](auto a, auto b, auto c) { return editor->DoListEdit(a, b, c); };
		}

		class CEditor : public CMyDialog
		{
		public:
			// Main Edit Constructor
			CEditor(_In_opt_ CWnd* pParentWnd, UINT uidTitle, UINT uidPrompt, ULONG ulButtonFlags);
			CEditor(
				_In_opt_ CWnd* pParentWnd,
				UINT uidTitle,
				UINT uidPrompt,
				ULONG ulButtonFlags,
				UINT uidActionButtonText1,
				UINT uidActionButtonText2,
				UINT uidActionButtonText3);

			_Check_return_ bool DisplayDialog();

			// These functions can be used to set up a data editing dialog
			void SetPromptPostFix(_In_ const std::wstring& szMsg);
			void AddPane(std::shared_ptr<viewpane::ViewPane> lpPane);
			void SetStringA(ULONG id, const std::string& szMsg) const;
			void SetStringW(ULONG id, const std::wstring& szMsg) const;
			void SetStringf(ULONG id, LPCWSTR szMsg, ...) const;
#ifdef CHECKFORMATPARAMS
#undef SetStringf
#define SetStringf(i, fmt, ...) SetStringf(i, fmt, __VA_ARGS__), wprintf(fmt, __VA_ARGS__)
#endif

			void LoadString(ULONG id, UINT uidMsg) const;
			void SetBinary(ULONG id, _In_opt_count_(cb) LPBYTE lpb, size_t cb) const;
			void SetHex(ULONG id, ULONG ulVal) const { SetStringf(id, L"0x%08X", ulVal); } // STRING_OK
			void SetDecimal(ULONG id, ULONG ulVal) const { SetStringf(id, L"%u", ulVal); } // STRING_OK
			// Updates pane using SetStringW
			void SetSize(ULONG id, size_t cb) const
			{
				SetStringf(id, L"0x%08X = %u", static_cast<int>(cb), static_cast<UINT>(cb)); // STRING_OK
			}
			void ClearHighlight(ULONG id) const
			{
				auto pane = std::dynamic_pointer_cast<viewpane::TextPane>(GetPane(id));
				if (pane)
				{
					pane->ClearHighlight();
				}
			}

			void HighlightHex(ULONG id, smartview::block* lpData) const
			{
				auto pane = std::dynamic_pointer_cast<viewpane::TextPane>(GetPane(id));
				if (pane)
				{
					pane->ClearHighlight();
					if (!lpData) return;
					const auto hex = pane->GetStringW();
					pane->AddHighlight(viewpane::Range{
						strings::OffsetToFilteredOffset(hex, lpData->getOffset() * 2),
						strings::OffsetToFilteredOffset(hex, (lpData->getOffset() + lpData->getSize()) * 2)});
				}
			}

			// Get values after we've done the DisplayDialog
			std::shared_ptr<viewpane::ViewPane> GetPane(int id) const;
			std::wstring GetStringW(ULONG id) const;
			_Check_return_ ULONG GetHex(ULONG id) const;
			_Check_return_ ULONG GetDecimal(ULONG id) const;
			_Check_return_ ULONG GetPropTag(ULONG id) const;
			_Check_return_ bool GetCheck(ULONG id) const;
			_Check_return_ int GetDropDown(ULONG id) const;
			_Check_return_ DWORD_PTR GetDropDownValue(ULONG id) const;
			_Check_return_ HRESULT
			GetEntryID(ULONG id, bool bIsBase64, _Out_ size_t* lpcbBin, _Out_ LPENTRYID* lppEID) const;
			// Returns a binary buffer which is represented by the hex string
			_Check_return_ std::vector<BYTE> GetBinary(ULONG id, bool bIsBase64 = false) const
			{
				if (bIsBase64) // entry was BASE64 encoded
				{
					return strings::Base64Decode(GetStringW(id));
				}
				else // Entry was hexized string
				{
					return strings::HexStringToBin(GetStringW(id));
				}
			}

			GUID GetSelectedGUID(ULONG id, bool bByteSwapped) const;

			// AddIn functions
			void SetAddInTitle(const std::wstring& szTitle);
			void SetAddInLabel(ULONG id, const std::wstring& szLabel) const;

			// Use this function to implement list editing
			// return true to indicate the entry was changed, false to indicate it was not
			_Check_return_ virtual bool
			DoListEdit(ULONG /*ulListNum*/, int /*iItem*/, _In_ sortlistdata::sortListData* /*lpData*/)
			{
				return false;
			}

		protected:
			// Functions used by derived classes during init
			void InsertColumn(ULONG id, int nCol, UINT uidText) const;
			void InsertColumn(ULONG id, int nCol, UINT uidText, ULONG ulPropType) const;
			void SetListString(ULONG id, ULONG iListRow, ULONG iListCol, const std::wstring& szListString) const;
			_Check_return_ sortlistdata::sortListData*
			InsertListRow(ULONG id, int iRow, const std::wstring& szText) const;
			void ClearList(ULONG id) const;
			void ResizeList(ULONG id, bool bSort) const;
			void SetListID(ULONG id) { m_listID = id; }

			// Functions used by derived classes during handle change events
			_Check_return_ std::string GetStringA(ULONG id) const;
			_Check_return_ ULONG GetListCount(ULONG id) const;
			_Check_return_ sortlistdata::sortListData* GetListRowData(ULONG id, int iRow) const;
			_Check_return_ bool IsDirty(ULONG id) const;

			// Called to enable/disable buttons based on number of items
			void UpdateButtons() const;
			BOOL OnInitDialog() override;
			void OnOK() override;
			void OnRecalcLayout();

			// protected since derived classes need to call the base implementation
			_Check_return_ virtual ULONG HandleChange(UINT nID);

			void EnableScroll() { m_bEnableScroll = true; }

		private:
			// Overridable functions
			// use these functions to implement custom edit buttons
			virtual void OnEditAction1() {}
			virtual void OnEditAction2() {}
			virtual void OnEditAction3() {}

			// private constructor invoked from public constructors
			void Constructor(
				_In_opt_ CWnd* pParentWnd,
				UINT uidTitle,
				UINT uidPrompt,
				ULONG ulButtonFlags,
				UINT uidActionButtonText1,
				UINT uidActionButtonText2,
				UINT uidActionButtonText3);

			_Check_return_ SIZE ComputeWorkArea(SIZE sScreen);
			void OnGetMinMaxInfo(_Inout_ MINMAXINFO* lpMMI)
			{
				lpMMI->ptMinTrackSize = POINT{m_iMinWidth, m_iMinHeight};
			}
			void OnSetDefaultSize();
			_Check_return_ LRESULT OnNcHitTest(CPoint point);
			void OnSize(UINT nType, int cx, int cy);
			LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;
			void SetMargins() const;

			// List functions and data
			_Check_return_ bool OnEditListEntry(ULONG id) const;
			ULONG m_listID{NOLIST}; // Only supporting one list right now - this is the pane ID for it

			// Our UI controls. Only valid during display.
			bool m_bHasPrompt{};
			CEdit m_Prompt;
			CButton m_OkButton;
			CButton m_ActionButton1;
			CButton m_ActionButton2;
			CButton m_ActionButton3;
			CButton m_CancelButton;
			ULONG m_cButtons{};

			// Variables that get set in the constructor
			ULONG m_bButtonFlags{};

			// Size calculations
			int m_iMargin{};
			int m_iSideMargin{};
			int m_iSmallHeightMargin{};
			int m_iLargeHeightMargin{};
			int m_iButtonWidth{};
			int m_iEditHeight{};
			int m_iTextHeight{};
			int m_iButtonHeight{};
			int m_iMinWidth{};
			int m_iMinHeight{};

			// Title bar, prompt and icon
			UINT m_uidTitle{};
			std::wstring m_szTitle;
			UINT m_uidPrompt{};
			std::wstring m_szPromptPostFix;
			std::wstring m_szAddInTitle;
			HICON m_hIcon{};

			UINT m_uidActionButtonText1{};
			UINT m_uidActionButtonText2{};
			UINT m_uidActionButtonText3{};

			// Panes are held in the order in which they render on screen
			std::vector<std::shared_ptr<viewpane::ViewPane>> m_Panes{};

			bool m_bEnableScroll{};
			CWnd m_ScrollWindow;
			HWND m_hWndVertScroll{};
			bool m_bScrollVisible{};

			int m_iScrollClient{};

			DECLARE_MESSAGE_MAP()
		};
	} // namespace editor
} // namespace dialog