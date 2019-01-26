#include <StdAfx.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/Editors/DbgView.h>
#include <UI/ParentWnd.h>
#include <core/utility/output.h>

namespace dialog
{
	namespace editor
	{
		class CDbgView : public CEditor
		{
		public:
			CDbgView(_In_ ui::CParentWnd* pParentWnd);
			virtual ~CDbgView();
			void AppendText(const std::wstring& szMsg) const;

		private:
			_Check_return_ ULONG HandleChange(UINT nID) override;
			void OnEditAction1() override;
			void OnEditAction2() override;

			// OnOK override does nothing except *not* call base OnOK
			void OnOK() override {}
			void OnCancel() override;
			bool m_bPaused;
		};

		CDbgView* g_DgbView = nullptr;

		// Displays the debug viewer - only one may exist at a time
		void DisplayDbgView(_In_ ui::CParentWnd* pParentWnd)
		{
			if (!g_DgbView) g_DgbView = new CDbgView(pParentWnd);
		}

		void OutputToDbgView(const std::wstring& szMsg)
		{
			if (!g_DgbView) return;
			g_DgbView->AppendText(szMsg);
		}

		enum __DbgViewFields
		{
			DBGVIEW_TAGS,
			DBGVIEW_PAUSE,
			DBGVIEW_VIEW,
			DBGVIEW_NUMFIELDS, // Must be last
		};

		static std::wstring CLASS = L"CDbgView";

		CDbgView::CDbgView(_In_ ui::CParentWnd* pParentWnd)
			: CEditor(
				  pParentWnd,
				  IDS_DBGVIEW,
				  NULL,
				  CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_ACTION2,
				  IDS_CLEAR,
				  IDS_CLOSE,
				  NULL)
		{
			TRACE_CONSTRUCTOR(CLASS);
			AddPane(viewpane::TextPane::CreateSingleLinePane(DBGVIEW_TAGS, IDS_REGKEY_DEBUG_TAG, false));
			SetHex(DBGVIEW_TAGS, output::GetDebugLevel());
			AddPane(viewpane::CheckPane::Create(DBGVIEW_PAUSE, IDS_PAUSE, false, false));
			AddPane(viewpane::TextPane::CreateMultiLinePane(DBGVIEW_VIEW, NULL, true));
			m_bPaused = false;
			DisplayParentedDialog(pParentWnd, 800);
		}

		CDbgView::~CDbgView()
		{
			TRACE_DESTRUCTOR(CLASS);
			g_DgbView = nullptr;
		}

		void CDbgView::OnCancel()
		{
			ShowWindow(SW_HIDE);
			delete this;
		}

		_Check_return_ ULONG CDbgView::HandleChange(UINT nID)
		{
			const auto paneID = CEditor::HandleChange(nID);

			if (paneID == static_cast<ULONG>(-1)) return static_cast<ULONG>(-1);

			switch (paneID)
			{
			case DBGVIEW_TAGS:
			{
				const auto ulTag = GetHex(DBGVIEW_TAGS);
				output::SetDebugLevel(ulTag);
				return true;
			}
			case DBGVIEW_PAUSE:
			{
				m_bPaused = GetCheck(DBGVIEW_PAUSE);
			}
			break;

			default:
				break;
			}

			return paneID;
		}

		// Clear
		void CDbgView::OnEditAction1()
		{
			auto lpPane = static_cast<viewpane::TextPane*>(GetPane(DBGVIEW_VIEW));
			if (lpPane)
			{
				return lpPane->Clear();
			}
		}

		// Close
		void CDbgView::OnEditAction2() { OnCancel(); }

		void CDbgView::AppendText(const std::wstring& szMsg) const
		{
			if (m_bPaused) return;

			auto lpPane = static_cast<viewpane::TextPane*>(GetPane(DBGVIEW_VIEW));
			if (lpPane)
			{
				lpPane->AppendString(szMsg);
			}
		}
	} // namespace editor
} // namespace dialog