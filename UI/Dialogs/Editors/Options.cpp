#include <StdAfx.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/Editors/Options.h>
#include <MAPI/StubUtils.h>

namespace dialog
{
	namespace editor
	{
		class COptions : public CEditor
		{
		public:
			COptions(_In_ CWnd* pWnd);
			virtual ~COptions();
			bool NeedPropRefresh() const;

		private:
			_Check_return_ ULONG HandleChange(UINT nID) override;

			void OnOK() override;

			bool m_bNeedPropRefresh{};
		};

		static std::wstring CLASS = L"COptions";

		COptions::COptions(_In_ CWnd* pWnd)
			: CEditor(pWnd, IDS_SETOPTS, NULL, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
		{
			TRACE_CONSTRUCTOR(CLASS);
			EnableScroll();

			output::DebugPrintEx(DBGGeneric, CLASS, L"COptions(", L"Building option sheet - adding fields\n");

			for (ULONG ulReg = 0; ulReg < NumRegOptionKeys; ulReg++)
			{
				if (registry::regoptCheck == registry::RegKeys[ulReg].ulRegOptType)
				{
					AddPane(viewpane::CheckPane::Create(
						ulReg,
						registry::RegKeys[ulReg].uiOptionsPrompt,
						0 != registry::RegKeys[ulReg].ulCurDWORD,
						false));
				}
				else if (registry::regoptString == registry::RegKeys[ulReg].ulRegOptType)
				{
					AddPane(viewpane::TextPane::CreateSingleLinePane(
						ulReg, registry::RegKeys[ulReg].uiOptionsPrompt, registry::RegKeys[ulReg].szCurSTRING, false));
				}
				else if (registry::regoptStringHex == registry::RegKeys[ulReg].ulRegOptType)
				{
					AddPane(viewpane::TextPane::CreateSingleLinePane(
						ulReg, registry::RegKeys[ulReg].uiOptionsPrompt, false));
					SetHex(ulReg, registry::RegKeys[ulReg].ulCurDWORD);
				}
				else if (registry::regoptStringDec == registry::RegKeys[ulReg].ulRegOptType)
				{
					AddPane(viewpane::TextPane::CreateSingleLinePane(
						ulReg, registry::RegKeys[ulReg].uiOptionsPrompt, false));
					SetDecimal(ulReg, registry::RegKeys[ulReg].ulCurDWORD);
				}
			}
		}

		COptions::~COptions() { TRACE_DESTRUCTOR(CLASS); }

		void COptions::OnOK()
		{
			// need to grab this FIRST
			registry::RegKeys[registry::regkeyDEBUG_FILE_NAME].szCurSTRING =
				GetStringW(registry::regkeyDEBUG_FILE_NAME);

			if (GetHex(registry::regkeyDEBUG_TAG) != registry::RegKeys[registry::regkeyDEBUG_TAG].ulCurDWORD)
			{
				output::SetDebugLevel(GetHex(registry::regkeyDEBUG_TAG));
				output::DebugPrintVersion(DBGVersionBanner);
			}

			output::SetDebugOutputToFile(GetCheck(registry::regkeyDEBUG_TO_FILE));

			// Remaining options require no special handling - loop through them
			for (ULONG ulReg = 0; ulReg < NumRegOptionKeys; ulReg++)
			{
				if (registry::regoptCheck == registry::RegKeys[ulReg].ulRegOptType)
				{
					if (registry::RegKeys[ulReg].bRefresh &&
						registry::RegKeys[ulReg].ulCurDWORD != static_cast<ULONG>(GetCheck(ulReg)))
					{
						m_bNeedPropRefresh = true;
					}

					registry::RegKeys[ulReg].ulCurDWORD = GetCheck(ulReg);
				}
				else if (registry::regoptStringHex == registry::RegKeys[ulReg].ulRegOptType)
				{
					registry::RegKeys[ulReg].ulCurDWORD = GetHex(ulReg);
				}
				else if (registry::regoptStringDec == registry::RegKeys[ulReg].ulRegOptType)
				{
					registry::RegKeys[ulReg].ulCurDWORD = GetDecimal(ulReg);
				}
			}

			// Commit our values to the registry
			registry::WriteToRegistry();

			CEditor::OnOK();
		}

		_Check_return_ ULONG COptions::HandleChange(UINT nID) { return CEditor::HandleChange(nID); }

		bool COptions::NeedPropRefresh() const { return m_bNeedPropRefresh; }

		bool DisplayOptionsDlg(_In_ CWnd* lpParentWnd)
		{
			COptions MyOptions(lpParentWnd);
			MyOptions.DoModal();

			mapistub::ForceOutlookMAPI(0 != registry::RegKeys[registry::regkeyFORCEOUTLOOKMAPI].ulCurDWORD);
			mapistub::ForceSystemMAPI(0 != registry::RegKeys[registry::regkeyFORCESYSTEMMAPI].ulCurDWORD);
			return MyOptions.NeedPropRefresh();
		}
	} // namespace editor
} // namespace dialog
