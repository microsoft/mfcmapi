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

			ULONG ulReg = 0;
			for (auto& regKey : registry::RegKeys)
			{
				if (regKey && regKey->uiOptionsPrompt)
				{
					if (registry::regoptCheck == regKey->ulRegOptType)
					{
						AddPane(viewpane::CheckPane::Create(
							ulReg, regKey->uiOptionsPrompt, 0 != regKey->ulCurDWORD, false));
					}
					else if (registry::regoptString == regKey->ulRegOptType)
					{
						AddPane(viewpane::TextPane::CreateSingleLinePane(
							ulReg, regKey->uiOptionsPrompt, regKey->szCurSTRING, false));
					}
					else if (registry::regoptStringHex == regKey->ulRegOptType)
					{
						AddPane(viewpane::TextPane::CreateSingleLinePane(ulReg, regKey->uiOptionsPrompt, false));
						SetHex(ulReg, regKey->ulCurDWORD);
					}
					else if (registry::regoptStringDec == regKey->ulRegOptType)
					{
						AddPane(viewpane::TextPane::CreateSingleLinePane(ulReg, regKey->uiOptionsPrompt, false));
						SetDecimal(ulReg, regKey->ulCurDWORD);
					}
				}

				ulReg++;
			}
		}

		COptions::~COptions() { TRACE_DESTRUCTOR(CLASS); }

		void COptions::OnOK()
		{
			// need to grab this FIRST
			registry::debugFileName = GetStringW(registry::regkeyDEBUG_FILE_NAME);

			if (GetHex(registry::regkeyDEBUG_TAG) != registry::debugTag)
			{
				output::SetDebugLevel(GetHex(registry::regkeyDEBUG_TAG));
				output::DebugPrintVersion(DBGVersionBanner);
			}

			output::SetDebugOutputToFile(GetCheck(registry::regkeyDEBUG_TO_FILE));

			// Remaining options require no special handling - loop through them
			ULONG ulReg = 0;
			for (auto& regKey : registry::RegKeys)
			{
				if (regKey && regKey->uiOptionsPrompt)
				{
					if (registry::regoptCheck == regKey->ulRegOptType)
					{
						if (regKey->bRefresh && regKey->ulCurDWORD != static_cast<ULONG>(GetCheck(ulReg)))
						{
							m_bNeedPropRefresh = true;
						}

						regKey->ulCurDWORD = GetCheck(ulReg);
					}
					else if (registry::regoptStringHex == regKey->ulRegOptType)
					{
						regKey->ulCurDWORD = GetHex(ulReg);
					}
					else if (registry::regoptStringDec == regKey->ulRegOptType)
					{
						regKey->ulCurDWORD = GetDecimal(ulReg);
					}
				}

				ulReg++;
			}

			// Commit our values to the registry
			registry::WriteToRegistry();

			CEditor::OnOK();
		}

		_Check_return_ ULONG COptions::HandleChange(const UINT nID) { return CEditor::HandleChange(nID); }

		bool COptions::NeedPropRefresh() const { return m_bNeedPropRefresh; }

		bool DisplayOptionsDlg(_In_ CWnd* lpParentWnd)
		{
			COptions MyOptions(lpParentWnd);
			MyOptions.DoModal();

			mapistub::ForceOutlookMAPI(registry::forceOutlookMAPI);
			mapistub::ForceSystemMAPI(registry::forceSystemMAPI);
			return MyOptions.NeedPropRefresh();
		}
	} // namespace editor
} // namespace dialog
