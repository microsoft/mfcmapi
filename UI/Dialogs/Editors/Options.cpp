#include <StdAfx.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/Editors/Options.h>
#include <MAPI/StubUtils.h>
#include <IO/Registry.h>
#include <IO/MFCOutput.h>
#include <core/utility/strings.h>

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

			for (auto& regKey : registry::RegKeys)
			{
				if (!regKey || !regKey->uiOptionsPrompt) continue;

				if (registry::regoptCheck == regKey->ulRegOptType)
				{
					AddPane(viewpane::CheckPane::Create(
						regKey->uiOptionsPrompt, regKey->uiOptionsPrompt, 0 != regKey->ulCurDWORD, false));
				}
				else if (registry::regoptString == regKey->ulRegOptType)
				{
					AddPane(viewpane::TextPane::CreateSingleLinePane(
						regKey->uiOptionsPrompt, regKey->uiOptionsPrompt, regKey->szCurSTRING, false));
				}
				else if (registry::regoptStringHex == regKey->ulRegOptType)
				{
					AddPane(viewpane::TextPane::CreateSingleLinePane(
						regKey->uiOptionsPrompt, regKey->uiOptionsPrompt, false));
					SetHex(regKey->uiOptionsPrompt, regKey->ulCurDWORD);
				}
				else if (registry::regoptStringDec == regKey->ulRegOptType)
				{
					AddPane(viewpane::TextPane::CreateSingleLinePane(
						regKey->uiOptionsPrompt, regKey->uiOptionsPrompt, false));
					SetDecimal(regKey->uiOptionsPrompt, regKey->ulCurDWORD);
				}
			}
		}

		COptions::~COptions() { TRACE_DESTRUCTOR(CLASS); }

		void COptions::OnOK()
		{
			// need to grab this FIRST
			registry::debugFileName = GetStringW(registry::debugFileName.uiOptionsPrompt);

			if (GetHex(registry::debugTag.uiOptionsPrompt) != registry::debugTag)
			{
				output::SetDebugLevel(GetHex(registry::debugTag.uiOptionsPrompt));
				output::DebugPrintVersion(DBGVersionBanner);
			}

			output::SetDebugOutputToFile(GetCheck(registry::debugToFile.uiOptionsPrompt));

			// Remaining options require no special handling - loop through them
			for (auto& regKey : registry::RegKeys)
			{
				if (!regKey || !regKey->uiOptionsPrompt) continue;

				if (registry::regoptCheck == regKey->ulRegOptType)
				{
					if (regKey->bRefresh && regKey->ulCurDWORD != static_cast<ULONG>(GetCheck(regKey->uiOptionsPrompt)))
					{
						m_bNeedPropRefresh = true;
					}

					regKey->ulCurDWORD = GetCheck(regKey->uiOptionsPrompt);
				}
				else if (registry::regoptStringHex == regKey->ulRegOptType)
				{
					regKey->ulCurDWORD = GetHex(regKey->uiOptionsPrompt);
				}
				else if (registry::regoptStringDec == regKey->ulRegOptType)
				{
					regKey->ulCurDWORD = GetDecimal(regKey->uiOptionsPrompt);
				}
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
