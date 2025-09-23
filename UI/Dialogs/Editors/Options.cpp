#include <StdAfx.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/Editors/Options.h>
#include <mapistub/library/stubutils.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/addin/mfcmapi.h>

namespace dialog::editor
{
	class COptions : public CEditor
	{
	public:
		COptions(_In_ CWnd* pWnd);
		~COptions();
		bool NeedPropRefresh() const noexcept;

	private:
		_Check_return_ ULONG HandleChange(UINT nID) override;

		void OnOK() override;

		bool m_bNeedPropRefresh{};
	};

	static std::wstring CLASS = L"COptions";

	COptions::COptions(_In_ CWnd* pWnd) : CEditor(pWnd, IDS_SETOPTS, NULL, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
	{
		TRACE_CONSTRUCTOR(CLASS);
		EnableScroll();

		output::DebugPrintEx(
			output::dbgLevel::Generic, CLASS, L"COptions(", L"Building option sheet - adding fields\n");

		for (auto& regKey : registry::RegKeys)
		{
			if (!regKey || !regKey->uiOptionsPrompt) continue;

			if (registry::regOptionType::check == regKey->ulRegOptType)
			{
				AddPane(viewpane::CheckPane::Create(
					regKey->uiOptionsPrompt, regKey->uiOptionsPrompt, 0 != regKey->ulCurDWORD, false));
			}
			else if (registry::regOptionType::string == regKey->ulRegOptType)
			{
				AddPane(viewpane::TextPane::CreateSingleLinePane(
					regKey->uiOptionsPrompt, regKey->uiOptionsPrompt, regKey->szCurSTRING, false));
			}
			else if (registry::regOptionType::stringHex == regKey->ulRegOptType)
			{
				AddPane(
					viewpane::TextPane::CreateSingleLinePane(regKey->uiOptionsPrompt, regKey->uiOptionsPrompt, false));
				SetHex(regKey->uiOptionsPrompt, regKey->ulCurDWORD);
			}
			else if (registry::regOptionType::stringDec == regKey->ulRegOptType)
			{
				AddPane(
					viewpane::TextPane::CreateSingleLinePane(regKey->uiOptionsPrompt, regKey->uiOptionsPrompt, false));
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
			registry::debugTag = GetHex(registry::debugTag.uiOptionsPrompt);
			output::outputVersion(output::dbgLevel::VersionBanner, nullptr);
		}

		output::SetDebugOutputToFile(GetCheck(registry::debugToFile.uiOptionsPrompt));

		// Remaining options require no special handling - loop through them
		for (auto& regKey : registry::RegKeys)
		{
			if (!regKey || !regKey->uiOptionsPrompt) continue;

			if (registry::regOptionType::check == regKey->ulRegOptType)
			{
				if (regKey->bRefresh && regKey->ulCurDWORD != static_cast<ULONG>(GetCheck(regKey->uiOptionsPrompt)))
				{
					m_bNeedPropRefresh = true;
				}

				regKey->ulCurDWORD = GetCheck(regKey->uiOptionsPrompt);
			}
			else if (registry::regOptionType::stringHex == regKey->ulRegOptType)
			{
				regKey->ulCurDWORD = GetHex(regKey->uiOptionsPrompt);
			}
			else if (registry::regOptionType::stringDec == regKey->ulRegOptType)
			{
				regKey->ulCurDWORD = GetDecimal(regKey->uiOptionsPrompt);
			}
		}

		// Commit our values to the registry
		registry::WriteToRegistry();

		CEditor::OnOK();
	}

	_Check_return_ ULONG COptions::HandleChange(const UINT nID) { return CEditor::HandleChange(nID); }

	bool COptions::NeedPropRefresh() const noexcept { return m_bNeedPropRefresh; }

	bool DisplayOptionsDlg(_In_ CWnd* lpParentWnd)
	{
		COptions MyOptions(lpParentWnd);
		MyOptions.DoModal();

		registry::PushOptionsToStub();
		return MyOptions.NeedPropRefresh();
	}
} // namespace dialog::editor
