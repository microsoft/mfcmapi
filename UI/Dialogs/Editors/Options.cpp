#include <StdAfx.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/Editors/Options.h>
#include <mapistub/library/stubutils.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/addin/mfcmapi.h>
#include <UI/UIFunctions.h>
#include <UI/ViewPane/SectionPane.h>

namespace dialog::editor
{
	class COptions : public CEditor
	{
	public:
		COptions(_In_ CWnd* pWnd);
		~COptions();
		bool NeedPropRefresh() const noexcept;

	private:
		void AddOptionPane(_In_ registry::__RegKey* regKey);
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

		const std::vector<std::pair<UINT, std::vector<registry::__RegKey*>>> optionGroups = {
			{0, {&registry::uiTheme}},
			{IDS_OPTIONS_ONLINE_CACHE, {&registry::forceMDBOnline, &registry::forceMapiNoCache}},
			{IDS_OPTIONS_TABLES_DISPLAY,
			 {&registry::throttleLevel,
			  &registry::doColumnNames,
			  &registry::editColumnsOnLoad,
			  &registry::allowDupeColumns,
			  &registry::onlyAdditionalProperties,
			  &registry::doSmartView}},
			{IDS_OPTIONS_PROPERTIES,
			 {&registry::parseNamedProps,
			  &registry::getPropNamesOnAllProps,
			  &registry::cacheNamedProps,
			  &registry::preferUnicodeProps,
			  &registry::useRowDataForSinglePropList,
			  &registry::useGetPropList,
			  &registry::suppressNotFound}},
			{IDS_OPTIONS_NOTIFICATIONS_OPENING,
			 {&registry::hierExpandNotifications,
			  &registry::hierRootNotifs,
			  &registry::useMessageRaw,
			  &registry::allowPersistCache}},
			{IDS_OPTIONS_MAPI_IMPLEMENTATION,
			 {&registry::preferOlmapi32,
			  &registry::forceOutlookMAPI,
			  &registry::forceSystemMAPI,
			  &registry::loadAddIns,
			  &registry::enableLegacyFeatures}},
			{IDS_OPTIONS_DIAGNOSTICS,
			 {&registry::uiDiag,
			  &registry::debugTag,
			  &registry::debugToFile,
			  &registry::debugFileName,
			  &registry::useIMAPIProgress,
			  &registry::heapEnableTerminationOnCorruption}}};

		for (const auto& [section, options] : optionGroups)
		{
			if (section) AddPane(std::make_shared<viewpane::SectionPane>(section, section));
			for (const auto& regKey : options)
			{
				AddOptionPane(regKey);
			}
		}
	}

	void COptions::AddOptionPane(_In_ registry::__RegKey* regKey)
	{
		if (!regKey || !regKey->uiOptionsPrompt) return;

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
			AddPane(viewpane::TextPane::CreateSingleLinePane(regKey->uiOptionsPrompt, regKey->uiOptionsPrompt, false));
			SetHex(regKey->uiOptionsPrompt, regKey->ulCurDWORD);
		}
		else if (registry::regOptionType::stringDec == regKey->ulRegOptType)
		{
			AddPane(viewpane::TextPane::CreateSingleLinePane(regKey->uiOptionsPrompt, regKey->uiOptionsPrompt, false));
			SetDecimal(regKey->uiOptionsPrompt, regKey->ulCurDWORD);
		}
		else if (registry::regOptionType::dropdown == regKey->ulRegOptType)
		{
			auto pane = viewpane::DropDownPane::Create(
				regKey->uiOptionsPrompt,
				regKey->uiOptionsPrompt,
				static_cast<ULONG>(regKey->uiDropDownOptions.size()),
				regKey->uiDropDownOptions.data(),
				true);
			const auto selection = regKey->ulCurDWORD < regKey->uiDropDownOptions.size() ? regKey->ulCurDWORD : 0;
			pane->SetDropDownSelection(static_cast<int>(selection));
			AddPane(pane);
		}
	}

	COptions::~COptions() { TRACE_DESTRUCTOR(CLASS); }

	void COptions::OnOK()
	{
		const auto oldTheme = registry::uiTheme.ulCurDWORD;
		const auto oldUIDiag = registry::uiDiag.ulCurDWORD;

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
			else if (registry::regOptionType::dropdown == regKey->ulRegOptType)
			{
				const auto selected = GetDropDown(regKey->uiOptionsPrompt);
				if (selected >= 0 && selected < static_cast<int>(regKey->uiDropDownOptions.size()))
				{
					regKey->ulCurDWORD = static_cast<DWORD>(selected);
				}
			}
		}

		// Commit our values to the registry
		registry::WriteToRegistry();

		if (oldTheme != registry::uiTheme.ulCurDWORD)
		{
			ui::ApplyThemeToAllWindows();
		}
		else if (oldUIDiag != registry::uiDiag.ulCurDWORD)
		{
			ui::RedrawAllWindows();
		}

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
