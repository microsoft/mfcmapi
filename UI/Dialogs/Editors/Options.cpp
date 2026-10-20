#include <StdAfx.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/Editors/Options.h>
#include <UI/ViewPane/AddInApprovalPane.h>
#include <mapistub/library/stubutils.h>
#include <core/utility/registry.h>
#include <core/utility/output.h>
#include <core/addin/mfcmapi.h>
#include <core/addin/addin.h>
#include <UI/UIFunctions.h>
#include <UI/ViewPane/SectionPane.h>
#include <map>

namespace dialog::editor
{
	static constexpr int ADDIN_APPROVAL_PANE_ID = 30000;

	struct OptionValue
	{
		DWORD number{};
		std::wstring text;
	};

	using OptionValues = std::map<registry::__RegKey*, OptionValue>;

	class COptions : public CEditor
	{
	public:
		COptions(_In_ CWnd* pWnd, OptionValues& values, bool advanced = false);
		~COptions();
		bool NeedPropRefresh() const noexcept;

	private:
		void AddOptionPane(_In_ registry::__RegKey* regKey);
		void OnOK() override;
		void OnEditAction1() override;

		OptionValues& m_values;
		std::vector<registry::__RegKey*> m_options;
		bool m_advanced{};
		bool m_bNeedPropRefresh{};
	};

	static std::wstring CLASS = L"COptions";

	COptions::COptions(_In_ CWnd* pWnd, OptionValues& values, bool advanced)
		: CEditor(
			  pWnd,
			  advanced ? IDS_ADVANCED_OPTIONS : IDS_SETOPTS,
			  NULL,
			  CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL | (advanced ? 0 : CEDITOR_BUTTON_ACTION1),
			  IDS_ACTION_ADVANCED_OPTIONS,
			  NULL,
			  NULL),
		  m_values(values), m_advanced(advanced)
	{
		TRACE_CONSTRUCTOR(CLASS);
		EnableScroll();

		output::DebugPrintEx(
			output::dbgLevel::Generic, CLASS, L"COptions(", L"Building option sheet - adding fields\n");

		std::vector<std::pair<UINT, std::vector<registry::__RegKey*>>> optionGroups;
		if (advanced)
		{
			optionGroups = {
				{IDS_OPTIONS_TABLES_DISPLAY,
				 {&registry::throttleLevel,
				  &registry::editColumnsOnLoad,
				  &registry::allowDupeColumns,
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
				  &registry::enableLegacyFeatures}},
				{IDS_OPTIONS_ADDINS, {}},
				{IDS_OPTIONS_DIAGNOSTICS,
				 {&registry::uiDiag, &registry::useIMAPIProgress, &registry::heapEnableTerminationOnCorruption}}};
		}
		else
		{
			optionGroups = {
				{0, {&registry::uiTheme}},
				{IDS_OPTIONS_ONLINE_CACHE, {&registry::forceMDBOnline, &registry::forceMapiNoCache}},
				{IDS_OPTIONS_TABLES_DISPLAY, {&registry::doColumnNames, &registry::onlyAdditionalProperties}},
				{IDS_OPTIONS_ADDINS, {&registry::addInDirectory}},
				{IDS_OPTIONS_LOGGING, {&registry::debugTag, &registry::debugToFile, &registry::debugFileName}}};
		}

		for (const auto& [section, options] : optionGroups)
		{
			if (section) AddPane(std::make_shared<viewpane::SectionPane>(section, section));
			for (const auto& regKey : options)
			{
				AddOptionPane(regKey);
			}
			if (advanced && section == IDS_OPTIONS_ADDINS)
			{
				AddPane(viewpane::AddInApprovalPane::Create(ADDIN_APPROVAL_PANE_ID));
			}
		}
	}

	void COptions::AddOptionPane(_In_ registry::__RegKey* regKey)
	{
		if (!regKey || !regKey->uiOptionsPrompt) return;
		m_options.push_back(regKey);
		const auto& value = m_values.at(regKey);

		if (registry::regOptionType::check == regKey->ulRegOptType)
		{
			AddPane(viewpane::CheckPane::Create(
				regKey->uiOptionsPrompt, regKey->uiOptionsPrompt, 0 != value.number, false));
		}
		else if (registry::regOptionType::string == regKey->ulRegOptType)
		{
			if (regKey == &registry::addInDirectory)
			{
				auto pane = viewpane::PathPane::Create(
					regKey->uiOptionsPrompt, regKey->uiOptionsPrompt, viewpane::pathType::directory, value.text);
				pane->SetHelpText(IDS_ADDINDIRECTORYHELP);
				AddPane(pane);
			}
			else if (regKey == &registry::debugFileName)
			{
				AddPane(viewpane::PathPane::Create(
					regKey->uiOptionsPrompt,
					regKey->uiOptionsPrompt,
					viewpane::pathType::fileSave,
					value.text,
					L"log"));
			}
			else
			{
				AddPane(viewpane::TextPane::CreateSingleLinePane(
					regKey->uiOptionsPrompt, regKey->uiOptionsPrompt, value.text, false));
			}
		}
		else if (registry::regOptionType::stringHex == regKey->ulRegOptType)
		{
			AddPane(viewpane::TextPane::CreateSingleLinePane(regKey->uiOptionsPrompt, regKey->uiOptionsPrompt, false));
			SetHex(regKey->uiOptionsPrompt, value.number);
		}
		else if (registry::regOptionType::stringDec == regKey->ulRegOptType)
		{
			AddPane(viewpane::TextPane::CreateSingleLinePane(regKey->uiOptionsPrompt, regKey->uiOptionsPrompt, false));
			SetDecimal(regKey->uiOptionsPrompt, value.number);
		}
		else if (registry::regOptionType::dropdown == regKey->ulRegOptType)
		{
			auto pane = viewpane::DropDownPane::Create(
				regKey->uiOptionsPrompt,
				regKey->uiOptionsPrompt,
				static_cast<ULONG>(regKey->uiDropDownOptions.size()),
				regKey->uiDropDownOptions.data(),
				true);
			const auto selection = value.number < regKey->uiDropDownOptions.size() ? value.number : 0;
			pane->SetDropDownSelection(static_cast<int>(selection));
			AddPane(pane);
		}
	}

	COptions::~COptions() { TRACE_DESTRUCTOR(CLASS); }

	void COptions::OnOK()
	{
		if (!m_advanced)
		{
			const auto directory = GetStringW(IDS_REGKEY_ADDINDIRECTORY);
			const auto validationResult = addin::ValidateAddInDirectory(directory);
			if (validationResult != ERROR_SUCCESS)
			{
				error::ErrDialog(__FILE__, __LINE__, IDS_ADDINDIRECTORYERROR, directory.c_str(), validationResult);
				return;
			}

			if (directory != registry::addInDirectory.szCurSTRING && !addin::ClearAddInApprovalCache())
			{
				error::ErrDialog(__FILE__, __LINE__, IDS_ADDINAPPROVALRESETFAILED);
				return;
			}
		}

		for (const auto regKey : m_options)
		{
			auto& value = m_values.at(regKey);
			if (registry::regOptionType::check == regKey->ulRegOptType)
			{
				value.number = GetCheck(regKey->uiOptionsPrompt);
			}
			else if (registry::regOptionType::string == regKey->ulRegOptType)
			{
				value.text = GetStringW(regKey->uiOptionsPrompt);
			}
			else if (registry::regOptionType::stringHex == regKey->ulRegOptType)
			{
				value.number = GetHex(regKey->uiOptionsPrompt);
			}
			else if (registry::regOptionType::stringDec == regKey->ulRegOptType)
			{
				value.number = GetDecimal(regKey->uiOptionsPrompt);
			}
			else if (registry::regOptionType::dropdown == regKey->ulRegOptType)
			{
				const auto selected = GetDropDown(regKey->uiOptionsPrompt);
				if (selected >= 0 && selected < static_cast<int>(regKey->uiDropDownOptions.size()))
				{
					value.number = static_cast<DWORD>(selected);
				}
			}
		}

		// Advanced edits stay pending until the enclosing Options dialog is accepted.
		if (m_advanced)
		{
			CEditor::OnOK();
			return;
		}

		const auto oldTheme = registry::uiTheme.ulCurDWORD;
		const auto oldUIDiag = registry::uiDiag.ulCurDWORD;
		const auto oldDebugTag = registry::debugTag.ulCurDWORD;

		for (const auto& [regKey, value] : m_values)
		{
			if (registry::regOptionType::string == regKey->ulRegOptType)
			{
				regKey->szCurSTRING = value.text;
			}
			else if (regKey != &registry::debugToFile)
			{
				if (regKey->bRefresh && regKey->ulCurDWORD != value.number) m_bNeedPropRefresh = true;
				regKey->ulCurDWORD = value.number;
			}
		}

		if (oldDebugTag != registry::debugTag.ulCurDWORD)
		{
			output::outputVersion(output::dbgLevel::VersionBanner, nullptr);
		}

		output::SetDebugOutputToFile(m_values.at(&registry::debugToFile).number != 0);

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

	void COptions::OnEditAction1()
	{
		COptions advancedOptions(this, m_values, true);
		advancedOptions.DoModal();
	}

	bool COptions::NeedPropRefresh() const noexcept { return m_bNeedPropRefresh; }

	bool DisplayOptionsDlg(_In_ CWnd* lpParentWnd)
	{
		OptionValues values;
		for (const auto regKey : registry::RegKeys)
		{
			if (regKey && regKey->uiOptionsPrompt)
			{
				values.emplace(regKey, OptionValue{regKey->ulCurDWORD, regKey->szCurSTRING});
			}
		}

		COptions MyOptions(lpParentWnd, values);
		MyOptions.DoModal();

		registry::PushOptionsToStub();
		return MyOptions.NeedPropRefresh();
	}
} // namespace dialog::editor
