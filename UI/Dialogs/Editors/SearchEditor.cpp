#include <StdAfx.h>
#include <UI/Dialogs/Editors/SearchEditor.h>
#include <core/utility/strings.h>
#include <UI/Dialogs/Editors/PropertyTagEditor.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/cache/namedPropCacheEntry2.h>
#include <core/mapi/cache/namedPropCache2.h>
#include <core/utility/output.h>
#include <core/interpret/proptags.h>
#include <core/addin/mfcmapi.h>

namespace dialog::editor
{
	static std::wstring CLASS = L"CSearchEditor";

	std::vector<std::pair<ULONG, ULONG>> FuzzyLevels = {
		{IDS_SEARCHSUBSTRING, FL_IGNORECASE | FL_SUBSTRING},
		{IDS_SEARCHPREFIX, FL_IGNORECASE | FL_PREFIX},
		{IDS_SEARCHFULLSTRING, FL_IGNORECASE | FL_FULLSTRING},
	};

	// TODO: Right now only supports string properties
	CSearchEditor::CSearchEditor(ULONG ulPropTag, _In_opt_ LPMAPIPROP lpMAPIProp, _In_ CWnd* pParentWnd)
		: CEditor(
			  pParentWnd,
			  IDS_SEARCHCRITERIA,
			  IDS_CREATEPROPRES,
			  CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL | CEDITOR_BUTTON_ACTION1,
			  IDS_ACTIONSELECTPTAG,
			  NULL,
			  NULL)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_ulPropTag = ulPropTag;
		m_ulFuzzyLevel = FuzzyLevels[0].second;

		m_lpMAPIProp = lpMAPIProp;
		if (m_lpMAPIProp) m_lpMAPIProp->AddRef();

		AddPane(viewpane::TextPane::CreateSingleLinePane(PROPNAME, IDS_PROPNAME, true));
		AddPane(viewpane::TextPane::CreateSingleLinePane(SEARCHTERM, IDS_PROPVALUE, false));
		AddPane(viewpane::DropDownPane::Create(FUZZYLEVEL, IDS_SEARCHTYPE, 0, nullptr, false));
		AddPane(viewpane::CheckPane::Create(FINDROW, IDS_APPLYUSINGFINDROW, false, false));

		// initialize our dropdowns here
		auto ulDropNum = 0;
		auto lpFuzzyPane = std::dynamic_pointer_cast<viewpane::DropDownPane>(GetPane(FUZZYLEVEL));
		if (lpFuzzyPane)
		{
			for (const auto fuzzyLevel : FuzzyLevels)
			{
				lpFuzzyPane->InsertDropString(strings::loadstring(fuzzyLevel.first), ulDropNum++);
			}

			lpFuzzyPane->SetDropDownSelection(strings::loadstring(FuzzyLevels[0].first));
		}

		PopulateFields(NOSKIPFIELD);
	}

	CSearchEditor::~CSearchEditor()
	{
		TRACE_DESTRUCTOR(CLASS);
		if (m_lpMAPIProp) m_lpMAPIProp->Release();
	}

	_Check_return_ const SRestriction* CSearchEditor::GetRestriction() const
	{
		// Allocate and create our SRestriction
		return mapi::CreatePropertyStringRestriction(
			CHANGE_PROP_TYPE(m_ulPropTag, PT_UNICODE), GetStringW(SEARCHTERM), m_ulFuzzyLevel, nullptr);
	}

	// Select a property tag
	void CSearchEditor::OnEditAction1()
	{
		CPropertyTagEditor MyData(
			IDS_EDITGIVENPROP, // Title
			NULL, // Prompt
			m_ulPropTag,
			false,
			m_lpMAPIProp,
			this);

		if (!MyData.DisplayDialog()) return;

		m_ulPropTag = MyData.GetPropertyTag();
		PopulateFields(NOSKIPFIELD);
	}

	_Check_return_ ULONG CSearchEditor::HandleChange(UINT nID)
	{
		const auto paneID = static_cast<SearchFields>(CEditor::HandleChange(nID));

		if (paneID == static_cast<SearchFields>(-1)) return static_cast<ULONG>(-1);

		switch (paneID)
		{
		case FUZZYLEVEL:
			m_ulFuzzyLevel = GetSelectedFuzzyLevel();
			break;
		default:
			return paneID;
		}

		PopulateFields(paneID);

		return paneID;
	}

	// Fill out the fields in the form
	// Don't touch the field passed in ulSkipField
	// Pass NOSKIPFIELD to fill out all fields
	void CSearchEditor::PopulateFields(ULONG ulSkipField) const
	{
		if (PROPNAME != ulSkipField)
		{
			auto propTagNames = proptags::PropTagToPropName(m_ulPropTag, false);

			if (PROP_ID(m_ulPropTag) && !propTagNames.bestGuess.empty())
			{
				SetStringW(PROPNAME, propTagNames.bestGuess);
			}
			else
			{
				auto namePropNames = cache::NameIDToStrings(m_ulPropTag, m_lpMAPIProp, nullptr, {}, false);

				if (!namePropNames.name.empty())
					SetStringW(PROPNAME, namePropNames.name);
				else
					LoadString(PROPNAME, IDS_UNKNOWNPROPERTY);
			}
		}
	}

	_Check_return_ ULONG CSearchEditor::GetSelectedFuzzyLevel() const
	{
		const auto lpFuzzyPane = std::dynamic_pointer_cast<viewpane::DropDownPane>(GetPane(FUZZYLEVEL));
		if (lpFuzzyPane)
		{
			const auto ulFuzzySelection = lpFuzzyPane->GetDropDownSelection();
			if (ulFuzzySelection == CB_ERR)
			{
				return strings::wstringToUlong(lpFuzzyPane->GetDropStringUseControl(), 16);
			}

			return FuzzyLevels[ulFuzzySelection].second;
		}

		return FL_FULLSTRING;
	}
} // namespace dialog::editor