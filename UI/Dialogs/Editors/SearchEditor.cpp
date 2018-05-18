#include "stdafx.h"
#include "SearchEditor.h"
#include <Interpret/InterpretProp2.h>
#include <Interpret/String.h>
#include "PropertyTagEditor.h"
#include <MAPI/MAPIFunctions.h>

static std::wstring CLASS = L"CSearchEditor";

vector<std::pair<ULONG, ULONG>> FuzzyLevels = {
	{ IDS_SEARCHSUBSTRING , FL_IGNORECASE | FL_SUBSTRING },
	{ IDS_SEARCHPREFIX, FL_IGNORECASE | FL_PREFIX },
	{ IDS_SEARCHFULLSTRING, FL_IGNORECASE | FL_FULLSTRING },
};

// TODO: Right now only supports string properties
CSearchEditor::CSearchEditor(
	ULONG ulPropTag,
	_In_opt_ LPMAPIPROP lpMAPIProp,
	_In_ CWnd* pParentWnd) :
	CEditor(pParentWnd,
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

	InitPane(SearchFields::PROPNAME, TextPane::CreateSingleLinePane(IDS_PROPNAME, true));
	InitPane(SearchFields::SEARCHTERM, TextPane::CreateSingleLinePane(IDS_PROPVALUE, false));
	InitPane(SearchFields::FUZZYLEVEL, DropDownPane::Create(IDS_SEARCHTYPE, 0, nullptr, false));
	InitPane(SearchFields::FINDROW, CheckPane::Create(IDS_APPLYUSINGFINDROW, false, false));

	// initialize our dropdowns here
	auto ulDropNum = 0;
	auto lpFuzzyPane = static_cast<DropDownPane*>(GetPane(SearchFields::FUZZYLEVEL));
	if (lpFuzzyPane)
	{
		for (auto fuzzyLevel : FuzzyLevels)
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

_Check_return_ LPSRestriction CSearchEditor::GetRestriction()
{
	auto hRes = S_OK;
	LPSRestriction lpRes = nullptr;
	// Allocate and create our SRestriction
	WC_H(CreatePropertyStringRestriction(
		CHANGE_PROP_TYPE(m_ulPropTag, PT_UNICODE),
		GetStringW(CSearchEditor::SearchFields::SEARCHTERM),
		m_ulFuzzyLevel,
		nullptr,
		&lpRes));
	if (S_OK != hRes)
	{
		MAPIFreeBuffer(lpRes);
		lpRes = nullptr;
	}

	return lpRes;
}

// Select a property tag
void CSearchEditor::OnEditAction1()
{
	auto hRes = S_OK;

	CPropertyTagEditor MyData(
		IDS_EDITGIVENPROP, // Title
		NULL, // Prompt
		m_ulPropTag,
		false,
		m_lpMAPIProp,
		this);

	WC_H(MyData.DisplayDialog());
	if (S_OK != hRes) return;

	m_ulPropTag = MyData.GetPropertyTag();
	PopulateFields(NOSKIPFIELD);
}

_Check_return_ ULONG CSearchEditor::HandleChange(UINT nID)
{
	auto i = (SearchFields)CEditor::HandleChange(nID);

	if (static_cast<ULONG>(-1) == i) return static_cast<ULONG>(-1);

	switch (i)
	{
	case SearchFields::FUZZYLEVEL:
		m_ulFuzzyLevel = GetSelectedFuzzyLevel();
		break;
	default:
		return i;
	}

	PopulateFields(i);

	return i;
}

// Fill out the fields in the form
// Don't touch the field passed in ulSkipField
// Pass NOSKIPFIELD to fill out all fields
void CSearchEditor::PopulateFields(ULONG ulSkipField) const
{
	if (SearchFields::PROPNAME != ulSkipField)
	{
		auto propTagNames = PropTagToPropName(m_ulPropTag, false);

		if (PROP_ID(m_ulPropTag) && !propTagNames.bestGuess.empty())
		{
			SetStringW(SearchFields::PROPNAME, propTagNames.bestGuess);
		}
		else
		{
			auto namePropNames = NameIDToStrings(
				m_ulPropTag,
				m_lpMAPIProp,
				nullptr,
				nullptr,
				false);

			if (!namePropNames.name.empty())
				SetStringW(SearchFields::PROPNAME, namePropNames.name);
			else
				LoadString(SearchFields::PROPNAME, IDS_UNKNOWNPROPERTY);
		}
	}
}

_Check_return_ ULONG CSearchEditor::GetSelectedFuzzyLevel() const
{
	auto lpFuzzyPane = static_cast<DropDownPane*>(GetPane(SearchFields::FUZZYLEVEL));
	if (lpFuzzyPane)
	{
		auto ulFuzzySelection = lpFuzzyPane->GetDropDownSelection();
		if (ulFuzzySelection == CB_ERR)
		{
			return strings::wstringToUlong(lpFuzzyPane->GetDropStringUseControl(), 16);
		}
		else
		{
			return FuzzyLevels[ulFuzzySelection].second;
		}
	}

	return FL_FULLSTRING;
}