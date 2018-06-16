#include <StdAfx.h>
#include <UI/Dialogs/Editors/SearchEditor.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/String.h>
#include <UI/Dialogs/Editors/PropertyTagEditor.h>
#include <MAPI/MAPIFunctions.h>
#include <MAPI/Cache/NamedPropCache.h>

namespace dialog
{
	namespace editor
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

			InitPane(PROPNAME, viewpane::TextPane::CreateSingleLinePane(IDS_PROPNAME, true));
			InitPane(SEARCHTERM, viewpane::TextPane::CreateSingleLinePane(IDS_PROPVALUE, false));
			InitPane(FUZZYLEVEL, viewpane::DropDownPane::Create(IDS_SEARCHTYPE, 0, nullptr, false));
			InitPane(FINDROW, viewpane::CheckPane::Create(IDS_APPLYUSINGFINDROW, false, false));

			// initialize our dropdowns here
			auto ulDropNum = 0;
			auto lpFuzzyPane = dynamic_cast<viewpane::DropDownPane*>(GetPane(FUZZYLEVEL));
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
			auto hRes = S_OK;
			LPSRestriction lpRes = nullptr;
			// Allocate and create our SRestriction
			WC_H(mapi::CreatePropertyStringRestriction(
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
			const auto i = static_cast<SearchFields>(CEditor::HandleChange(nID));

			if (i == static_cast<SearchFields>(-1)) return static_cast<ULONG>(-1);

			switch (i)
			{
			case FUZZYLEVEL:
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
			if (PROPNAME != ulSkipField)
			{
				auto propTagNames = interpretprop::PropTagToPropName(m_ulPropTag, false);

				if (PROP_ID(m_ulPropTag) && !propTagNames.bestGuess.empty())
				{
					SetStringW(PROPNAME, propTagNames.bestGuess);
				}
				else
				{
					auto namePropNames = cache::NameIDToStrings(m_ulPropTag, m_lpMAPIProp, nullptr, nullptr, false);

					if (!namePropNames.name.empty())
						SetStringW(PROPNAME, namePropNames.name);
					else
						LoadString(PROPNAME, IDS_UNKNOWNPROPERTY);
				}
			}
		}

		_Check_return_ ULONG CSearchEditor::GetSelectedFuzzyLevel() const
		{
			const auto lpFuzzyPane = dynamic_cast<viewpane::DropDownPane*>(GetPane(FUZZYLEVEL));
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
	}
}