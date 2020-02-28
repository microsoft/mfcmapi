#include <StdAfx.h>
#include <UI/Dialogs/Editors/PropertyTagEditor.h>
#include <core/utility/strings.h>
#include <core/mapi/cache/namedPropCache.h>
#include <core/interpret/guid.h>
#include <UI/Dialogs/Editors/PropertySelector.h>
#include <core/addin/addin.h>
#include <core/addin/mfcmapi.h>
#include <core/utility/output.h>
#include <core/interpret/proptype.h>
#include <core/interpret/proptags.h>

namespace dialog::editor
{
	static std::wstring CLASS = L"CPropertyTagEditor";

	enum __PropTagFields
	{
		PROPTAG_TAG,
		PROPTAG_ID,
		PROPTAG_TYPE,
		PROPTAG_NAME,
		PROPTAG_TYPESTRING,
		PROPTAG_NAMEPROPKIND,
		PROPTAG_NAMEPROPNAME,
		PROPTAG_NAMEPROPGUID,
	};

	CPropertyTagEditor::CPropertyTagEditor(
		UINT uidTitle,
		UINT uidPrompt,
		ULONG ulPropTag,
		bool bIsAB,
		_In_opt_ LPMAPIPROP lpMAPIProp,
		_In_ CWnd* pParentWnd)
		: CEditor(
			  pParentWnd,
			  uidTitle ? uidTitle : IDS_PROPTAGEDITOR,
			  uidPrompt,
			  CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL | CEDITOR_BUTTON_ACTION1 |
				  (lpMAPIProp ? CEDITOR_BUTTON_ACTION2 : 0),
			  IDS_ACTIONSELECTPTAG,
			  IDS_ACTIONCREATENAMEDPROP,
			  NULL)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_ulPropTag = ulPropTag;
		m_bIsAB = bIsAB;
		m_lpMAPIProp = lpMAPIProp;

		if (m_lpMAPIProp) m_lpMAPIProp->AddRef();

		AddPane(viewpane::TextPane::CreateSingleLinePane(PROPTAG_TAG, IDS_PROPTAG, false));
		AddPane(viewpane::TextPane::CreateSingleLinePane(PROPTAG_ID, IDS_PROPID, false));
		AddPane(viewpane::DropDownPane::Create(PROPTAG_TYPE, IDS_PROPTYPE, 0, nullptr, false));
		AddPane(viewpane::TextPane::CreateSingleLinePane(PROPTAG_NAME, IDS_PROPNAME, true));
		AddPane(viewpane::TextPane::CreateSingleLinePane(PROPTAG_TYPESTRING, IDS_PROPTYPE, true));

		// Map named properties if we can, but not for Address Books
		if (m_lpMAPIProp && !m_bIsAB)
		{
			AddPane(viewpane::DropDownPane::Create(PROPTAG_NAMEPROPKIND, IDS_NAMEPROPKIND, 0, nullptr, true));
			AddPane(viewpane::TextPane::CreateSingleLinePane(PROPTAG_NAMEPROPNAME, IDS_NAMEPROPNAME, false));
			AddPane(viewpane::DropDownPane::CreateGuid(PROPTAG_NAMEPROPGUID, IDS_NAMEPROPGUID, false));
		}

		// initialize our dropdowns here
		// prop types
		for (ULONG ulDropNum = 0; ulDropNum < PropTypeArray.size(); ulDropNum++)
		{
			InsertDropString(PROPTAG_TYPE, ulDropNum, PropTypeArray[ulDropNum].lpszName);
		}

		if (m_lpMAPIProp)
		{
			InsertDropString(PROPTAG_NAMEPROPKIND, 0, L"MNID_STRING"); // STRING_OK
			InsertDropString(PROPTAG_NAMEPROPKIND, 1, L"MNID_ID"); // STRING_OK
		}

		PopulateFields(NOSKIPFIELD);
	}

	CPropertyTagEditor::~CPropertyTagEditor()
	{
		TRACE_DESTRUCTOR(CLASS);
		if (m_lpMAPIProp) m_lpMAPIProp->Release();
	}

	_Check_return_ ULONG CPropertyTagEditor::GetPropertyTag() const { return m_ulPropTag; }

	// Select a property tag
	void CPropertyTagEditor::OnEditAction1()
	{
		CPropertySelector MyData(m_bIsAB, m_lpMAPIProp, this);

		if (!MyData.DisplayDialog()) return;

		m_ulPropTag = MyData.GetPropertyTag();
		PopulateFields(NOSKIPFIELD);
	}

	// GetNamesFromIDs - always with MAPI_CREATE
	void CPropertyTagEditor::OnEditAction2()
	{
		if (!m_lpMAPIProp) return;

		LookupNamedProp(NOSKIPFIELD, true);

		PopulateFields(NOSKIPFIELD);
	}

	// Search for properties matching lpszDispIDName on a substring
	_Check_return_ LPNAMEID_ARRAY_ENTRY GetDispIDFromName(_In_z_ LPCWSTR lpszDispIDName)
	{
		if (!lpszDispIDName) return nullptr;

		const auto entry = find_if(begin(NameIDArray), end(NameIDArray), [&](NAMEID_ARRAY_ENTRY& nameID) {
			if (0 == wcscmp(nameID.lpszName, lpszDispIDName))
			{
				// PSUNKNOWN is used as a placeholder in NameIDArray - don't return matching entries
				if (!IsEqualGUID(*nameID.lpGuid, guid::PSUNKNOWN)) return true;
			}

			return false;
		});

		return entry != end(NameIDArray) ? &(*entry) : nullptr;
	}

	void CPropertyTagEditor::LookupNamedProp(ULONG ulSkipField, bool bCreate)
	{
		auto ulPropType = GetSelectedPropType();

		MAPINAMEID NamedID = {};

		// Assume an ID to help with the dispid case
		NamedID.ulKind = MNID_ID;

		const auto iCurSel = GetDropDownSelection(PROPTAG_NAMEPROPKIND);
		if (iCurSel != CB_ERR)
		{
			if (0 == iCurSel) NamedID.ulKind = MNID_STRING;
			if (1 == iCurSel) NamedID.ulKind = MNID_ID;
		}

		auto szName = GetStringW(PROPTAG_NAMEPROPNAME);

		auto guid = GetSelectedGUID(PROPTAG_NAMEPROPGUID, false);
		NamedID.lpguid = &guid;

		if (MNID_ID == NamedID.ulKind)
		{
			// Now check if that string is a known dispid
			const auto lpNameIDEntry = GetDispIDFromName(szName.c_str());

			// If we matched on a dispid name, use that for our lookup
			// Note that we should only ever reach this case if the user typed a dispid name
			if (lpNameIDEntry)
			{
				NamedID.Kind.lID = lpNameIDEntry->lValue;
				NamedID.lpguid = const_cast<LPGUID>(lpNameIDEntry->lpGuid);
				ulPropType = lpNameIDEntry->ulType;

				// We found something in our lookup, but later GetIDsFromNames call may fail
				// Make sure we write what we found back to the dialog
				// However, don't overwrite the field the user changed
				if (PROPTAG_NAMEPROPKIND != ulSkipField)
					SetDropDownSelection(PROPTAG_NAMEPROPKIND, L"MNID_ID"); // STRING_OK

				if (PROPTAG_NAMEPROPGUID != ulSkipField)
				{
					SetDropDownSelection(PROPTAG_NAMEPROPGUID, guid::GUIDToString(lpNameIDEntry->lpGuid));
				}

				// This will accomplish setting the type field
				// If the stored type was PT_UNSPECIFIED, we'll just keep the user selected type
				if (PT_UNSPECIFIED != lpNameIDEntry->ulType)
				{
					m_ulPropTag = CHANGE_PROP_TYPE(m_ulPropTag, lpNameIDEntry->ulType);
				}
			}
			else
			{
				NamedID.Kind.lID = strings::wstringToUlong(szName, 16);
			}
		}
		else if (MNID_STRING == NamedID.ulKind)
		{
			NamedID.Kind.lpwstrName = const_cast<LPWSTR>(szName.c_str());
		}

		if (NamedID.lpguid &&
			(MNID_ID == NamedID.ulKind && NamedID.Kind.lID || MNID_STRING == NamedID.ulKind && NamedID.Kind.lpwstrName))
		{
			const auto lpNamedPropTags = cache::GetIDsFromNames(m_lpMAPIProp, {NamedID}, bCreate ? MAPI_CREATE : 0);
			if (lpNamedPropTags && lpNamedPropTags->cValues == 1)
			{
				m_ulPropTag = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[0], ulPropType);
			}

			MAPIFreeBuffer(lpNamedPropTags);
		}
	}

	_Check_return_ ULONG CPropertyTagEditor::GetSelectedPropType() const
	{
		std::wstring szType;
		const auto iCurSel = GetDropDownSelection(PROPTAG_TYPE);
		if (iCurSel != CB_ERR)
		{
			szType = PropTypeArray[iCurSel].lpszName;
		}
		else
		{
			szType = GetDropStringUseControl(PROPTAG_TYPE);
		}

		return proptype::PropTypeNameToPropType(szType);
	}

	_Check_return_ ULONG CPropertyTagEditor::HandleChange(UINT nID)
	{
		const auto paneID = CEditor::HandleChange(nID);

		if (paneID == static_cast<ULONG>(-1)) return static_cast<ULONG>(-1);

		switch (paneID)
		{
		case PROPTAG_TAG: // Prop tag changed
			m_ulPropTag = GetPropTag(PROPTAG_TAG);
			break;
		case PROPTAG_ID: // Prop ID changed
		{
			const auto szID = GetStringW(PROPTAG_ID);
			const auto ulID = strings::wstringToUlong(szID, 16);

			m_ulPropTag = PROP_TAG(PROP_TYPE(m_ulPropTag), ulID);
		}
		break;
		case PROPTAG_TYPE: // Prop Type changed
		{
			const auto ulType = GetSelectedPropType();

			m_ulPropTag = CHANGE_PROP_TYPE(m_ulPropTag, ulType);
		}
		break;
		case PROPTAG_NAMEPROPKIND:
		case PROPTAG_NAMEPROPNAME:
		case PROPTAG_NAMEPROPGUID:
			LookupNamedProp(paneID, false);
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
	void CPropertyTagEditor::PopulateFields(ULONG ulSkipField) const
	{
		auto namePropNames = cache::NameIDToStrings(m_ulPropTag, m_lpMAPIProp, nullptr, {}, m_bIsAB);

		if (PROPTAG_TAG != ulSkipField) SetHex(PROPTAG_TAG, m_ulPropTag);
		if (PROPTAG_ID != ulSkipField) SetStringf(PROPTAG_ID, L"0x%04X", PROP_ID(m_ulPropTag)); // STRING_OK
		if (PROPTAG_TYPE != ulSkipField) SetDropDownSelection(PROPTAG_TYPE, proptype::TypeToString(m_ulPropTag));
		if (PROPTAG_NAME != ulSkipField)
		{
			auto propTagNames = proptags::PropTagToPropName(m_ulPropTag, m_bIsAB);

			if (PROP_ID(m_ulPropTag) && !propTagNames.bestGuess.empty())
				SetStringf(
					PROPTAG_NAME,
					L"%ws (%ws)",
					propTagNames.bestGuess.c_str(),
					propTagNames.otherMatches.c_str()); // STRING_OK
			else if (!namePropNames.name.empty())
				SetStringf(PROPTAG_NAME, L"%ws", namePropNames.name.c_str()); // STRING_OK
			else
				LoadString(PROPTAG_NAME, IDS_UNKNOWNPROPERTY);
		}

		if (PROPTAG_TYPESTRING != ulSkipField) SetStringW(PROPTAG_TYPESTRING, proptype::TypeToString(m_ulPropTag));

		// Do a named property lookup and fill out fields
		// But only if PROPTAG_TAG or PROPTAG_ID is what the user changed
		// And never for Address Books
		if (m_lpMAPIProp && !m_bIsAB && (PROPTAG_TAG == ulSkipField || PROPTAG_ID == ulSkipField))
		{
			const auto name = cache::GetNameFromID(m_lpMAPIProp, m_ulPropTag, NULL);
			if (name->valid())
			{
				const auto mnid = name->getMapiNameId();
				if (mnid->ulKind == MNID_STRING)
				{
					if (PROPTAG_NAMEPROPKIND != ulSkipField)
						SetDropDownSelection(PROPTAG_NAMEPROPKIND, L"MNID_STRING"); // STRING_OK
					if (PROPTAG_NAMEPROPNAME != ulSkipField) SetStringW(PROPTAG_NAMEPROPNAME, mnid->Kind.lpwstrName);
				}
				else if (mnid->ulKind == MNID_ID)
				{
					if (PROPTAG_NAMEPROPKIND != ulSkipField)
						SetDropDownSelection(PROPTAG_NAMEPROPKIND, L"MNID_ID"); // STRING_OK
					if (PROPTAG_NAMEPROPNAME != ulSkipField) SetHex(PROPTAG_NAMEPROPNAME, mnid->Kind.lID);
				}
				else
				{
					if (PROPTAG_NAMEPROPNAME != ulSkipField) SetStringW(PROPTAG_NAMEPROPNAME, L"");
				}

				if (PROPTAG_NAMEPROPGUID != ulSkipField)
				{
					SetDropDownSelection(PROPTAG_NAMEPROPGUID, guid::GUIDToString(mnid->lpguid));
				}
			}
			else
			{
				if (PROPTAG_NAMEPROPKIND != ulSkipField && PROPTAG_NAMEPROPNAME != ulSkipField &&
					PROPTAG_NAMEPROPGUID != ulSkipField)
				{
					SetDropDownSelection(PROPTAG_NAMEPROPKIND, strings::emptystring);
					SetStringW(PROPTAG_NAMEPROPNAME, L"");
					SetDropDownSelection(PROPTAG_NAMEPROPGUID, strings::emptystring);
				}
			}
		}
	}

	void CPropertyTagEditor::SetDropDownSelection(ULONG i, _In_ const std::wstring& szText) const
	{
		auto lpPane = std::dynamic_pointer_cast<viewpane::DropDownPane>(GetPane(i));
		if (lpPane)
		{
			return lpPane->SetDropDownSelection(szText);
		}
	}

	_Check_return_ std::wstring CPropertyTagEditor::GetDropStringUseControl(ULONG id) const
	{
		const auto lpPane = std::dynamic_pointer_cast<viewpane::DropDownPane>(GetPane(id));
		if (lpPane)
		{
			return lpPane->GetDropStringUseControl();
		}

		return strings::emptystring;
	}

	_Check_return_ int CPropertyTagEditor::GetDropDownSelection(ULONG id) const
	{
		const auto lpPane = std::dynamic_pointer_cast<viewpane::DropDownPane>(GetPane(id));
		if (lpPane)
		{
			return lpPane->GetDropDownSelection();
		}

		return CB_ERR;
	}

	void CPropertyTagEditor::InsertDropString(ULONG id, int iRow, _In_ const std::wstring& szText) const
	{
		auto lpPane = std::dynamic_pointer_cast<viewpane::DropDownPane>(GetPane(id));
		if (lpPane)
		{
			lpPane->InsertDropString(szText, iRow);
		}
	}
} // namespace dialog::editor