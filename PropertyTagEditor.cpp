// PropertyTagEditor.cpp : implementation file
//

#include "stdafx.h"
#include "PropertyTagEditor.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"
#include "PropTagArray.h"
#include "MAPIFunctions.h"

static TCHAR* CLASS = _T("CPropertyTagEditor");

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
									   UINT			uidTitle,
									   UINT			uidPrompt,
									   ULONG		ulPropTag,
									   BOOL			bIncludeABProps,
									   LPMAPIPROP	lpMAPIProp,
									   CWnd*		pParentWnd):
CEditor(pParentWnd,
		uidTitle?uidTitle:IDS_PROPTAGEDITOR,
		uidPrompt?uidPrompt:IDS_PROPTAGEDITORPROMPT,
		0,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL|CEDITOR_BUTTON_ACTION1|(lpMAPIProp?CEDITOR_BUTTON_ACTION2:0),
		IDS_ACTIONSELECTPTAG,
		IDS_ACTIONCREATENAMEDPROP)
{
	TRACE_CONSTRUCTOR(CLASS);
	m_ulPropTag = ulPropTag;
	m_bIncludeABProps = bIncludeABProps;
	m_lpMAPIProp = lpMAPIProp;

	if (m_lpMAPIProp) m_lpMAPIProp->AddRef();

	CreateControls(m_lpMAPIProp?8:5);
	InitSingleLine(	PROPTAG_TAG,IDS_PROPTAG,NULL,false);
	InitSingleLine(	PROPTAG_ID,IDS_PROPID,NULL,false);
	InitDropDown(	PROPTAG_TYPE,IDS_PROPTYPE,0,NULL,false);
	InitSingleLine(	PROPTAG_NAME,IDS_PROPNAME,NULL,true);
	InitSingleLine(	PROPTAG_TYPESTRING,IDS_PROPTYPE,NULL,true);
	if (m_lpMAPIProp)
	{
		InitDropDown(	PROPTAG_NAMEPROPKIND,IDS_NAMEPROPKIND,0,NULL,true);
		InitSingleLine(	PROPTAG_NAMEPROPNAME,IDS_NAMEPROPNAME,NULL,false);
		InitDropDown(	PROPTAG_NAMEPROPGUID,IDS_NAMEPROPGUID,0,NULL,false);
	}
}

CPropertyTagEditor::~CPropertyTagEditor()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_lpMAPIProp) m_lpMAPIProp->Release();
}

BOOL CPropertyTagEditor::OnInitDialog()
{
	BOOL bRet = CEditor::OnInitDialog();

	// initialize our dropdowns here
	ULONG ulDropNum = 0;

	// prop types
	for (ulDropNum=0 ; ulDropNum < ulPropTypeArray ; ulDropNum++)
	{
#ifdef UNICODE
		InsertDropString(PROPTAG_TYPE,ulDropNum,PropTypeArray[ulDropNum].lpszName);
#else
		HRESULT hRes = S_OK;
		LPSTR szAnsiName = NULL;
		EC_H(UnicodeToAnsi(PropTypeArray[ulDropNum].lpszName,&szAnsiName));
		if (SUCCEEDED(hRes))
		{
			InsertDropString(PROPTAG_TYPE,ulDropNum,szAnsiName);
		}
		delete[] szAnsiName;
#endif
	}

	if (m_lpMAPIProp)
	{
		InsertDropString(PROPTAG_NAMEPROPKIND,0,_T("MNID_STRING")); // STRING_OK
		InsertDropString(PROPTAG_NAMEPROPKIND,1,_T("MNID_ID")); // STRING_OK

		for (ulDropNum=0 ; ulDropNum < ulPropGuidArray ; ulDropNum++)
		{
			LPTSTR szGUID = GUIDToStringAndName(PropGuidArray[ulDropNum].lpGuid);
			InsertDropString(PROPTAG_NAMEPROPGUID,ulDropNum,szGUID);
			delete[] szGUID;
		}
	}

	PopulateFields(NOSKIPFIELD);

	return bRet;
}

ULONG CPropertyTagEditor::GetPropertyTag()
{
	return m_ulPropTag;
}

// Select a property tag
void CPropertyTagEditor::OnEditAction1()
{
	HRESULT hRes = S_OK;

	CPropertySelector MyData(
			m_bIncludeABProps,
			m_lpMAPIProp,
			this);

	WC_H(MyData.DisplayDialog());
	if (S_OK != hRes) return;

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

void CPropertyTagEditor::LookupNamedProp(ULONG ulSkipField, BOOL bCreate)
{
	HRESULT hRes = S_OK;

	ULONG ulPropType = PT_NULL;
	GetSelectedPropType(&ulPropType);

	GUID guid = {0};
	MAPINAMEID NamedID = {0};
	LPMAPINAMEID lpNamedID = NULL;
	lpNamedID = &NamedID;

	// Assume an ID to help with the dispid case
	NamedID.ulKind = MNID_ID;

	int iCurSel = 0;
	iCurSel = GetDropDownSelection(PROPTAG_NAMEPROPKIND);
	if (iCurSel != CB_ERR)
	{
		if (0 == iCurSel) NamedID.ulKind = MNID_STRING;
		if (1 == iCurSel) NamedID.ulKind = MNID_ID;
	}
	CString szName;
	LPWSTR	szWideName = NULL;
	szName = GetStringUseControl(PROPTAG_NAMEPROPNAME);

	// Convert our prop tag name to a wide character string
#ifdef UNICODE
	szWideName = (LPWSTR) (LPCWSTR) szName;
#else
	EC_H(AnsiToUnicode(szName,&szWideName));
#endif

	if (GetSelectedGUID(&guid))
	{
		NamedID.lpguid = &guid;
	}

	// Now check if that string is a known dispid
	LPNAMEID_ARRAY_ENTRY lpNameIDEntry = NULL;

	if (MNID_ID == NamedID.ulKind)
	{
		lpNameIDEntry = GetDispIDFromName(szWideName);

		// If we matched on a dispid name, use that for our lookup
		// Note that we should only ever reach this case if the user typed a dispid name
		if (lpNameIDEntry)
		{
			NamedID.Kind.lID = lpNameIDEntry->lValue;
			NamedID.lpguid = (LPGUID) lpNameIDEntry->lpGuid;
			ulPropType = lpNameIDEntry->ulType;
			lpNamedID = &NamedID;

			// We found something in our lookup, but later GetIDsFromNames call may fail
			// Make sure we write what we found back to the dialog
			// However, don't overwrite the field the user changed
			if (PROPTAG_NAMEPROPKIND != ulSkipField) SetDropDownSelection(PROPTAG_NAMEPROPKIND,_T("MNID_ID")); // STRING_OK

			if (PROPTAG_NAMEPROPGUID != ulSkipField)
			{
				LPTSTR szGUID = GUIDToString(lpNameIDEntry->lpGuid);
				SetDropDownSelection(PROPTAG_NAMEPROPGUID,szGUID);
				delete[] szGUID;
			}

			// This will accomplish setting the type field
			// If the stored type was PT_UNSPECIFIED, we'll just keep the user selected type
			if (PT_UNSPECIFIED != lpNameIDEntry->ulType)
			{
				m_ulPropTag = CHANGE_PROP_TYPE(m_ulPropTag,lpNameIDEntry->ulType);
			}
		}
		else
		{
			NamedID.Kind.lID = _tcstoul((LPCTSTR)szName,NULL,16);
		}
	}
	else if (MNID_STRING == NamedID.ulKind)
	{
		NamedID.Kind.lpwstrName = szWideName;
	}

	if (NamedID.lpguid &&
		((MNID_ID == NamedID.ulKind && NamedID.Kind.lID) || (MNID_STRING == NamedID.ulKind && NamedID.Kind.lpwstrName)))
	{
		LPSPropTagArray lpNamedPropTags = NULL;

		EC_H(m_lpMAPIProp->GetIDsFromNames(
			1,
			&lpNamedID,
			bCreate?MAPI_CREATE:0,
			&lpNamedPropTags));

		if (lpNamedPropTags)
		{
			m_ulPropTag = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[0],ulPropType);
			MAPIFreeBuffer(lpNamedPropTags);
		}
	}

	if (MNID_STRING == NamedID.ulKind)
	{
		MAPIFreeBuffer(NamedID.Kind.lpwstrName);
	}
#ifndef UNICODE
	delete[] szWideName;
#endif
}

BOOL CPropertyTagEditor::GetSelectedGUID(LPGUID lpSelectedGUID)
{
	if (!lpSelectedGUID) return false;
	if (!IsValidDropDown(PROPTAG_NAMEPROPGUID)) return false;

	LPCGUID lpGUID = NULL;
	int iCurSel = 0;
	iCurSel = GetDropDownSelection(PROPTAG_NAMEPROPGUID);
	if (iCurSel != CB_ERR)
	{
		lpGUID = PropGuidArray[iCurSel].lpGuid;
	}
	else
	{
		// no match - need to do a lookup
		CString szText;
		GUID	guid = {0};
		szText = GetDropStringUseControl(PROPTAG_NAMEPROPGUID);
		// try the GUID like PS_* first
		GUIDNameToGUID((LPCTSTR) szText,&lpGUID);
		if (!lpGUID) // no match - try it like a guid {}
		{
			HRESULT hRes = S_OK;
			WC_H(StringToGUID((LPCTSTR) szText,&guid));

			if (SUCCEEDED(hRes))
			{
				lpGUID = &guid;
			}
		}
	}
	if (lpGUID)
	{
		memcpy(lpSelectedGUID,lpGUID,sizeof(GUID));
		return true;
	}
	return false;
}

BOOL CPropertyTagEditor::GetSelectedPropType(ULONG* ulPropType)
{
	if (!ulPropType) return false;
	if (!IsValidDropDown(PROPTAG_TYPE)) return false;

	HRESULT hRes = S_OK;
	CString szType;
	int iCurSel = 0;
	iCurSel = GetDropDownSelection(PROPTAG_TYPE);
	if (iCurSel != CB_ERR)
	{
		szType = PropTypeArray[iCurSel].lpszName;
	}
	else
	{
		szType = GetDropStringUseControl(PROPTAG_TYPE);
	}
	LPTSTR szEnd = NULL;
	ULONG ulType = _tcstoul((LPCTSTR) szType,&szEnd,16);

	if (*szEnd != NULL) // If we didn't consume the whole string, try a lookup
	{
		EC_H(PropTypeNameToPropType((LPCTSTR) szType,&ulType));
	}

	*ulPropType = ulType;
	return true;
}

ULONG CPropertyTagEditor::HandleChange(UINT nID)
{
	HRESULT hRes = S_OK;
	ULONG i = CEditor::HandleChange(nID);

	if ((ULONG) -1 == i) return (ULONG) -1;

	switch (i)
	{
	case(PROPTAG_TAG): // Prop tag changed
		{
			CString szTag;
			szTag = GetStringUseControl(PROPTAG_TAG);
			LPTSTR szEnd = NULL;
			ULONG ulTag = _tcstoul((LPCTSTR) szTag,&szEnd,16);

			if (*szEnd != NULL) // If we didn't consume the whole string, try a lookup
			{
				EC_H(PropNameToPropTag((LPCTSTR) szTag,&ulTag));
			}

			// Figure if this is a full tag or just an ID
			if (ulTag & 0xffff0000) // Full prop tag
			{
				m_ulPropTag = ulTag;
			}
			else // Just an ID
			{
				m_ulPropTag = PROP_TAG(PT_UNSPECIFIED,ulTag);
			}
		}
		break;
	case(PROPTAG_ID): // Prop ID changed
		{
			CString szID;
			szID = GetStringUseControl(PROPTAG_ID);
			ULONG ulID = _tcstoul((LPCTSTR) szID,NULL,16);

			m_ulPropTag = PROP_TAG(PROP_TYPE(m_ulPropTag),ulID);
		}
		break;
	case(PROPTAG_TYPE): // Prop Type changed
		{
			ULONG ulType = PT_NULL;
			GetSelectedPropType(&ulType);

			m_ulPropTag = CHANGE_PROP_TYPE(m_ulPropTag,ulType);
		}
		break;
	case(PROPTAG_NAMEPROPKIND):
	case(PROPTAG_NAMEPROPNAME):
	case(PROPTAG_NAMEPROPGUID):
		{
			LookupNamedProp(i, false);
		}
		break;
	default:
		return i;
		break;
	}

	PopulateFields(i);

	return i;
}

// Fill out the fields in the form
// Don't touch the field passed in ulSkipField
// Pass NOSKIPFIELD to fill out all fields
void CPropertyTagEditor::PopulateFields(ULONG ulSkipField)
{
	HRESULT hRes = S_OK;

	CString PropType;
	LPTSTR szExactMatch = NULL;
	LPTSTR szPartialMatch = NULL;
	LPTSTR szNamedPropName = NULL;

	InterpretProp(
		NULL,
		m_ulPropTag,
		m_lpMAPIProp,
		NULL,
		m_bIncludeABProps,
		&szExactMatch, // Built from ulPropTag & bIsAB
		&szPartialMatch, // Built from ulPropTag & bIsAB
		&PropType,
		NULL,
		NULL,
		NULL,
		NULL,
		&szNamedPropName,
		NULL,
		NULL);

	if (PROPTAG_TAG != ulSkipField) SetHex(PROPTAG_TAG,m_ulPropTag);
	if (PROPTAG_ID != ulSkipField) SetStringf(PROPTAG_ID,_T("0x%04X"),PROP_ID(m_ulPropTag)); // STRING_OK
	if (PROPTAG_TYPE != ulSkipField) SetDropDownSelection(PROPTAG_TYPE,PropType);
	if (PROPTAG_NAME != ulSkipField)
	{
		if (PROP_ID(m_ulPropTag) && (szExactMatch || szPartialMatch))
			SetStringf(PROPTAG_NAME,_T("%s (%s)"),szExactMatch?szExactMatch:_T(""),szPartialMatch?szPartialMatch:_T("")); // STRING_OK
		else if (szNamedPropName)
			SetStringf(PROPTAG_NAME,_T("%s"),szNamedPropName); // STRING_OK
		else
			LoadString(PROPTAG_NAME,IDS_UNKNOWNPROPERTY);
	}
	if (PROPTAG_TYPESTRING != ulSkipField) SetString(PROPTAG_TYPESTRING,(LPCTSTR) TypeToString(m_ulPropTag));

	// do a named property lookup and fill out fields
	// but only if PROPTAG_TAG or PROPTAG_ID is what the user changed
	if (m_lpMAPIProp &&
		(PROPTAG_TAG == ulSkipField || PROPTAG_ID == ulSkipField ))
	{
		ULONG			ulPropNames = 0;
		SPropTagArray	sTagArray = {0};
		LPSPropTagArray lpTagArray = &sTagArray;
		LPMAPINAMEID*	lppPropNames = NULL;

		lpTagArray->cValues = 1;
		lpTagArray->aulPropTag[0] = m_ulPropTag;

		WC_H_GETPROPS(m_lpMAPIProp->GetNamesFromIDs(
			&lpTagArray,
			NULL,
			NULL,
			&ulPropNames,
			&lppPropNames));
		if (SUCCEEDED(hRes) && ulPropNames == lpTagArray->cValues && lppPropNames && lppPropNames[0])
		{
			if (MNID_STRING == lppPropNames[0]->ulKind)
			{
				if (PROPTAG_NAMEPROPKIND != ulSkipField) SetDropDownSelection(PROPTAG_NAMEPROPKIND,_T("MNID_STRING")); // STRING_OK
				if (PROPTAG_NAMEPROPNAME != ulSkipField) SetStringW(PROPTAG_NAMEPROPNAME,lppPropNames[0]->Kind.lpwstrName);
			}
			else if (MNID_ID == lppPropNames[0]->ulKind)
			{
				if (PROPTAG_NAMEPROPKIND != ulSkipField) SetDropDownSelection(PROPTAG_NAMEPROPKIND,_T("MNID_ID")); // STRING_OK
				if (PROPTAG_NAMEPROPNAME != ulSkipField) SetHex(PROPTAG_NAMEPROPNAME,lppPropNames[0]->Kind.lID);
			}
			else
			{
				if (PROPTAG_NAMEPROPNAME != ulSkipField) SetString(PROPTAG_NAMEPROPNAME,NULL);
			}
			if (PROPTAG_NAMEPROPGUID != ulSkipField)
			{
				LPTSTR szGUID = GUIDToString(lppPropNames[0]->lpguid);
				SetDropDownSelection(PROPTAG_NAMEPROPGUID,szGUID);
				delete[] szGUID;
			}
		}
		else
		{
			if (PROPTAG_NAMEPROPKIND != ulSkipField &&
				PROPTAG_NAMEPROPNAME != ulSkipField &&
				PROPTAG_NAMEPROPGUID != ulSkipField) 
			{
				SetDropDownSelection(PROPTAG_NAMEPROPKIND,NULL);
				SetString(PROPTAG_NAMEPROPNAME,NULL);
				SetDropDownSelection(PROPTAG_NAMEPROPGUID,NULL);
			}
		}
		MAPIFreeBuffer(lppPropNames);
	}

	delete[] szNamedPropName;
	delete[] szPartialMatch;
	delete[] szExactMatch;
}


//////////////////////////////////////////////////////////////////////////////////////
// CPropertySelector
// Property selection dialog
// Displays a list of known property tags - no add or delete
//////////////////////////////////////////////////////////////////////////////////////

CPropertySelector::CPropertySelector(
									   BOOL			bIncludeABProps,
									   LPMAPIPROP	lpMAPIProp,
									   CWnd*		pParentWnd):
CEditor(pParentWnd,IDS_PROPSELECTOR,IDS_PROPSELECTORPROMPT,0,CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(CLASS);
	m_ulPropTag = PR_NULL;
	m_bIncludeABProps = bIncludeABProps;
	m_lpMAPIProp = lpMAPIProp;

	if (m_lpMAPIProp) m_lpMAPIProp->AddRef();

	CreateControls(1);
	InitList(0,IDS_KNOWNPROPTAGS,true,true);
}

CPropertySelector::~CPropertySelector()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_lpMAPIProp) m_lpMAPIProp->Release();
}

BOOL CPropertySelector::OnInitDialog()
{
	BOOL bRet = CEditor::OnInitDialog();

	if (IsValidList(0))
	{
		CString szTmp;
		SortListData* lpData = NULL;
		ULONG i = 0;
		InsertColumn(0,1,IDS_PROPERTYNAMES);
		InsertColumn(0,2,IDS_TAG);
		InsertColumn(0,3,IDS_TYPE);

		ULONG ulCurRow = 0;
		for (i = 0;i< ulPropTagArray;i++)
		{
			if (!m_bIncludeABProps && (PropTagArray[i].ulValue & 0x80000000)) continue;
#ifdef UNICODE
			lpData = InsertListRow(0,ulCurRow,PropTagArray[i].lpszName);
#else
			HRESULT hRes = S_OK;
			LPSTR szAnsiName = NULL;
			EC_H(UnicodeToAnsi(PropTagArray[i].lpszName,&szAnsiName));
			if (SUCCEEDED(hRes))
			{
				lpData = InsertListRow(0,ulCurRow,szAnsiName);
			}
			delete[] szAnsiName;
#endif

			if (lpData)
			{
				lpData->ulSortDataType = SORTLIST_PROP;
				lpData->data.Prop.ulPropTag = PropTagArray[i].ulValue;
				lpData->bItemFullyLoaded = true;
			}

			szTmp.Format(_T("0x%08X"),PropTagArray[i].ulValue); // STRING_OK
			SetListString(0,ulCurRow,1,szTmp);
			SetListString(0,ulCurRow,2,TypeToString(PropTagArray[i].ulValue));
			ulCurRow++;
		}

		// Initial sort is by property tag
		ResizeList(0,true);
	}

	return bRet;
}

void CPropertySelector::OnOK()
{
	SortListData* lpListData = GetSelectedListRowData(0);
	if (lpListData)
		m_ulPropTag = lpListData->data.Prop.ulPropTag;

	CEditor::OnOK();
}

// We're not actually editing the list here - just overriding this to allow double-click
// So it's OK to return false
BOOL CPropertySelector::DoListEdit(ULONG /*ulListNum*/, int /*iItem*/, SortListData* /*lpData*/)
{
	OnOK();
	return false;
}

ULONG CPropertySelector::GetPropertyTag()
{
	return m_ulPropTag;
}
