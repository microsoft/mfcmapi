// PropertyTagEditor.cpp : implementation file
//

#include "stdafx.h"
#include "Error.h"

#include "PropertyTagEditor.h"

#include "InterpretProp.h"
#include "InterpretProp2.h"
#include "PropTagArray.h"
#include "MAPIFunctions.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

static TCHAR* CLASS = _T("CPropertyTagEditor");

CPropertyTagEditor::CPropertyTagEditor(
									   UINT			uidTitle,
									   UINT			uidPrompt,
									   ULONG		ulPropTag,
									   BOOL			bIncludeABProps,
									   LPMAPIPROP	lpMAPIProp,
									   CWnd*		pParentWnd):
CEditor(pParentWnd,IDS_PROPTAGEDITOR,IDS_PROPTAGEDITORPROMPT,0,CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL|CEDITOR_BUTTON_ACTION1|(lpMAPIProp?CEDITOR_BUTTON_ACTION2:0))
{
	TRACE_CONSTRUCTOR(CLASS);
	if (uidTitle) m_uidTitle = uidTitle;
	if (uidPrompt) m_uidPrompt = uidPrompt;
	m_ulPropTag = ulPropTag;
	m_bIncludeABProps = bIncludeABProps;
	m_lpMAPIProp = lpMAPIProp;

	if (m_lpMAPIProp) m_lpMAPIProp->AddRef();

	m_uidActionButtonText1 = IDS_ACTIONSELECTPTAG;
	m_uidActionButtonText2 = IDS_ACTIONLOOKUPNAMEDPROP;
	CreateControls(m_lpMAPIProp?8:5);
	InitSingleLine(	0,IDS_PROPTAG,NULL,false);
	InitSingleLine(	1,IDS_PROPID,NULL,false);
	InitDropDown(	2,IDS_PROPTYPE,0,NULL,false);
	InitSingleLine(	3,IDS_PROPNAME,NULL,true);
	InitSingleLine(	4,IDS_PROPTYPE,NULL,true);
	if (m_lpMAPIProp)
	{
		InitDropDown(	5,IDS_NAMEPROPKIND,0,NULL,true);
		InitSingleLine(	6,IDS_NAMEPROPNAME,NULL,false);
		InitDropDown(	7,IDS_NAMEPROPGUID,0,NULL,false);
	}
}

CPropertyTagEditor::~CPropertyTagEditor()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_lpMAPIProp) m_lpMAPIProp->Release();
}

BEGIN_MESSAGE_MAP(CPropertyTagEditor, CEditor)
//{{AFX_MSG_MAP(CPropertyTagEditor)
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BOOL CPropertyTagEditor::OnInitDialog()
{
	HRESULT hRes = S_OK;

	EC_B(CEditor::OnInitDialog());

	//initialize our dropdowns here
	ULONG ulDropNum = 0;

	//prop types
	if (m_lpControls[2].UI.lpDropDown)
	{
		m_lpControls[2].UI.lpDropDown->ulDropList = ulPropTypeArray;
		for (ulDropNum=0 ; ulDropNum < ulPropTypeArray ; ulDropNum++)
		{
#ifdef UNICODE
			m_lpControls[2].UI.lpDropDown->DropDown.InsertString(
				ulDropNum,
				PropTypeArray[ulDropNum].lpszName);
#else
			LPSTR szAnsiName = NULL;
			EC_H(UnicodeToAnsi(PropTypeArray[ulDropNum].lpszName,&szAnsiName));
			if (SUCCEEDED(hRes))
			{
				m_lpControls[2].UI.lpDropDown->DropDown.InsertString(
					ulDropNum,
					szAnsiName);
			}
			delete[] szAnsiName;
#endif
		}
	}

	if (m_lpMAPIProp)
	{
		if (m_lpControls[5].UI.lpDropDown)
		{
			m_lpControls[5].UI.lpDropDown->ulDropList = 2;
			m_lpControls[5].UI.lpDropDown->DropDown.InsertString(0,_T("MNID_STRING"));// STRING_OK
			m_lpControls[5].UI.lpDropDown->DropDown.InsertString(1,_T("MNID_ID"));// STRING_OK
		}

		if (m_lpControls[7].UI.lpDropDown)
		{
			m_lpControls[7].UI.lpDropDown->ulDropList = ulPropGuidArray;
			for (ulDropNum=0 ; ulDropNum < ulPropGuidArray ; ulDropNum++)
			{
				LPTSTR szGUID = GUIDToStringAndName(PropGuidArray[ulDropNum].lpGuid);
				m_lpControls[7].UI.lpDropDown->DropDown.InsertString(
					ulDropNum,
					szGUID);
				delete[] szGUID;
			}
		}
	}

	PopulateFields(NOSKIPFIELD);

	return HRES_TO_BOOL(hRes);
}

ULONG CPropertyTagEditor::GetPropertyTag()
{
	return m_ulPropTag;
}

//Select a property tag
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

//GetNamesFromIDs
void CPropertyTagEditor::OnEditAction2()
{
	if (!m_lpMAPIProp) return;

	HRESULT hRes = S_OK;

	CEditor MyData(
		this,
		IDS_CALLGETIDSFROMNAMES,
		IDS_CALLGETIDSFROMNAMESPROMPT,
		1,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
	MyData.InitCheck(0,IDS_MAPICREATE,false,false);

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		GUID	guid = {0};

		if (GetSelectedGUID(&guid))
		{
			LPSPropTagArray lpNamedPropTags = NULL;
			MAPINAMEID		NamedID;
			LPMAPINAMEID	lpNamedID = NULL;
			NamedID.lpguid = &guid;
			NamedID.ulKind = MNID_STRING;

			int iCurSel = 0;
			iCurSel = m_lpControls[5].UI.lpDropDown->DropDown.GetCurSel();
			if (iCurSel != CB_ERR)
			{
				if (0 == iCurSel) NamedID.ulKind = MNID_STRING;
				if (1 == iCurSel) NamedID.ulKind = MNID_ID;
			}
			CString szName;
			m_lpControls[6].UI.lpEdit->EditBox.GetWindowText(szName);
			if (MNID_STRING == NamedID.ulKind)
			{
#ifdef _UNICODE
				EC_H(CopyStringW(&NamedID.Kind.lpwstrName,szName,NULL));
#else
				LPWSTR	szWideName = NULL;
				EC_H(AnsiToUnicode(
					szName,
					&szWideName));
				EC_H(CopyStringW(&NamedID.Kind.lpwstrName,szWideName,NULL));
				delete[] szWideName;
#endif
			}
			else
			{

				NamedID.Kind.lID = _tcstoul((LPCTSTR)szName,NULL,16);
			}
			lpNamedID = &NamedID;

			EC_H(m_lpMAPIProp->GetIDsFromNames(
				1,
				&lpNamedID,
				MyData.GetCheck(0)?MAPI_CREATE:0,
				&lpNamedPropTags));
			if (MNID_STRING == NamedID.ulKind)
			{
				MAPIFreeBuffer(NamedID.Kind.lpwstrName);
			}

			if (lpNamedPropTags)
			{
				ULONG ulPropType = PT_NULL;
				GetSelectedPropType(&ulPropType);

				m_ulPropTag = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[0],ulPropType);
				MAPIFreeBuffer(lpNamedPropTags);
			}
			PopulateFields(NOSKIPFIELD);
		}
	}
}

BOOL CPropertyTagEditor::GetSelectedGUID(LPGUID lpSelectedGUID)
{
	if (!lpSelectedGUID) return false;
	if (!IsValidDropDown(7)) return false;

	LPCGUID lpGUID = NULL;
	int iCurSel = 0;
	iCurSel = m_lpControls[7].UI.lpDropDown->DropDown.GetCurSel();
	if (iCurSel != CB_ERR)
	{
		lpGUID = PropGuidArray[iCurSel].lpGuid;
	}
	else
	{
		//no match - need to do a lookup
		CString szText;
		GUID	guid = {0};
		m_lpControls[7].UI.lpDropDown->DropDown.GetWindowText(szText);
		//try the GUID like PS_* first
		GUIDNameToGUID((LPCTSTR) szText,&lpGUID);
		if (!lpGUID)//no match - try it like a guid {}
		{
			HRESULT hRes = S_OK;
			WC_H(StringToGUID((LPCTSTR) szText,&guid));
			lpGUID = &guid;
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
	if (!IsValidDropDown(2)) return false;

	HRESULT hRes = S_OK;
	CString szType;
	int iCurSel = 0;
	iCurSel = m_lpControls[2].UI.lpDropDown->DropDown.GetCurSel();
	if (iCurSel != CB_ERR)
	{
		m_lpControls[2].UI.lpDropDown->DropDown.GetLBText(iCurSel,szType);
	}
	else
	{
		m_lpControls[2].UI.lpDropDown->DropDown.GetWindowText(szType);
	}
	LPTSTR szEnd = NULL;
	ULONG ulType = _tcstoul((LPCTSTR) szType,&szEnd,16);

	if (*szEnd != NULL) // If we didn't consume the whole string, try a lookup
	{
		EC_H(PropTypeNameToPropType((LPCTSTR) szType,&ulType));
	}

	*ulPropType = CHANGE_PROP_TYPE(m_ulPropTag,ulType);
	return true;
}

ULONG CPropertyTagEditor::HandleChange(UINT nID)
{
	HRESULT hRes = S_OK;
	if (!m_lpControls) return (ULONG) -1;
	ULONG i = CEditor::HandleChange(nID);

	if ((ULONG) -1 == i) return (ULONG) -1;

	switch (i)
	{
	case(0)://Prop tag changed
		{
			CString szTag;
			m_lpControls[i].UI.lpEdit->EditBox.GetWindowText(szTag);
			LPTSTR szEnd = NULL;
			ULONG ulTag = _tcstoul((LPCTSTR) szTag,&szEnd,16);

			if (*szEnd != NULL) // If we didn't consume the whole string, try a lookup
			{
				EC_H(PropNameToPropTag((LPCTSTR) szTag,&ulTag));
			}

			m_ulPropTag = ulTag;
		}
		break;
	case(1)://Prop ID changed
		{
			CString szID;
			m_lpControls[i].UI.lpEdit->EditBox.GetWindowText(szID);
			ULONG ulID = _tcstoul((LPCTSTR) szID,NULL,16);

			m_ulPropTag = PROP_TAG(PROP_TYPE(m_ulPropTag),ulID);
		}
		break;
	case(2)://Prop Type changed
		{
			ULONG ulType = PT_NULL;
			GetSelectedPropType(&ulType);

			m_ulPropTag = CHANGE_PROP_TYPE(m_ulPropTag,ulType);
		}
		break;
	default:
		return i;
		break;
	}

	PopulateFields(i);

	return i;
}

//Fill out the fields in the form
//Don't touch the field passed in ulSkipField
//Pass NOSKIPFIELD to fill out all fields
void CPropertyTagEditor::PopulateFields(ULONG ulSkipField)
{
	HRESULT hRes = S_OK;

	CString PropType;
	LPTSTR szExactMatch = NULL;
	LPTSTR szPartialMatch = NULL;

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
		NULL,
		NULL,
		NULL);

	if (0 != ulSkipField) SetHex(0,m_ulPropTag);
	if (1 != ulSkipField) SetStringf(1,_T("0x%04X"),PROP_ID(m_ulPropTag));// STRING_OK
	if (2 != ulSkipField) SetDropDown(2,PropType);
	if (3 != ulSkipField)
	{
		SetStringf(3,_T("%s (%s)"),szExactMatch?szExactMatch:_T(""),szPartialMatch?szPartialMatch:_T(""));// STRING_OK
	}
	if (4 != ulSkipField) SetString(4,(LPCTSTR) TypeToString(m_ulPropTag));

	if (m_lpMAPIProp)
	{
		//do a named property lookup and fill out fields
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
				if (5 != ulSkipField) SetDropDown(5,_T("MNID_STRING"));// STRING_OK
				if (6 != ulSkipField) SetStringW(6,lppPropNames[0]->Kind.lpwstrName);
			}
			else if (MNID_ID == lppPropNames[0]->ulKind)
			{
				if (5 != ulSkipField) SetDropDown(5,_T("MNID_ID"));// STRING_OK
				if (6 != ulSkipField) SetHex(6,lppPropNames[0]->Kind.lID);
			}
			else
			{
				if (6 != ulSkipField) SetString(6,NULL);
			}
			if (7 != ulSkipField)
			{
				LPTSTR szGUID = GUIDToString(lppPropNames[0]->lpguid);
				SetDropDown(7,szGUID);
				delete[] szGUID;
			}
		}
		else
		{
			if (5 != ulSkipField) SetDropDown(5,NULL);
			if (6 != ulSkipField) SetString(6,NULL);
			if (7 != ulSkipField) SetDropDown(7,NULL);
		}
		MAPIFreeBuffer(lppPropNames);
	}

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

BEGIN_MESSAGE_MAP(CPropertySelector, CEditor)
//{{AFX_MSG_MAP(CPropertySelector)
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BOOL CPropertySelector::OnInitDialog()
{
	HRESULT hRes = S_OK;

	EC_B(CEditor::OnInitDialog());

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
			lpData = m_lpControls[0].UI.lpList->List.InsertRow(ulCurRow,PropTagArray[i].lpszName);
#else
			LPSTR szAnsiName = NULL;
			EC_H(UnicodeToAnsi(PropTagArray[i].lpszName,&szAnsiName));
			if (SUCCEEDED(hRes))
			{
				lpData = m_lpControls[0].UI.lpList->List.InsertRow(ulCurRow,szAnsiName);
			}
			delete[] szAnsiName;
#endif

			if (lpData)
			{
				lpData->ulSortDataType = SORTLIST_PROP;
				lpData->data.Prop.ulPropTag = PropTagArray[i].ulValue;
				lpData->bItemFullyLoaded = true;
			}

			szTmp.Format(_T("0x%08X"),PropTagArray[i].ulValue);// STRING_OK
			WC_B(m_lpControls[0].UI.lpList->List.SetItemText(ulCurRow,1,(LPCTSTR)szTmp));
			WC_B(m_lpControls[0].UI.lpList->List.SetItemText(ulCurRow,2,TypeToString(PropTagArray[i].ulValue)));
			ulCurRow++;
		}
		// Initial sort is by property tag
		m_lpControls[0].UI.lpList->List.SortColumn(0);
		m_lpControls[0].UI.lpList->List.AutoSizeColumns();
	}

	return HRES_TO_BOOL(hRes);
}

void CPropertySelector::OnOK()
{
	if (IsValidList(0))
	{
		int	iItem = NULL;

		iItem = m_lpControls[0].UI.lpList->List.GetNextItem(-1,LVNI_FOCUSED | LVNI_SELECTED);

		if (-1 != iItem)
		{
			SortListData* lpListData = ((SortListData*)m_lpControls[0].UI.lpList->List.GetItemData(iItem));
			if (lpListData)
				m_ulPropTag = lpListData->data.Prop.ulPropTag;
		}
	}
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
