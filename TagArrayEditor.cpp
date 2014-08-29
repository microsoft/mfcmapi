// TagArrayEditor.cpp : implementation file
//

#include "stdafx.h"
#include "TagArrayEditor.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"
#include "MAPIFunctions.h"
#include "PropertyTagEditor.h"

static TCHAR* CLASS = _T("CTagArrayEditor");

CTagArrayEditor::CTagArrayEditor(
	_In_opt_ CWnd* pParentWnd,
	UINT uidTitle,
	UINT uidPrompt,
	_In_opt_ LPMAPITABLE lpContentsTable,
	_In_opt_ LPSPropTagArray lpTagArray,
	bool bIsAB,
	_In_opt_ LPMAPIPROP lpMAPIProp) :
	CEditor(pParentWnd, uidTitle, uidPrompt, 0, CEDITOR_BUTTON_OK | (lpContentsTable ? (CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_ACTION2) : 0) | CEDITOR_BUTTON_CANCEL, IDS_QUERYCOLUMNS, IDS_FLAGS, NULL)
{
	TRACE_CONSTRUCTOR(CLASS);

	m_lpContentsTable = lpContentsTable;
	if (m_lpContentsTable) m_lpContentsTable->AddRef();
	m_ulSetColumnsFlags = TBL_BATCH;
	m_lpTagArray = lpTagArray;
	m_bIsAB = bIsAB;
	m_lpOutputTagArray = NULL;
	m_lpMAPIProp = lpMAPIProp;
	if (m_lpMAPIProp) m_lpMAPIProp->AddRef();

	CreateControls(1);
	InitPane(0, CreateListPane(IDS_PROPTAGARRAY, false, false, this));
} // CTagArrayEditor::CTagArrayEditor

CTagArrayEditor::~CTagArrayEditor()
{
	TRACE_DESTRUCTOR(CLASS);
	MAPIFreeBuffer(m_lpOutputTagArray);
	if (m_lpMAPIProp) m_lpMAPIProp->Release();
} // CTagArrayEditor::~CTagArrayEditor

// Used to call functions which need to be called AFTER controls are created
BOOL CTagArrayEditor::OnInitDialog()
{
	BOOL bRet = CEditor::OnInitDialog();

	ReadTagArrayToList(0);

	UpdateListButtons();
	return bRet;
} // CTagArrayEditor::OnInitDialog

void CTagArrayEditor::OnOK()
{
	WriteListToTagArray(0);
	if (m_lpOutputTagArray && m_lpContentsTable)
	{
		// Apply lpFinalTagArray through SetColumns
		HRESULT hRes = S_OK;
		EC_MAPI(m_lpContentsTable->SetColumns(
			m_lpOutputTagArray,
			m_ulSetColumnsFlags)); // Flags
	}

	CMyDialog::OnOK(); // don't need to call CEditor::OnOK
} // CTagArrayEditor::OnOK

_Check_return_ bool CTagArrayEditor::DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData)
{
	if (!IsValidList(ulListNum)) return false;
	if (!lpData) return false;

	HRESULT	hRes = S_OK;
	ULONG	ulOrigPropTag = lpData->data.Tag.ulPropTag;
	ULONG	ulNewPropTag = ulOrigPropTag;

	CPropertyTagEditor MyPropertyTag(
		NULL, // title
		NULL, // prompt
		ulOrigPropTag,
		m_bIsAB,
		m_lpMAPIProp,
		this);

	WC_H(MyPropertyTag.DisplayDialog());
	if (S_OK != hRes) return false;
	ulNewPropTag = MyPropertyTag.GetPropertyTag();

	if (ulNewPropTag != ulOrigPropTag)
	{
		CString szTmp;
		lpData->data.Tag.ulPropTag = ulNewPropTag;

		szTmp.Format(_T("0x%08X"), ulNewPropTag); // STRING_OK
		SetListString(ulListNum, iItem, 1, szTmp);

		LPTSTR szExactMatch = NULL;
		LPTSTR szPartialMatch = NULL;
		LPTSTR szNamedPropName = NULL;
		LPTSTR szNamedPropGUID = NULL;

		InterpretProp(
			ulNewPropTag,
			m_lpMAPIProp,
			NULL,
			NULL,
			m_bIsAB,
			&szNamedPropName, // Built from lpProp & lpMAPIProp
			&szNamedPropGUID, // Built from lpProp & lpMAPIProp
			NULL);

		EC_H(PropTagToPropName(ulNewPropTag, m_bIsAB, &szExactMatch, &szPartialMatch));

		SetListString(ulListNum, iItem, 2, szExactMatch);
		SetListString(ulListNum, iItem, 3, szPartialMatch);
		SetListString(ulListNum, iItem, 4, TypeToString(ulNewPropTag));
		SetListString(ulListNum, iItem, 5, szNamedPropName);
		SetListString(ulListNum, iItem, 6, szNamedPropGUID);

		delete[] szPartialMatch;
		delete[] szExactMatch;
		FreeNameIDStrings(szNamedPropName, szNamedPropGUID, NULL);

		ResizeList(ulListNum, false);
		return true;
	}
	return false;
} // CTagArrayEditor::DoListEdit

void CTagArrayEditor::ReadTagArrayToList(ULONG ulListNum)
{
	if (!IsValidList(ulListNum)) return;

	ClearList(ulListNum);

	InsertColumn(ulListNum, 0, IDS_SHARP);
	InsertColumn(ulListNum, 1, IDS_TAG);
	InsertColumn(ulListNum, 2, IDS_PROPERTYNAMES);
	InsertColumn(ulListNum, 3, IDS_OTHERNAMES);
	InsertColumn(ulListNum, 4, IDS_TYPE);
	InsertColumn(ulListNum, 5, IDS_NAMEDPROPNAME);
	InsertColumn(ulListNum, 6, IDS_NAMEDPROPGUID);

	if (m_lpTagArray)
	{
		CString szTmp;
		ULONG iTagCount = 0;
		ULONG cValues = m_lpTagArray->cValues;

		for (iTagCount = 0; iTagCount < cValues; iTagCount++)
		{
			ULONG ulPropTag = m_lpTagArray->aulPropTag[iTagCount];
			SortListData* lpData = NULL;
			szTmp.Format(_T("%u"), iTagCount); // STRING_OK
			lpData = InsertListRow(ulListNum, iTagCount, szTmp);
			if (lpData)
			{
				lpData->ulSortDataType = SORTLIST_TAGARRAY;
				lpData->data.Tag.ulPropTag = ulPropTag;
				lpData->bItemFullyLoaded = true;
			}

			CString PropTag;
			LPTSTR szExactMatch = NULL;
			LPTSTR szPartialMatch = NULL;
			LPTSTR szNamedPropName = NULL;
			LPTSTR szNamedPropGUID = NULL;

			InterpretProp(
				ulPropTag,
				m_lpMAPIProp,
				NULL,
				NULL,
				m_bIsAB,
				&szNamedPropName, // Built from lpProp & lpMAPIProp
				&szNamedPropGUID, // Built from lpProp & lpMAPIProp
				NULL);

			PropTag.Format(_T("0x%08X"), ulPropTag);
			SetListString(ulListNum, iTagCount, 1, PropTag);

			HRESULT hRes = S_OK;
			EC_H(PropTagToPropName(ulPropTag, m_bIsAB, &szExactMatch, &szPartialMatch));
			if (SUCCEEDED(hRes))
			{
				SetListString(ulListNum, iTagCount, 2, szExactMatch);
				SetListString(ulListNum, iTagCount, 3, szPartialMatch);
			}

			SetListString(ulListNum, iTagCount, 4, TypeToString(ulPropTag));
			SetListString(ulListNum, iTagCount, 5, szNamedPropName);
			SetListString(ulListNum, iTagCount, 6, szNamedPropGUID);

			delete[] szPartialMatch;
			delete[] szExactMatch;
			FreeNameIDStrings(szNamedPropName, szNamedPropGUID, NULL);
		}
	}

	ResizeList(ulListNum, false);
} // CTagArrayEditor::ReadTagArrayToList

void CTagArrayEditor::WriteListToTagArray(ULONG ulListNum)
{
	if (!IsValidList(ulListNum)) return;

	// If we're not dirty, don't write
	if (!IsDirty(ulListNum)) return;

	HRESULT hRes = S_OK;
	ULONG ulListCount = GetListCount(ulListNum);
	EC_H(MAPIAllocateBuffer(
		CbNewSPropTagArray(ulListCount),
		(LPVOID*)&m_lpOutputTagArray));
	if (m_lpOutputTagArray)
	{
		m_lpOutputTagArray->cValues = ulListCount;

		ULONG iTagCount = 0;
		for (iTagCount = 0; iTagCount < m_lpOutputTagArray->cValues; iTagCount++)
		{
			SortListData* lpData = GetListRowData(ulListNum, iTagCount);
			if (lpData)
				m_lpOutputTagArray->aulPropTag[iTagCount] = lpData->data.Tag.ulPropTag;
		}

	}
} // CTagArrayEditor::WriteListToTagArray

_Check_return_ LPSPropTagArray CTagArrayEditor::DetachModifiedTagArray()
{
	LPSPropTagArray lpRetArray = m_lpOutputTagArray;
	m_lpOutputTagArray = NULL;
	return lpRetArray;
} // CTagArrayEditor::DetachModifiedTagArray

// QueryColumns flags
void CTagArrayEditor::OnEditAction1()
{
	if (!m_lpContentsTable) return;
	HRESULT hRes = S_OK;
	ULONG ulQueryColumnFlags = NULL;

	CEditor MyData(
		this,
		IDS_QUERYCOLUMNS,
		IDS_QUERYCOLUMNSPROMPT,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CreateSingleLinePane(IDS_QUERYCOLUMNFLAGS, NULL, false));
	MyData.SetHex(0, ulQueryColumnFlags);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		ulQueryColumnFlags = MyData.GetHex(0);
		LPSPropTagArray lpTagArray = NULL;

		EC_MAPI(m_lpContentsTable->QueryColumns(
			ulQueryColumnFlags,
			&lpTagArray));

		if (SUCCEEDED(hRes))
		{
			MAPIFreeBuffer(m_lpTagArray);
			m_lpTagArray = lpTagArray;

			ReadTagArrayToList(0);
			UpdateListButtons();
		}
	}
}

// SetColumns flags
void CTagArrayEditor::OnEditAction2()
{
	if (!m_lpContentsTable) return;
	HRESULT hRes = S_OK;

	CEditor MyData(
		this,
		IDS_SETCOLUMNS,
		IDS_SETCOLUMNSPROMPT,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CreateSingleLinePane(IDS_SETCOLUMNFLAGS, NULL, false));
	MyData.SetHex(0, m_ulSetColumnsFlags);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		m_ulSetColumnsFlags = MyData.GetHex(0);
	}
}