#include "stdafx.h"
#include "TagArrayEditor.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"
#include "PropertyTagEditor.h"
#include "SortList/PropListData.h"

static wstring CLASS = L"CTagArrayEditor";

CTagArrayEditor::CTagArrayEditor(
	_In_opt_ CWnd* pParentWnd,
	UINT uidTitle,
	UINT uidPrompt,
	_In_opt_ LPMAPITABLE lpContentsTable,
	_In_opt_ LPSPropTagArray lpTagArray,
	bool bIsAB,
	_In_opt_ LPMAPIPROP lpMAPIProp) :
	CEditor(pParentWnd, uidTitle, uidPrompt, CEDITOR_BUTTON_OK | (lpContentsTable ? CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_ACTION2 : 0) | CEDITOR_BUTTON_CANCEL, IDS_QUERYCOLUMNS, IDS_FLAGS, NULL)
{
	TRACE_CONSTRUCTOR(CLASS);

	m_lpContentsTable = lpContentsTable;
	if (m_lpContentsTable) m_lpContentsTable->AddRef();
	m_ulSetColumnsFlags = TBL_BATCH;
	m_lpTagArray = lpTagArray;
	m_bIsAB = bIsAB;
	m_lpOutputTagArray = nullptr;
	m_lpMAPIProp = lpMAPIProp;
	if (m_lpMAPIProp) m_lpMAPIProp->AddRef();

	InitPane(0, ListPane::Create(IDS_PROPTAGARRAY, false, false, this));
}

CTagArrayEditor::~CTagArrayEditor()
{
	TRACE_DESTRUCTOR(CLASS);
	MAPIFreeBuffer(m_lpOutputTagArray);
	if (m_lpMAPIProp) m_lpMAPIProp->Release();
}

// Used to call functions which need to be called AFTER controls are created
BOOL CTagArrayEditor::OnInitDialog()
{
	auto bRet = CEditor::OnInitDialog();

	ReadTagArrayToList(0);

	UpdateListButtons();
	return bRet;
}

void CTagArrayEditor::OnOK()
{
	WriteListToTagArray(0);
	if (m_lpOutputTagArray && m_lpContentsTable)
	{
		// Apply lpFinalTagArray through SetColumns
		auto hRes = S_OK;
		EC_MAPI(m_lpContentsTable->SetColumns(
			m_lpOutputTagArray,
			m_ulSetColumnsFlags)); // Flags
	}

	CMyDialog::OnOK(); // don't need to call CEditor::OnOK
}

_Check_return_ bool CTagArrayEditor::DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData)
{
	if (!IsValidList(ulListNum)) return false;
	if (!lpData) return false;
	if (!lpData->Prop())
	{
		lpData->InitializePropList(0);
	}

	auto hRes = S_OK;
	auto ulOrigPropTag = lpData->Prop()->m_ulPropTag;

	CPropertyTagEditor MyPropertyTag(
		NULL, // title
		NULL, // prompt
		ulOrigPropTag,
		m_bIsAB,
		m_lpMAPIProp,
		this);

	WC_H(MyPropertyTag.DisplayDialog());
	if (S_OK != hRes) return false;
	auto ulNewPropTag = MyPropertyTag.GetPropertyTag();

	if (ulNewPropTag != ulOrigPropTag)
	{
		lpData->Prop()->m_ulPropTag = ulNewPropTag;

		wstring szExactMatch;
		wstring szPartialMatch;
		wstring szNamedPropName;
		wstring szNamedPropGUID;
		wstring szNamedPropDASL;

		NameIDToStrings(
			ulNewPropTag,
			m_lpMAPIProp,
			nullptr,
			nullptr,
			m_bIsAB,
			szNamedPropName, // Built from lpProp & lpMAPIProp
			szNamedPropGUID, // Built from lpProp & lpMAPIProp
			szNamedPropDASL);

		PropTagToPropName(ulNewPropTag, m_bIsAB, szExactMatch, szPartialMatch);

		SetListString(ulListNum, iItem, 1, format(L"0x%08X", ulNewPropTag));
		SetListString(ulListNum, iItem, 2, szExactMatch);
		SetListString(ulListNum, iItem, 3, szPartialMatch);
		SetListString(ulListNum, iItem, 4, TypeToString(ulNewPropTag));
		SetListString(ulListNum, iItem, 5, szNamedPropName);
		SetListString(ulListNum, iItem, 6, szNamedPropGUID);

		ResizeList(ulListNum, false);
		return true;
	}
	return false;
}

void CTagArrayEditor::ReadTagArrayToList(ULONG ulListNum) const
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
		auto cValues = m_lpTagArray->cValues;

		for (ULONG iTagCount = 0; iTagCount < cValues; iTagCount++)
		{
			auto ulPropTag = m_lpTagArray->aulPropTag[iTagCount];
			auto lpData = InsertListRow(ulListNum, iTagCount, format(L"%u", iTagCount)); // STRING_OK
			if (lpData)
			{
				lpData->InitializePropList(ulPropTag);
			}

			wstring szExactMatch;
			wstring szPartialMatch;
			wstring szNamedPropName;
			wstring szNamedPropGUID;
			wstring szNamedPropDASL;

			NameIDToStrings(
				ulPropTag,
				m_lpMAPIProp,
				nullptr,
				nullptr,
				m_bIsAB,
				szNamedPropName, // Built from lpProp & lpMAPIProp
				szNamedPropGUID, // Built from lpProp & lpMAPIProp
				szNamedPropDASL);

			PropTagToPropName(ulPropTag, m_bIsAB, szExactMatch, szPartialMatch);

			SetListString(ulListNum, iTagCount, 1, format(L"0x%08X", ulPropTag));
			SetListString(ulListNum, iTagCount, 2, szExactMatch);
			SetListString(ulListNum, iTagCount, 3, szPartialMatch);
			SetListString(ulListNum, iTagCount, 4, TypeToString(ulPropTag));
			SetListString(ulListNum, iTagCount, 5, szNamedPropName);
			SetListString(ulListNum, iTagCount, 6, szNamedPropGUID);
		}
	}

	ResizeList(ulListNum, false);
}

void CTagArrayEditor::WriteListToTagArray(ULONG ulListNum)
{
	if (!IsValidList(ulListNum)) return;

	// If we're not dirty, don't write
	if (!IsDirty(ulListNum)) return;

	auto hRes = S_OK;
	auto ulListCount = GetListCount(ulListNum);
	EC_H(MAPIAllocateBuffer(
		CbNewSPropTagArray(ulListCount),
		reinterpret_cast<LPVOID*>(&m_lpOutputTagArray)));
	if (m_lpOutputTagArray)
	{
		m_lpOutputTagArray->cValues = ulListCount;

		for (ULONG iTagCount = 0; iTagCount < m_lpOutputTagArray->cValues; iTagCount++)
		{
			auto lpData = GetListRowData(ulListNum, iTagCount);
			if (lpData && lpData->Prop())
				m_lpOutputTagArray->aulPropTag[iTagCount] = lpData->Prop()->m_ulPropTag;
		}
	}
}

_Check_return_ LPSPropTagArray CTagArrayEditor::DetachModifiedTagArray()
{
	auto lpRetArray = m_lpOutputTagArray;
	m_lpOutputTagArray = nullptr;
	return lpRetArray;
}

// QueryColumns flags
void CTagArrayEditor::OnEditAction1()
{
	if (!m_lpContentsTable) return;
	auto hRes = S_OK;
	ULONG ulQueryColumnFlags = NULL;

	CEditor MyData(
		this,
		IDS_QUERYCOLUMNS,
		IDS_QUERYCOLUMNSPROMPT,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, TextPane::CreateSingleLinePane(IDS_QUERYCOLUMNFLAGS, false));
	MyData.SetHex(0, ulQueryColumnFlags);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		ulQueryColumnFlags = MyData.GetHex(0);
		LPSPropTagArray lpTagArray = nullptr;

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
	auto hRes = S_OK;

	CEditor MyData(
		this,
		IDS_SETCOLUMNS,
		IDS_SETCOLUMNSPROMPT,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, TextPane::CreateSingleLinePane(IDS_SETCOLUMNFLAGS, false));
	MyData.SetHex(0, m_ulSetColumnsFlags);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		m_ulSetColumnsFlags = MyData.GetHex(0);
	}
}