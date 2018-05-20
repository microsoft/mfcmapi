#include "StdAfx.h"
#include "TagArrayEditor.h"
#include <Interpret/InterpretProp.h>
#include "PropertyTagEditor.h"
#include <UI/Controls/SortList/PropListData.h>
#include <MAPI/NamedPropCache.h>

static std::wstring CLASS = L"CTagArrayEditor";

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
	m_lpSourceTagArray = lpTagArray;
	m_bIsAB = bIsAB;
	m_lpOutputTagArray = nullptr;
	m_lpMAPIProp = lpMAPIProp;
	if (m_lpMAPIProp) m_lpMAPIProp->AddRef();

	InitPane(0, ListPane::Create(IDS_PROPTAGARRAY, false, false, ListEditCallBack(this)));
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
	const auto bRet = CEditor::OnInitDialog();

	ReadTagArrayToList(0, m_lpSourceTagArray);

	UpdateButtons();
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
	if (!lpData) return false;
	if (!lpData->Prop())
	{
		lpData->InitializePropList(0);
	}

	auto hRes = S_OK;
	const auto ulOrigPropTag = lpData->Prop()->m_ulPropTag;

	CPropertyTagEditor MyPropertyTag(
		NULL, // title
		NULL, // prompt
		ulOrigPropTag,
		m_bIsAB,
		m_lpMAPIProp,
		this);

	WC_H(MyPropertyTag.DisplayDialog());
	if (S_OK != hRes) return false;
	const auto ulNewPropTag = MyPropertyTag.GetPropertyTag();

	if (ulNewPropTag != ulOrigPropTag)
	{
		lpData->Prop()->m_ulPropTag = ulNewPropTag;

		const auto namePropNames = NameIDToStrings(
			ulNewPropTag,
			m_lpMAPIProp,
			nullptr,
			nullptr,
			m_bIsAB);

		const auto propTagNames = PropTagToPropName(ulNewPropTag, m_bIsAB);

		SetListString(ulListNum, iItem, 1, strings::format(L"0x%08X", ulNewPropTag));
		SetListString(ulListNum, iItem, 2, propTagNames.bestGuess);
		SetListString(ulListNum, iItem, 3, propTagNames.otherMatches);
		SetListString(ulListNum, iItem, 4, TypeToString(ulNewPropTag));
		SetListString(ulListNum, iItem, 5, namePropNames.name);
		SetListString(ulListNum, iItem, 6, namePropNames.guid);

		ResizeList(ulListNum, false);
		return true;
	}
	return false;
}

void CTagArrayEditor::ReadTagArrayToList(ULONG ulListNum, LPSPropTagArray lpTagArray) const
{
	ClearList(ulListNum);

	InsertColumn(ulListNum, 0, IDS_SHARP);
	InsertColumn(ulListNum, 1, IDS_TAG);
	InsertColumn(ulListNum, 2, IDS_PROPERTYNAME);
	InsertColumn(ulListNum, 3, IDS_OTHERNAMES);
	InsertColumn(ulListNum, 4, IDS_TYPE);
	InsertColumn(ulListNum, 5, IDS_NAMEDPROPNAME);
	InsertColumn(ulListNum, 6, IDS_NAMEDPROPGUID);

	if (lpTagArray)
	{
		const auto cValues = lpTagArray->cValues;

		for (ULONG iTagCount = 0; iTagCount < cValues; iTagCount++)
		{
			const auto ulPropTag = lpTagArray->aulPropTag[iTagCount];
			auto lpData = InsertListRow(ulListNum, iTagCount, std::to_wstring(iTagCount));
			if (lpData)
			{
				lpData->InitializePropList(ulPropTag);
			}

			const auto namePropNames = NameIDToStrings(
				ulPropTag,
				m_lpMAPIProp,
				nullptr,
				nullptr,
				m_bIsAB);

			const auto propTagNames = PropTagToPropName(ulPropTag, m_bIsAB);

			SetListString(ulListNum, iTagCount, 1, strings::format(L"0x%08X", ulPropTag));
			SetListString(ulListNum, iTagCount, 2, propTagNames.bestGuess);
			SetListString(ulListNum, iTagCount, 3, propTagNames.otherMatches);
			SetListString(ulListNum, iTagCount, 4, TypeToString(ulPropTag));
			SetListString(ulListNum, iTagCount, 5, namePropNames.name);
			SetListString(ulListNum, iTagCount, 6, namePropNames.guid);
		}
	}

	ResizeList(ulListNum, false);
}

void CTagArrayEditor::WriteListToTagArray(ULONG ulListNum)
{
	// If we're not dirty, don't write
	if (!IsDirty(ulListNum)) return;

	auto hRes = S_OK;
	const auto ulListCount = GetListCount(ulListNum);
	EC_H(MAPIAllocateBuffer(
		CbNewSPropTagArray(ulListCount),
		reinterpret_cast<LPVOID*>(&m_lpOutputTagArray)));
	if (m_lpOutputTagArray)
	{
		m_lpOutputTagArray->cValues = ulListCount;

		for (ULONG iTagCount = 0; iTagCount < m_lpOutputTagArray->cValues; iTagCount++)
		{
			const auto lpData = GetListRowData(ulListNum, iTagCount);
			if (lpData && lpData->Prop())
				m_lpOutputTagArray->aulPropTag[iTagCount] = lpData->Prop()->m_ulPropTag;
		}
	}
}

_Check_return_ LPSPropTagArray CTagArrayEditor::DetachModifiedTagArray()
{
	const auto lpRetArray = m_lpOutputTagArray;
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
			ReadTagArrayToList(0, lpTagArray);
			UpdateButtons();
		}

		MAPIFreeBuffer(lpTagArray);
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