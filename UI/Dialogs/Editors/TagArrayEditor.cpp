#include <StdAfx.h>
#include <UI/Dialogs/Editors/TagArrayEditor.h>
#include <UI/Dialogs/Editors/PropertyTagEditor.h>
#include <core/sortlistdata/propListData.h>
#include <core/mapi/cache/namedPropCache.h>
#include <core/mapi/mapiMemory.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/interpret/proptags.h>
#include <core/interpret/proptype.h>
#include <core/addin/mfcmapi.h>
#include <core/mapi/mapiFunctions.h>

namespace dialog::editor
{
	static std::wstring CLASS = L"CTagArrayEditor";

	CTagArrayEditor::CTagArrayEditor(
		_In_opt_ CWnd* pParentWnd,
		UINT uidTitle,
		UINT uidPrompt,
		_In_opt_ LPMAPITABLE lpContentsTable,
		_In_opt_ LPSPropTagArray lpTagArray,
		bool bIsAB,
		_In_opt_ LPMAPIPROP lpMAPIProp)
		: CEditor(
			  pParentWnd,
			  uidTitle,
			  uidPrompt,
			  CEDITOR_BUTTON_OK | (lpContentsTable ? CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_ACTION2 : 0) |
				  CEDITOR_BUTTON_CANCEL,
			  IDS_QUERYCOLUMNS,
			  IDS_FLAGS,
			  NULL)
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

		AddPane(viewpane::ListPane::Create(0, IDS_PROPTAGARRAY, false, false, ListEditCallBack(this)));
		SetListID(0);
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
			EC_MAPI_S(m_lpContentsTable->SetColumns(m_lpOutputTagArray,
													m_ulSetColumnsFlags)); // Flags
		}

		CMyDialog::OnOK(); // don't need to call CEditor::OnOK
	}

	_Check_return_ bool CTagArrayEditor::DoListEdit(ULONG ulListNum, int iItem, _In_ sortlistdata::sortListData* lpData)
	{
		if (!lpData) return false;
		if (!lpData->cast<sortlistdata::propListData>())
		{
			sortlistdata::propListData::init(lpData, 0);
		}

		const auto prop = lpData->cast<sortlistdata::propListData>();
		if (!prop) return false;

		const auto ulOrigPropTag = prop->m_ulPropTag;

		CPropertyTagEditor MyPropertyTag(
			NULL, // title
			NULL, // prompt
			ulOrigPropTag,
			m_bIsAB,
			m_lpMAPIProp,
			this);

		if (!MyPropertyTag.DisplayDialog()) return false;
		const auto ulNewPropTag = MyPropertyTag.GetPropertyTag();

		if (ulNewPropTag != ulOrigPropTag)
		{
			prop->m_ulPropTag = ulNewPropTag;

			const auto namePropNames = cache::NameIDToStrings(ulNewPropTag, m_lpMAPIProp, nullptr, nullptr, m_bIsAB);

			const auto propTagNames = proptags::PropTagToPropName(ulNewPropTag, m_bIsAB);

			SetListString(ulListNum, iItem, 1, strings::format(L"0x%08X", ulNewPropTag));
			SetListString(ulListNum, iItem, 2, propTagNames.bestGuess);
			SetListString(ulListNum, iItem, 3, propTagNames.otherMatches);
			SetListString(ulListNum, iItem, 4, proptype::TypeToString(ulNewPropTag));
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
				const auto ulPropTag = mapi::getTag(lpTagArray, iTagCount);
				auto lpData = InsertListRow(ulListNum, iTagCount, std::to_wstring(iTagCount));
				if (lpData)
				{
					sortlistdata::propListData::init(lpData, ulPropTag);
				}

				const auto namePropNames = cache::NameIDToStrings(ulPropTag, m_lpMAPIProp, nullptr, nullptr, m_bIsAB);

				const auto propTagNames = proptags::PropTagToPropName(ulPropTag, m_bIsAB);

				SetListString(ulListNum, iTagCount, 1, strings::format(L"0x%08X", ulPropTag));
				SetListString(ulListNum, iTagCount, 2, propTagNames.bestGuess);
				SetListString(ulListNum, iTagCount, 3, propTagNames.otherMatches);
				SetListString(ulListNum, iTagCount, 4, proptype::TypeToString(ulPropTag));
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

		const auto ulListCount = GetListCount(ulListNum);
		m_lpOutputTagArray = mapi::allocate<LPSPropTagArray>(CbNewSPropTagArray(ulListCount));
		if (m_lpOutputTagArray)
		{
			m_lpOutputTagArray->cValues = ulListCount;

			for (ULONG iTagCount = 0; iTagCount < m_lpOutputTagArray->cValues; iTagCount++)
			{
				const auto lpData = GetListRowData(ulListNum, iTagCount);
				if (lpData)
				{
					const auto prop = lpData->cast<sortlistdata::propListData>();
					if (prop)
					{
						mapi::setTag(m_lpOutputTagArray, iTagCount) = prop->m_ulPropTag;
					}
				}
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
		ULONG ulQueryColumnFlags = NULL;

		CEditor MyData(this, IDS_QUERYCOLUMNS, IDS_QUERYCOLUMNSPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_QUERYCOLUMNFLAGS, false));
		MyData.SetHex(0, ulQueryColumnFlags);

		if (!MyData.DisplayDialog()) return;

		ulQueryColumnFlags = MyData.GetHex(0);
		LPSPropTagArray lpTagArray = nullptr;

		const auto hRes = EC_MAPI(m_lpContentsTable->QueryColumns(ulQueryColumnFlags, &lpTagArray));

		if (SUCCEEDED(hRes))
		{
			ReadTagArrayToList(0, lpTagArray);
			UpdateButtons();
		}

		MAPIFreeBuffer(lpTagArray);
	}

	// SetColumns flags
	void CTagArrayEditor::OnEditAction2()
	{
		if (!m_lpContentsTable) return;

		CEditor MyData(this, IDS_SETCOLUMNS, IDS_SETCOLUMNSPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_SETCOLUMNFLAGS, false));
		MyData.SetHex(0, m_ulSetColumnsFlags);

		if (MyData.DisplayDialog())
		{
			m_ulSetColumnsFlags = MyData.GetHex(0);
		}
	}
} // namespace dialog::editor