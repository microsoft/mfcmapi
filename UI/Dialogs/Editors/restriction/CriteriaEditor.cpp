#include <StdAfx.h>
#include <UI/Dialogs/Editors/restriction/CriteriaEditor.h>
#include <UI/Dialogs/Editors/restriction/RestrictEditor.h>
#include <core/mapi/extraPropTags.h>
#include <core/sortlistdata/binaryData.h>
#include <core/mapi/mapiMemory.h>
#include <core/utility/output.h>
#include <core/interpret/flags.h>
#include <core/addin/mfcmapi.h>
#include <core/property/parseProperty.h>

namespace dialog::editor
{
	// Note that no alloc parent is passed in to CriteriaEditor. So we're completely responsible for freeing any memory we allocate.
	// If we return (detach) memory to a caller, they must MAPIFreeBuffer
	static std::wstring CRITERIACLASS = L"CriteriaEditor"; // STRING_OK
#define LISTNUM 4
	CriteriaEditor::CriteriaEditor(
		_In_ CWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		_In_ const _SRestriction* lpRes,
		_In_ LPENTRYLIST lpEntryList,
		ULONG ulSearchState)
		: CEditor(
			  pParentWnd,
			  IDS_CRITERIAEDITOR,
			  IDS_CRITERIAEDITORPROMPT,
			  CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL,
			  IDS_ACTIONEDITRES,
			  NULL,
			  NULL)
	{
		TRACE_CONSTRUCTOR(CRITERIACLASS);
		m_lpMapiObjects = lpMapiObjects;

		m_lpSourceRes = lpRes;
		m_lpNewRes = nullptr;
		m_lpSourceEntryList = lpEntryList;

		m_lpNewEntryList = mapi::allocate<SBinaryArray*>(sizeof(SBinaryArray));

		m_ulNewSearchFlags = NULL;

		SetPromptPostFix(flags::AllFlagsToString(flagSearchFlag, true));
		AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_SEARCHSTATE, true));
		SetHex(0, ulSearchState);
		const auto szFlags = flags::InterpretFlags(flagSearchState, ulSearchState);
		AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_SEARCHSTATE, szFlags, true));
		AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_SEARCHFLAGS, false));
		SetHex(2, 0);
		AddPane(viewpane::TextPane::CreateSingleLinePane(3, IDS_SEARCHFLAGS, true));
		AddPane(viewpane::ListPane::CreateCollapsibleListPane(4, IDS_EIDLIST, false, false, ListEditCallBack(this)));
		SetListID(4);
		AddPane(viewpane::TextPane::CreateMultiLinePane(
			5, IDS_RESTRICTIONTEXT, property::RestrictionToString(m_lpSourceRes, nullptr), true));
	}

	CriteriaEditor::~CriteriaEditor()
	{
		// If these structures weren't detached, we need to free them
		MAPIFreeBuffer(m_lpNewEntryList);
		MAPIFreeBuffer(m_lpNewRes);
	}

	// Used to call functions which need to be called AFTER controls are created
	BOOL CriteriaEditor::OnInitDialog()
	{
		const auto bRet = CEditor::OnInitDialog();

		InitListFromEntryList(LISTNUM, m_lpSourceEntryList);

		UpdateButtons();

		return bRet;
	}

	_Check_return_ const _SRestriction* CriteriaEditor::GetSourceRes() const noexcept
	{
		if (m_lpNewRes) return m_lpNewRes;
		return m_lpSourceRes;
	}

	_Check_return_ ULONG CriteriaEditor::HandleChange(UINT nID)
	{
		const auto paneID = CEditor::HandleChange(nID);

		if (paneID == 2)
		{
			SetStringW(3, flags::InterpretFlags(flagSearchFlag, GetHex(paneID)));
		}

		return paneID;
	}

	// Whoever gets this MUST MAPIFreeBuffer
	_Check_return_ LPSRestriction CriteriaEditor::DetachModifiedSRestriction() noexcept
	{
		const auto lpRet = m_lpNewRes;
		m_lpNewRes = nullptr;
		return lpRet;
	}

	// Whoever gets this MUST MAPIFreeBuffer
	_Check_return_ LPENTRYLIST CriteriaEditor::DetachModifiedEntryList() noexcept
	{
		const auto lpRet = m_lpNewEntryList;
		m_lpNewEntryList = nullptr;
		return lpRet;
	}

	_Check_return_ ULONG CriteriaEditor::GetSearchFlags() const noexcept { return m_ulNewSearchFlags; }

	void CriteriaEditor::InitListFromEntryList(ULONG ulListNum, _In_ const SBinaryArray* lpEntryList) const
	{
		ClearList(ulListNum);

		InsertColumn(ulListNum, 0, IDS_SHARP);
		InsertColumn(ulListNum, 1, IDS_CB);
		InsertColumn(ulListNum, 2, IDS_BINARY);
		InsertColumn(ulListNum, 3, IDS_TEXTVIEW);

		if (lpEntryList)
		{
			for (ULONG iRow = 0; iRow < lpEntryList->cValues; iRow++)
			{
				auto lpData = InsertListRow(ulListNum, iRow, std::to_wstring(iRow));
				if (lpData)
				{
					sortlistdata::binaryData::init(lpData, &lpEntryList->lpbin[iRow]);
				}

				SetListString(ulListNum, iRow, 1, std::to_wstring(lpEntryList->lpbin[iRow].cb));
				SetListString(ulListNum, iRow, 2, strings::BinToHexString(&lpEntryList->lpbin[iRow], false));
				SetListString(ulListNum, iRow, 3, strings::BinToTextString(&lpEntryList->lpbin[iRow], true));
				if (lpData) lpData->setFullyLoaded(true);
			}
		}

		ResizeList(ulListNum, false);
	}

	void CriteriaEditor::OnEditAction1()
	{
		const auto lpSourceRes = GetSourceRes();

		RestrictEditor MyResEditor(this, m_lpMapiObjects, nullptr,
									lpSourceRes); // pass source res into editor
		if (MyResEditor.DisplayDialog())
		{
			const auto lpModRes = MyResEditor.DetachModifiedSRestriction();
			if (lpModRes)
			{
				// We didn't pass an alloc parent to RestrictEditor, so we must free what came back
				MAPIFreeBuffer(m_lpNewRes);
				m_lpNewRes = lpModRes;
				SetStringW(5, property::RestrictionToString(m_lpNewRes, nullptr));
			}
		}
	}

	_Check_return_ bool CriteriaEditor::DoListEdit(ULONG ulListNum, int iItem, _In_ sortlistdata::sortListData* lpData)
	{
		if (!lpData) return false;

		if (!lpData->cast<sortlistdata::binaryData>())
		{
			sortlistdata::binaryData::init(lpData, nullptr);
		}

		const auto binary = lpData->cast<sortlistdata::binaryData>();
		if (!binary) return false;

		CEditor BinEdit(this, IDS_EIDEDITOR, IDS_EIDEDITORPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		auto lpSourcebin = binary->getCurrentBin();

		BinEdit.AddPane(
			viewpane::TextPane::CreateSingleLinePane(0, IDS_EID, strings::BinToHexString(lpSourcebin, false), false));

		if (BinEdit.DisplayDialog())
		{
			auto bin = strings::HexStringToBin(BinEdit.GetStringW(0));
			auto newBin = SBinary{};
			newBin.lpb = mapi::ByteVectorToMAPI(bin, m_lpNewEntryList);
			if (newBin.lpb)
			{
				newBin.cb = static_cast<ULONG>(bin.size());
				binary->setCurrentBin(newBin);
				const auto szTmp = std::to_wstring(newBin.cb);
				SetListString(ulListNum, iItem, 1, szTmp);
				SetListString(ulListNum, iItem, 2, strings::BinToHexString(&newBin, false));
				SetListString(ulListNum, iItem, 3, strings::BinToTextString(&newBin, true));
				return true;
			}
		}

		return false;
	}

	void CriteriaEditor::OnOK()
	{
		CMyDialog::OnOK(); // don't need to call CEditor::OnOK
		const auto ulValues = GetListCount(LISTNUM);

		if (m_lpNewEntryList && ulValues < ULONG_MAX / sizeof(SBinary))
		{
			m_lpNewEntryList->cValues = ulValues;
			m_lpNewEntryList->lpbin =
				mapi::allocate<LPSBinary>(m_lpNewEntryList->cValues * sizeof(SBinary), m_lpNewEntryList);

			for (ULONG paneID = 0; paneID < m_lpNewEntryList->cValues; paneID++)
			{
				const auto lpData = GetListRowData(LISTNUM, paneID);
				if (lpData)
				{
					const auto binary = lpData->cast<sortlistdata::binaryData>();
					if (binary)
					{
						m_lpNewEntryList->lpbin[paneID] = binary->detachBin(m_lpNewEntryList);
					}
				}
			}
		}

		if (!m_lpNewRes && m_lpSourceRes)
		{
			EC_H_S(mapi::HrCopyRestriction(m_lpSourceRes, nullptr, &m_lpNewRes));
		}

		m_ulNewSearchFlags = GetHex(2);
	}
} // namespace dialog::editor