#include <StdAfx.h>
#include <UI/Dialogs/Editors/restriction/RestrictEditor.h>
#include <UI/Dialogs/Editors/restriction/CResAndOrEditor.h>
#include <UI/Dialogs/Editors/propeditor/ipropeditor.h>
#include <core/sortlistdata/resData.h>
#include <core/mapi/mapiMemory.h>
#include <core/utility/output.h>
#include <core/addin/mfcmapi.h>
#include <core/property/parseProperty.h>

namespace dialog::editor
{
	CResAndOrEditor::CResAndOrEditor(
		_In_ CWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		_In_ const _SRestriction* lpRes,
		_In_ LPVOID lpAllocParent)
		: CEditor(pParentWnd, IDS_RESED, IDS_RESEDANDORPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
	{
		TRACE_CONSTRUCTOR(ANDORCLASS);
		m_lpMapiObjects = lpMapiObjects;
		m_lpRes = lpRes;
		m_lpNewResArray = nullptr;
		m_ulNewResCount = NULL;
		m_lpAllocParent = lpAllocParent;

		AddPane(viewpane::ListPane::Create(0, IDS_SUBRESTRICTIONS, false, false, ListEditCallBack(this)));
		SetListID(0);
	}

	// Used to call functions which need to be called AFTER controls are created
	BOOL CResAndOrEditor::OnInitDialog()
	{
		const auto bRet = CEditor::OnInitDialog();

		InitListFromRestriction(0, m_lpRes);

		UpdateButtons();

		return bRet;
	}

	_Check_return_ LPSRestriction CResAndOrEditor::DetachModifiedSRestrictionArray() noexcept
	{
		const auto lpRet = m_lpNewResArray;
		m_lpNewResArray = nullptr;
		return lpRet;
	}

	_Check_return_ ULONG CResAndOrEditor::GetResCount() const noexcept { return m_ulNewResCount; }

	void CResAndOrEditor::InitListFromRestriction(ULONG ulListNum, _In_ const _SRestriction* lpRes) const
	{
		ClearList(ulListNum);
		InsertColumn(ulListNum, 0, IDS_SHARP);

		if (!lpRes) return;
		switch (lpRes->rt)
		{
		case RES_AND:
			InsertColumn(ulListNum, 1, IDS_SUBRESTRICTION);

			for (ULONG paneID = 0; paneID < lpRes->res.resAnd.cRes; paneID++)
			{
				auto lpData = InsertListRow(ulListNum, paneID, std::to_wstring(paneID));
				if (lpData)
				{
					sortlistdata::resData::init(lpData, &lpRes->res.resAnd.lpRes[paneID]);
					SetListString(
						ulListNum, paneID, 1, property::RestrictionToString(&lpRes->res.resAnd.lpRes[paneID], nullptr));
				}
			}
			break;
		case RES_OR:
			InsertColumn(ulListNum, 1, IDS_SUBRESTRICTION);

			for (ULONG paneID = 0; paneID < lpRes->res.resOr.cRes; paneID++)
			{
				auto lpData = InsertListRow(ulListNum, paneID, std::to_wstring(paneID));
				if (lpData)
				{
					sortlistdata::resData::init(lpData, &lpRes->res.resOr.lpRes[paneID]);
					SetListString(
						ulListNum, paneID, 1, property::RestrictionToString(&lpRes->res.resOr.lpRes[paneID], nullptr));
				}
			}
			break;
		}
		ResizeList(ulListNum, false);
	}

	_Check_return_ bool CResAndOrEditor::DoListEdit(ULONG ulListNum, int iItem, _In_ sortlistdata::sortListData* lpData)
	{
		if (!lpData) return false;
		if (!lpData->cast<sortlistdata::resData>())
		{
			sortlistdata::resData::init(lpData, nullptr);
		}

		const auto res = lpData->cast<sortlistdata::resData>();
		if (!res) return false;

		const auto lpSourceRes = res->getCurrentRes();

		CRestrictEditor MyResEditor(this, m_lpMapiObjects, m_lpAllocParent,
									lpSourceRes); // pass source res into editor
		if (!MyResEditor.DisplayDialog()) return false;
		// Since lpData->data.Res.lpNewRes was owned by an m_lpAllocParent, we don't free it directly
		const auto newRes = MyResEditor.DetachModifiedSRestriction();
		res->setCurrentRes(newRes);
		SetListString(ulListNum, iItem, 1, property::RestrictionToString(newRes, nullptr));
		return true;
	}

	// Create our LPSRestriction array from the dialog here
	void CResAndOrEditor::OnOK()
	{
		CMyDialog::OnOK(); // don't need to call CEditor::OnOK

		const auto ulNewResCount = GetListCount(0);

		if (ulNewResCount > ULONG_MAX / sizeof(SRestriction)) return;
		auto lpNewResArray = mapi::allocate<LPSRestriction>(sizeof(SRestriction) * ulNewResCount, m_lpAllocParent);
		if (lpNewResArray)
		{
			for (ULONG paneID = 0; paneID < ulNewResCount; paneID++)
			{
				const auto lpData = GetListRowData(0, paneID);
				if (lpData)
				{
					const auto res = lpData->cast<sortlistdata::resData>();
					if (res)
					{
						lpNewResArray[paneID] = res->detachRes(m_lpAllocParent);
					}
				}
			}

			m_ulNewResCount = ulNewResCount;
			m_lpNewResArray = lpNewResArray;
		}
	}
} // namespace dialog::editor