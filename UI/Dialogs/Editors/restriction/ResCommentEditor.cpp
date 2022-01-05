#include <StdAfx.h>
#include <UI/Dialogs/Editors/restriction/RestrictEditor.h>
#include <UI/Dialogs/Editors/restriction/ResCommentEditor.h>
#include <UI/Dialogs/Editors/propeditor/ipropeditor.h>
#include <core/sortlistdata/commentData.h>
#include <core/mapi/mapiMemory.h>
#include <core/utility/output.h>
#include <core/interpret/proptags.h>
#include <core/addin/mfcmapi.h>
#include <core/property/parseProperty.h>

namespace dialog::editor
{
	ResCommentEditor::ResCommentEditor(
		_In_ CWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		_In_ const _SRestriction* lpRes,
		_In_ LPVOID lpAllocParent)
		: CEditor(
			  pParentWnd,
			  IDS_COMMENTRESED,
			  IDS_RESEDCOMMENTPROMPT,
			  CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL,
			  IDS_ACTIONEDITRES,
			  NULL,
			  NULL)
	{
		TRACE_CONSTRUCTOR(COMMENTCLASS);
		m_lpMapiObjects = lpMapiObjects;

		m_lpSourceRes = lpRes;
		m_lpNewCommentRes = nullptr;
		m_lpNewCommentProp = nullptr;
		m_lpAllocParent = lpAllocParent;

		AddPane(
			viewpane::ListPane::CreateCollapsibleListPane(0, IDS_SUBRESTRICTION, false, false, ListEditCallBack(this)));
		SetListID(0);
		AddPane(viewpane::TextPane::CreateMultiLinePane(
			1, IDS_RESTRICTIONTEXT, property::RestrictionToString(m_lpSourceRes->res.resComment.lpRes, nullptr), true));
	}

	// Used to call functions which need to be called AFTER controls are created
	BOOL ResCommentEditor::OnInitDialog()
	{
		const auto bRet = CEditor::OnInitDialog();

		InitListFromPropArray(0, m_lpSourceRes->res.resComment.cValues, m_lpSourceRes->res.resComment.lpProp);

		UpdateButtons();

		return bRet;
	}

	_Check_return_ const _SRestriction* ResCommentEditor::GetSourceRes() const noexcept
	{
		if (m_lpNewCommentRes) return m_lpNewCommentRes;
		if (m_lpSourceRes && m_lpSourceRes->res.resComment.lpRes) return m_lpSourceRes->res.resComment.lpRes;
		return nullptr;
	}

	_Check_return_ LPSRestriction ResCommentEditor::DetachModifiedSRestriction() noexcept
	{
		const auto lpRet = m_lpNewCommentRes;
		m_lpNewCommentRes = nullptr;
		return lpRet;
	}

	_Check_return_ LPSPropValue ResCommentEditor::DetachModifiedSPropValue() noexcept
	{
		const auto lpRet = m_lpNewCommentProp;
		m_lpNewCommentProp = nullptr;
		return lpRet;
	}

	_Check_return_ ULONG ResCommentEditor::GetSPropValueCount() const noexcept { return m_ulNewCommentProp; }

	void ResCommentEditor::InitListFromPropArray(
		ULONG ulListNum,
		ULONG cProps,
		_In_count_(cProps) const _SPropValue* lpProps) const
	{
		ClearList(ulListNum);

		InsertColumn(ulListNum, 0, IDS_SHARP);
		std::wstring szProp;
		std::wstring szAltProp;
		InsertColumn(ulListNum, 1, IDS_PROPERTY);
		InsertColumn(ulListNum, 2, IDS_VALUE);
		InsertColumn(ulListNum, 3, IDS_ALTERNATEVIEW);

		for (ULONG paneID = 0; paneID < cProps; paneID++)
		{
			auto lpData = InsertListRow(ulListNum, paneID, std::to_wstring(paneID));
			if (lpData)
			{
				sortlistdata::commentData::init(lpData, &lpProps[paneID]);
				SetListString(
					ulListNum, paneID, 1, proptags::TagToString(lpProps[paneID].ulPropTag, nullptr, false, true));
				property::parseProperty(&lpProps[paneID], &szProp, &szAltProp);
				SetListString(ulListNum, paneID, 2, szProp);
				SetListString(ulListNum, paneID, 3, szAltProp);
			}
		}

		ResizeList(ulListNum, false);
	}

	void ResCommentEditor::OnEditAction1()
	{
		const auto lpSourceRes = GetSourceRes();

		RestrictEditor MyResEditor(this, m_lpMapiObjects, m_lpAllocParent,
									lpSourceRes); // pass source res into editor
		if (!MyResEditor.DisplayDialog()) return;

		// Since m_lpNewCommentRes was owned by an m_lpAllocParent, we don't free it directly
		m_lpNewCommentRes = MyResEditor.DetachModifiedSRestriction();
		SetStringW(1, property::RestrictionToString(m_lpNewCommentRes, nullptr));
	}

	_Check_return_ bool
	ResCommentEditor::DoListEdit(ULONG ulListNum, int iItem, _In_ sortlistdata::sortListData* lpData)
	{
		if (!m_lpAllocParent) return false;
		if (!lpData) return false;
		if (!lpData->cast<sortlistdata::commentData>())
		{
			sortlistdata::commentData::init(lpData, nullptr);
		}

		const auto comment = lpData->cast<sortlistdata::commentData>();
		if (!comment) return false;

		auto lpSourceProp = comment->getCurrentProp();

		auto sProp = SPropValue{};

		if (!lpSourceProp)
		{
			CEditor MyTag(this, IDS_TAG, IDS_TAGPROMPT, true);

			MyTag.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_TAG, false));

			if (!MyTag.DisplayDialog()) return false;
			sProp.ulPropTag = MyTag.GetHex(0);
			lpSourceProp = &sProp;
		}

		const auto propEditor = DisplayPropertyEditor(
			this, m_lpMapiObjects, IDS_PROPEDITOR, L"", false, nullptr, NULL, false, lpSourceProp);

		if (propEditor)
		{
			LPSPropValue prop = nullptr;
			// Populate prop with ownership by m_lpAllocParent
			WC_MAPI_S(mapi::HrDupPropset(1, propEditor->getValue(), m_lpAllocParent, &prop));
			// Now give prop to comment to hold
			comment->setCurrentProp(prop);
			std::wstring szTmp;
			std::wstring szAltTmp;
			SetListString(ulListNum, iItem, 1, proptags::TagToString(prop->ulPropTag, nullptr, false, true));
			property::parseProperty(prop, &szTmp, &szAltTmp);
			SetListString(ulListNum, iItem, 2, szTmp);
			SetListString(ulListNum, iItem, 3, szAltTmp);
			return true;
		}

		return false;
	}

	void ResCommentEditor::OnOK()
	{
		CMyDialog::OnOK(); // don't need to call CEditor::OnOK

		const auto ulNewCommentProp = GetListCount(0);

		if (ulNewCommentProp && ulNewCommentProp < ULONG_MAX / sizeof(SPropValue))
		{
			auto lpNewCommentProp =
				mapi::allocate<LPSPropValue>(sizeof(SPropValue) * ulNewCommentProp, m_lpAllocParent);
			if (lpNewCommentProp)
			{
				for (ULONG paneID = 0; paneID < ulNewCommentProp; paneID++)
				{
					const auto lpData = GetListRowData(0, paneID);
					if (lpData)
					{
						const auto comment = lpData->cast<sortlistdata::commentData>();
						if (comment)
						{
							EC_H_S(mapi::MyPropCopyMore(
								&lpNewCommentProp[paneID],
								comment->getCurrentProp(),
								MAPIAllocateMore,
								m_lpAllocParent));
						}
					}
				}

				m_ulNewCommentProp = ulNewCommentProp;
				m_lpNewCommentProp = lpNewCommentProp;
			}
		}
		if (!m_lpNewCommentRes && m_lpSourceRes && m_lpSourceRes->res.resComment.lpRes)
		{
			EC_H_S(mapi::HrCopyRestriction(m_lpSourceRes->res.resComment.lpRes, m_lpAllocParent, &m_lpNewCommentRes));
		}
	}
} // namespace dialog::editor