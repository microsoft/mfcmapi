#include <StdAfx.h>
#include <UI/Dialogs/Editors/restriction/RestrictEditor.h>
#include <UI/Dialogs/Editors/restriction/CResSubResEditor.h>
#include <core/utility/output.h>
#include <core/interpret/proptags.h>
#include <core/addin/mfcmapi.h>
#include <core/property/parseProperty.h>

namespace dialog::editor
{
	CResSubResEditor::CResSubResEditor(
		_In_ CWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		ULONG ulSubObject,
		_In_ const _SRestriction* lpRes,
		_In_ LPVOID lpAllocParent)
		: CEditor(
			  pParentWnd,
			  IDS_SUBRESED,
			  IDS_RESEDSUBPROMPT,
			  CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL,
			  IDS_ACTIONEDITRES,
			  NULL,
			  NULL)
	{
		TRACE_CONSTRUCTOR(SUBRESCLASS);
		m_lpMapiObjects = lpMapiObjects;

		m_lpOldRes = lpRes;
		m_lpNewRes = nullptr;
		m_lpAllocParent = lpAllocParent;

		AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_ULSUBOBJECT, false));
		SetHex(0, ulSubObject);
		AddPane(viewpane::TextPane::CreateSingleLinePane(
			1, IDS_ULSUBOBJECT, proptags::TagToString(ulSubObject, nullptr, false, true), true));

		AddPane(
			viewpane::TextPane::CreateMultiLinePane(2, IDS_LPRES, property::RestrictionToString(lpRes, nullptr), true));
	}

	_Check_return_ ULONG CResSubResEditor::HandleChange(UINT nID)
	{
		const auto paneID = CEditor::HandleChange(nID);

		if (paneID == 0)
		{
			SetStringW(1, proptags::TagToString(GetPropTag(0), nullptr, false, true));
		}

		return paneID;
	}

	_Check_return_ LPSRestriction CResSubResEditor::DetachModifiedSRestriction() noexcept
	{
		const auto lpRet = m_lpNewRes;
		m_lpNewRes = nullptr;
		return lpRet;
	}

	void CResSubResEditor::OnEditAction1()
	{
		CRestrictEditor ResEdit(this, m_lpMapiObjects, m_lpAllocParent, m_lpNewRes ? m_lpNewRes : m_lpOldRes);

		if (!ResEdit.DisplayDialog()) return;

		// Since m_lpNewRes was owned by an m_lpAllocParent, we don't free it directly
		m_lpNewRes = ResEdit.DetachModifiedSRestriction();

		SetStringW(2, property::RestrictionToString(m_lpNewRes, nullptr));
	}
} // namespace dialog::editor