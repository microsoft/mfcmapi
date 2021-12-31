#include <StdAfx.h>
#include <UI/Dialogs/Editors/restriction/RestrictEditor.h>
#include <UI/Dialogs/Editors/restriction/ResExistEditor.h>
#include <core/utility/output.h>
#include <core/interpret/proptags.h>
#include <core/addin/mfcmapi.h>

namespace dialog::editor
{
	ResExistEditor::ResExistEditor(_In_ CWnd* pParentWnd, ULONG ulPropTag)
		: CEditor(pParentWnd, IDS_RESED, IDS_RESEDEXISTPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
	{
		TRACE_CONSTRUCTOR(EXISTCLASS);

		AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_ULPROPTAG, false));
		SetHex(0, ulPropTag);
		AddPane(viewpane::TextPane::CreateSingleLinePane(
			1, IDS_ULPROPTAG, proptags::TagToString(ulPropTag, nullptr, false, true), true));
	}

	_Check_return_ ULONG ResExistEditor::HandleChange(UINT nID)
	{
		const auto paneID = CEditor::HandleChange(nID);

		if (paneID == 0)
		{
			SetStringW(1, proptags::TagToString(GetPropTag(0), nullptr, false, true));
		}

		return paneID;
	}
} // namespace dialog::editor