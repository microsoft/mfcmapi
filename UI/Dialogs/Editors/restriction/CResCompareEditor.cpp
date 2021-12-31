#include <StdAfx.h>
#include <UI/Dialogs/Editors/restriction/CResCompareEditor.h>
#include <core/mapi/extraPropTags.h>
#include <core/utility/output.h>
#include <core/interpret/flags.h>
#include <core/interpret/proptags.h>
#include <core/addin/mfcmapi.h>

namespace dialog::editor
{
	CResCompareEditor::CResCompareEditor(_In_ CWnd* pParentWnd, ULONG ulRelop, ULONG ulPropTag1, ULONG ulPropTag2)
		: CEditor(pParentWnd, IDS_RESED, IDS_RESEDCOMPPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
	{
		TRACE_CONSTRUCTOR(COMPCLASS);

		SetPromptPostFix(flags::AllFlagsToString(flagRelop, false));
		AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_RELOP, false));
		SetHex(0, ulRelop);
		const auto szFlags = flags::InterpretFlags(flagRelop, ulRelop);
		AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_RELOP, szFlags, true));
		AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_ULPROPTAG1, false));
		SetHex(2, ulPropTag1);
		AddPane(viewpane::TextPane::CreateSingleLinePane(
			3, IDS_ULPROPTAG1, proptags::TagToString(ulPropTag1, nullptr, false, true), true));

		AddPane(viewpane::TextPane::CreateSingleLinePane(4, IDS_ULPROPTAG2, false));
		SetHex(4, ulPropTag2);
		AddPane(viewpane::TextPane::CreateSingleLinePane(
			5, IDS_ULPROPTAG1, proptags::TagToString(ulPropTag2, nullptr, false, true), true));
	}

	_Check_return_ ULONG CResCompareEditor::HandleChange(UINT nID)
	{
		const auto paneID = CEditor::HandleChange(nID);

		if (paneID == 0)
		{
			SetStringW(1, flags::InterpretFlags(flagRelop, GetHex(0)));
		}
		else if (paneID == 2)
		{
			SetStringW(3, proptags::TagToString(GetPropTag(2), nullptr, false, true));
		}
		else if (paneID == 4)
		{
			SetStringW(5, proptags::TagToString(GetPropTag(4), nullptr, false, true));
		}

		return paneID;
	}
} // namespace dialog::editor