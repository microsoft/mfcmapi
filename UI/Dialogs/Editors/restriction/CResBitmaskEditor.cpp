#include <StdAfx.h>
#include <UI/Dialogs/Editors/restriction/CResBitmaskEditor.h>
#include <core/mapi/extraPropTags.h>
#include <core/utility/output.h>
#include <core/interpret/flags.h>
#include <core/interpret/proptags.h>
#include <core/addin/mfcmapi.h>

namespace dialog::editor
{
	CResBitmaskEditor::CResBitmaskEditor(_In_ CWnd* pParentWnd, ULONG relBMR, ULONG ulPropTag, ULONG ulMask)
		: CEditor(pParentWnd, IDS_RESED, IDS_RESEDBITPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
	{
		TRACE_CONSTRUCTOR(BITMASKCLASS);

		SetPromptPostFix(flags::AllFlagsToString(flagBitmask, false));
		AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_RELBMR, false));
		SetHex(0, relBMR);
		const auto szFlags = flags::InterpretFlags(flagBitmask, relBMR);
		AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_RELBMR, szFlags, true));
		AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_ULPROPTAG, false));
		SetHex(2, ulPropTag);
		AddPane(viewpane::TextPane::CreateSingleLinePane(
			3, IDS_ULPROPTAG, proptags::TagToString(ulPropTag, nullptr, false, true), true));

		AddPane(viewpane::TextPane::CreateSingleLinePane(4, IDS_MASK, false));
		SetHex(4, ulMask);
	}

	_Check_return_ ULONG CResBitmaskEditor::HandleChange(UINT nID)
	{
		const auto paneID = CEditor::HandleChange(nID);

		if (paneID == 0)
		{
			SetStringW(1, flags::InterpretFlags(flagBitmask, GetHex(0)));
		}
		else if (paneID == 2)
		{
			SetStringW(3, proptags::TagToString(GetPropTag(2), nullptr, false, true));
		}

		return paneID;
	}
} // namespace dialog::editor