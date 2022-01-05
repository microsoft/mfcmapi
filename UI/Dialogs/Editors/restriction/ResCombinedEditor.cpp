#include <StdAfx.h>
#include <UI/Dialogs/Editors/restriction/RestrictEditor.h>
#include <UI/Dialogs/Editors/restriction/ResCombinedEditor.h>
#include <UI/Dialogs/Editors/propeditor/ipropeditor.h>
#include <core/mapi/extraPropTags.h>
#include <core/utility/output.h>
#include <core/interpret/flags.h>
#include <core/interpret/proptags.h>
#include <core/addin/mfcmapi.h>
#include <core/property/parseProperty.h>

namespace dialog::editor
{
	ResCombinedEditor::ResCombinedEditor(
		_In_ CWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		ULONG ulResType,
		ULONG ulCompare,
		ULONG ulPropTag,
		_In_ const _SPropValue* lpProp,
		_In_ LPVOID lpAllocParent)
		: CEditor(
			  pParentWnd,
			  IDS_RESED,
			  ulResType == RES_CONTENT ? IDS_RESEDCONTPROMPT : // Content case
				  ulResType == RES_PROPERTY ? IDS_RESEDPROPPROMPT
											: // Property case
				  0, // else case
			  CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL,
			  IDS_ACTIONEDITPROP,
			  NULL,
			  NULL)
	{
		TRACE_CONSTRUCTOR(CONTENTCLASS);
		m_lpMapiObjects = lpMapiObjects;

		m_ulResType = ulResType;
		m_lpOldProp = lpProp;
		m_lpNewProp = nullptr;
		m_lpAllocParent = lpAllocParent;

		std::wstring szFlags;

		if (RES_CONTENT == m_ulResType)
		{
			SetPromptPostFix(flags::AllFlagsToString(flagFuzzyLevel, true));

			AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_ULFUZZYLEVEL, false));
			SetHex(0, ulCompare);
			szFlags = flags::InterpretFlags(flagFuzzyLevel, ulCompare);
			AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_ULFUZZYLEVEL, szFlags, true));
		}
		else if (RES_PROPERTY == m_ulResType)
		{
			SetPromptPostFix(flags::AllFlagsToString(flagRelop, false));
			AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_RELOP, false));
			SetHex(0, ulCompare);
			szFlags = flags::InterpretFlags(flagRelop, ulCompare);
			AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_RELOP, szFlags, true));
		}

		AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_ULPROPTAG, false));
		SetHex(2, ulPropTag);
		AddPane(viewpane::TextPane::CreateSingleLinePane(
			3, IDS_ULPROPTAG, proptags::TagToString(ulPropTag, nullptr, false, true), true));

		AddPane(viewpane::TextPane::CreateSingleLinePane(4, IDS_LPPROPULPROPTAG, false));
		if (lpProp) SetHex(4, lpProp->ulPropTag);
		AddPane(viewpane::TextPane::CreateSingleLinePane(
			5,
			IDS_LPPROPULPROPTAG,
			lpProp ? proptags::TagToString(lpProp->ulPropTag, nullptr, false, true) : strings::emptystring,
			true));

		std::wstring szProp;
		std::wstring szAltProp;
		if (lpProp) property::parseProperty(lpProp, &szProp, &szAltProp);
		AddPane(viewpane::TextPane::CreateMultiLinePane(6, IDS_LPPROP, szProp, true));
		AddPane(viewpane::TextPane::CreateMultiLinePane(7, IDS_LPPROPALTVIEW, szAltProp, true));
	}

	_Check_return_ ULONG ResCombinedEditor::HandleChange(UINT nID)
	{
		const auto paneID = CEditor::HandleChange(nID);

		std::wstring szFlags;
		if (paneID == 0)
		{
			if (RES_CONTENT == m_ulResType)
			{
				SetStringW(1, flags::InterpretFlags(flagFuzzyLevel, GetHex(0)));
			}
			else if (RES_PROPERTY == m_ulResType)
			{
				SetStringW(1, flags::InterpretFlags(flagRelop, GetHex(0)));
			}
		}
		else if (paneID == 2)
		{
			SetStringW(3, proptags::TagToString(GetPropTag(2), nullptr, false, true));
		}
		else if (paneID == 4)
		{
			SetStringW(5, proptags::TagToString(GetPropTag(4), nullptr, false, true));
			m_lpOldProp = nullptr;
			m_lpNewProp = nullptr;
			SetStringW(6, L"");
			SetStringW(7, L"");
		}

		return paneID;
	}

	_Check_return_ LPSPropValue ResCombinedEditor::DetachModifiedSPropValue() noexcept
	{
		const auto lpRet = m_lpNewProp;
		m_lpNewProp = nullptr;
		return lpRet;
	}

	void ResCombinedEditor::OnEditAction1()
	{
		if (!m_lpAllocParent) return;

		auto lpEditProp = m_lpOldProp;
		if (m_lpNewProp) lpEditProp = m_lpNewProp;

		const auto propEditor = DisplayPropertyEditor(
			this, m_lpMapiObjects, IDS_PROPEDITOR, L"", false, nullptr, GetPropTag(4), false, lpEditProp);

		if (propEditor)
		{
			// Populate m_lpNewProp with ownership by m_lpAllocParent
			WC_MAPI_S(mapi::HrDupPropset(1, propEditor->getValue(), m_lpAllocParent, &m_lpNewProp));

			std::wstring szProp;
			std::wstring szAltProp;
			property::parseProperty(m_lpNewProp, &szProp, &szAltProp);
			SetStringW(6, szProp);
			SetStringW(7, szAltProp);
		}
	}
} // namespace dialog::editor