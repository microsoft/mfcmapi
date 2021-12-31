#include <StdAfx.h>
#include <UI/Dialogs/Editors/restriction/RestrictEditor.h>
#include <UI/Dialogs/Editors/restriction/ResAndOrEditor.h>
#include <UI/Dialogs/Editors/restriction/ResBitmaskEditor.h>
#include <UI/Dialogs/Editors/restriction/ResCombinedEditor.h>
#include <UI/Dialogs/Editors/restriction/ResCommentEditor.h>
#include <UI/Dialogs/Editors/restriction/ResCompareEditor.h>
#include <UI/Dialogs/Editors/restriction/ResExistEditor.h>
#include <UI/Dialogs/Editors/restriction/ResSizeEditor.h>
#include <UI/Dialogs/Editors/restriction/ResSubResEditor.h>
#include <core/mapi/extraPropTags.h>
#include <core/mapi/mapiMemory.h>
#include <core/utility/output.h>
#include <core/interpret/flags.h>
#include <core/addin/mfcmapi.h>
#include <core/property/parseProperty.h>

namespace dialog::editor
{
	// Note that an alloc parent is passed in to CRestrictEditor. If a parent isn't passed, we allocate one ourselves.
	// All other memory allocated in CRestrictEditor is owned by the parent and must not be freed manually
	// If we return (detach) memory to a caller, they must MAPIFreeBuffer only if they did not pass in a parent
	static std::wstring CLASS = L"CRestrictEditor"; // STRING_OK
	// Create an editor for a restriction
	// Takes LPSRestriction lpRes as input
	CRestrictEditor::CRestrictEditor(
		_In_ CWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		_In_opt_ LPVOID lpAllocParent,
		_In_opt_ const _SRestriction* lpRes)
		: CEditor(
			  pParentWnd,
			  IDS_RESED,
			  IDS_RESEDPROMPT,
			  CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL,
			  IDS_ACTIONEDITRES,
			  NULL,
			  NULL)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_lpMapiObjects = lpMapiObjects;

		// Not copying the source restriction, but since we're modal it won't matter
		m_lpRes = lpRes;
		m_lpOutputRes = nullptr;
		m_bModified = false;
		if (!lpRes) m_bModified = true; // if we don't have a source, make sure we never look at it

		m_lpAllocParent = lpAllocParent;

		// Set up our output restriction and determine m_lpAllocParent- this will be the basis for all other allocations
		// Allocate base memory:
		m_lpOutputRes = mapi::allocate<LPSRestriction>(sizeof(SRestriction), m_lpAllocParent);
		if (!m_lpAllocParent)
		{
			m_lpAllocParent = m_lpOutputRes;
		}

		if (m_lpOutputRes)
		{
			memset(m_lpOutputRes, 0, sizeof(SRestriction));
			if (m_lpRes) m_lpOutputRes->rt = m_lpRes->rt;
		}

		SetPromptPostFix(flags::AllFlagsToString(flagRestrictionType, true));
		AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_RESTRICTIONTYPE, false)); // type as a number
		AddPane(viewpane::TextPane::CreateSingleLinePane(
			1, IDS_RESTRICTIONTYPE, true)); // type as a string (flagRestrictionType)
		AddPane(viewpane::TextPane::CreateMultiLinePane(
			2, IDS_RESTRICTIONTEXT, property::RestrictionToString(GetSourceRes(), nullptr), true));
	}

	CRestrictEditor::~CRestrictEditor()
	{
		TRACE_DESTRUCTOR(CLASS);
		// Only if we self allocated m_lpOutputRes and did not detach it can we free it
		if (m_lpAllocParent == m_lpOutputRes)
		{
			MAPIFreeBuffer(m_lpOutputRes);
		}
	}

	// Used to call functions which need to be called AFTER controls are created
	BOOL CRestrictEditor::OnInitDialog()
	{
		const auto bRet = CEditor::OnInitDialog();

		const auto lpSourceRes = GetSourceRes();

		if (lpSourceRes)
		{
			SetHex(0, lpSourceRes->rt);
			SetStringW(1, flags::InterpretFlags(flagRestrictionType, lpSourceRes->rt));
		}

		return bRet;
	}

	_Check_return_ const _SRestriction* CRestrictEditor::GetSourceRes() const noexcept
	{
		if (m_lpRes && !m_bModified) return m_lpRes;
		return m_lpOutputRes;
	}

	// Create our LPSRestriction from the dialog here
	void CRestrictEditor::OnOK()
	{
		CMyDialog::OnOK(); // don't need to call CEditor::OnOK
	}

	_Check_return_ LPSRestriction CRestrictEditor::DetachModifiedSRestriction() noexcept
	{
		if (!m_bModified) return nullptr;
		const auto lpRet = m_lpOutputRes;
		m_lpOutputRes = nullptr;
		return lpRet;
	}

	// CEditor::HandleChange will return the control number of what changed
	// so I can react to it if I need to
	_Check_return_ ULONG CRestrictEditor::HandleChange(UINT nID)
	{
		const auto paneID = CEditor::HandleChange(nID);

		// If the restriction type changed
		if (paneID == 0)
		{
			if (!m_lpOutputRes) return paneID;
			const auto ulOldResType = m_lpOutputRes->rt;
			const auto ulNewResType = GetHex(paneID);

			if (ulOldResType == ulNewResType) return paneID;

			SetStringW(1, flags::InterpretFlags(flagRestrictionType, ulNewResType));

			m_bModified = true;
			if ((ulOldResType == RES_AND || ulOldResType == RES_OR) &&
				(ulNewResType == RES_AND || ulNewResType == RES_OR))
			{
				m_lpOutputRes->rt = ulNewResType; // just need to adjust our type and refresh
			}
			else
			{
				// Need to wipe out our current output restriction and rebuild it
				if (m_lpAllocParent == m_lpOutputRes)
				{
					// We allocated m_lpOutputRes directly, so we can and should free it before replacing the pointer
					MAPIFreeBuffer(m_lpOutputRes);
					m_lpOutputRes = mapi::allocate<LPSRestriction>(sizeof(SRestriction));
					m_lpAllocParent = m_lpOutputRes;
				}
				else
				{
					// If the pointers are different, m_lpOutputRes was allocated with MAPIAllocateMore
					// Since m_lpOutputRes is owned by m_lpAllocParent, we don't free it directly
					m_lpOutputRes = mapi::allocate<LPSRestriction>(sizeof(SRestriction), m_lpAllocParent);
				}

				m_lpOutputRes->rt = ulNewResType;
			}

			SetStringW(2, property::RestrictionToString(m_lpOutputRes, nullptr));
		}
		return paneID;
	}

	void CRestrictEditor::OnEditAction1()
	{
		auto hRes = S_OK;
		const auto lpSourceRes = GetSourceRes();
		if (!lpSourceRes || !m_lpOutputRes) return;

		switch (lpSourceRes->rt)
		{
		case RES_COMPAREPROPS:
			hRes = WC_H(EditCompare(lpSourceRes));
			break;
			// Structures for these two types are identical
		case RES_OR:
		case RES_AND:
			hRes = WC_H(EditAndOr(lpSourceRes));
			break;
			// Structures for these two types are identical
		case RES_NOT:
		case RES_COUNT:
			hRes = WC_H(EditRestrict(lpSourceRes));
			break;
			// Structures for these two types are identical
		case RES_PROPERTY:
		case RES_CONTENT:
			hRes = WC_H(EditCombined(lpSourceRes));
			break;
		case RES_BITMASK:
			hRes = WC_H(EditBitmask(lpSourceRes));
			break;
		case RES_SIZE:
			hRes = WC_H(EditSize(lpSourceRes));
			break;
		case RES_EXIST:
			hRes = WC_H(EditExist(lpSourceRes));
			break;
		case RES_SUBRESTRICTION:
			hRes = WC_H(EditSubrestriction(lpSourceRes));
			break;
			// Structures for these two types are identical
		case RES_COMMENT:
		case RES_ANNOTATION:
			hRes = WC_H(EditComment(lpSourceRes));
			break;
		}

		if (hRes == S_OK)
		{
			m_bModified = true;
			SetStringW(2, property::RestrictionToString(m_lpOutputRes, nullptr));
		}
	}

	HRESULT CRestrictEditor::EditCompare(const _SRestriction* lpSourceRes)
	{
		ResCompareEditor MyEditor(
			this,
			lpSourceRes->res.resCompareProps.relop,
			lpSourceRes->res.resCompareProps.ulPropTag1,
			lpSourceRes->res.resCompareProps.ulPropTag2);
		if (MyEditor.DisplayDialog())
		{
			m_lpOutputRes->rt = lpSourceRes->rt;
			m_lpOutputRes->res.resCompareProps.relop = MyEditor.GetHex(0);
			m_lpOutputRes->res.resCompareProps.ulPropTag1 = MyEditor.GetPropTag(2);
			m_lpOutputRes->res.resCompareProps.ulPropTag2 = MyEditor.GetPropTag(4);
			return S_OK;
		}

		return S_FALSE;
	}

	HRESULT CRestrictEditor::EditAndOr(const _SRestriction* lpSourceRes)
	{
		ResAndOrEditor MyResEditor(this, m_lpMapiObjects, lpSourceRes,
									m_lpAllocParent); // pass source res into editor
		if (MyResEditor.DisplayDialog())
		{
			m_lpOutputRes->rt = lpSourceRes->rt;
			m_lpOutputRes->res.resAnd.cRes = MyResEditor.GetResCount();

			const auto lpNewResArray = MyResEditor.DetachModifiedSRestrictionArray();
			if (lpNewResArray)
			{
				// I can do this because our new memory was allocated from a common parent
				MAPIFreeBuffer(m_lpOutputRes->res.resAnd.lpRes);
				m_lpOutputRes->res.resAnd.lpRes = lpNewResArray;
			}

			return S_OK;
		}

		return S_FALSE;
	}

	HRESULT CRestrictEditor::EditRestrict(const _SRestriction* lpSourceRes)
	{
		CRestrictEditor MyResEditor(
			this,
			m_lpMapiObjects,
			m_lpAllocParent,
			lpSourceRes->res.resNot.lpRes); // pass source res into editor
		if (MyResEditor.DisplayDialog())
		{
			m_lpOutputRes->rt = lpSourceRes->rt;
			// Since m_lpOutputRes->res.resNot.lpRes was owned by an m_lpAllocParent, we don't free it directly
			m_lpOutputRes->res.resNot.lpRes = MyResEditor.DetachModifiedSRestriction();
			return S_OK;
		}

		return S_FALSE;
	}

	HRESULT CRestrictEditor::EditCombined(const _SRestriction* lpSourceRes)
	{
		auto hRes = S_OK;
		ResCombinedEditor MyEditor(
			this,
			m_lpMapiObjects,
			lpSourceRes->rt,
			lpSourceRes->res.resContent.ulFuzzyLevel,
			lpSourceRes->res.resContent.ulPropTag,
			lpSourceRes->res.resContent.lpProp,
			m_lpAllocParent);
		if (!MyEditor.DisplayDialog()) return S_FALSE;

		m_lpOutputRes->rt = lpSourceRes->rt;
		m_lpOutputRes->res.resContent.ulFuzzyLevel = MyEditor.GetHex(0);
		m_lpOutputRes->res.resContent.ulPropTag = MyEditor.GetPropTag(2);

		// Since m_lpOutputRes->res.resContent.lpProp was owned by an m_lpAllocParent, we don't free it directly
		m_lpOutputRes->res.resContent.lpProp = MyEditor.DetachModifiedSPropValue();

		if (!m_lpOutputRes->res.resContent.lpProp)
		{
			// Got a problem here - the relop or fuzzy level was changed, but not the property
			// Need to copy the property from the source Res to the output Res
			m_lpOutputRes->res.resContent.lpProp = mapi::allocate<LPSPropValue>(sizeof(SPropValue), m_lpAllocParent);

			if (m_lpOutputRes->res.resContent.lpProp)
			{
				hRes = EC_H(mapi::MyPropCopyMore(
					m_lpOutputRes->res.resContent.lpProp,
					lpSourceRes->res.resContent.lpProp,
					MAPIAllocateMore,
					m_lpAllocParent));
			}
		}

		return hRes;
	}

	HRESULT CRestrictEditor::EditBitmask(const _SRestriction* lpSourceRes)
	{
		ResBitmaskEditor MyEditor(
			this,
			lpSourceRes->res.resBitMask.relBMR,
			lpSourceRes->res.resBitMask.ulPropTag,
			lpSourceRes->res.resBitMask.ulMask);
		if (MyEditor.DisplayDialog())
		{
			m_lpOutputRes->rt = lpSourceRes->rt;
			m_lpOutputRes->res.resBitMask.relBMR = MyEditor.GetHex(0);
			m_lpOutputRes->res.resBitMask.ulPropTag = MyEditor.GetPropTag(2);
			m_lpOutputRes->res.resBitMask.ulMask = MyEditor.GetHex(4);
			return S_OK;
		}

		return S_FALSE;
	}

	HRESULT CRestrictEditor::EditSize(const _SRestriction* lpSourceRes)
	{
		ResSizeEditor MyEditor(
			this, lpSourceRes->res.resSize.relop, lpSourceRes->res.resSize.ulPropTag, lpSourceRes->res.resSize.cb);
		if (MyEditor.DisplayDialog())
		{
			m_lpOutputRes->rt = lpSourceRes->rt;
			m_lpOutputRes->res.resSize.relop = MyEditor.GetHex(0);
			m_lpOutputRes->res.resSize.ulPropTag = MyEditor.GetPropTag(2);
			m_lpOutputRes->res.resSize.cb = MyEditor.GetHex(4);
			return S_OK;
		}

		return S_FALSE;
	}

	HRESULT CRestrictEditor::EditExist(const _SRestriction* lpSourceRes)
	{
		ResExistEditor MyEditor(this, lpSourceRes->res.resExist.ulPropTag);
		if (MyEditor.DisplayDialog())
		{
			m_lpOutputRes->rt = lpSourceRes->rt;
			m_lpOutputRes->res.resExist.ulPropTag = MyEditor.GetPropTag(0);
			m_lpOutputRes->res.resExist.ulReserved1 = 0;
			m_lpOutputRes->res.resExist.ulReserved2 = 0;
			return S_OK;
		}

		return S_FALSE;
	}

	HRESULT CRestrictEditor::EditSubrestriction(const _SRestriction* lpSourceRes)
	{
		ResSubResEditor MyEditor(
			this, m_lpMapiObjects, lpSourceRes->res.resSub.ulSubObject, lpSourceRes->res.resSub.lpRes, m_lpAllocParent);
		if (MyEditor.DisplayDialog())
		{
			m_lpOutputRes->rt = lpSourceRes->rt;
			m_lpOutputRes->res.resSub.ulSubObject = MyEditor.GetHex(1);

			// Since m_lpOutputRes->res.resSub.lpRes was owned by an m_lpAllocParent, we don't free it directly
			m_lpOutputRes->res.resSub.lpRes = MyEditor.DetachModifiedSRestriction();
			return S_OK;
		}

		return S_FALSE;
	}

	HRESULT CRestrictEditor::EditComment(const _SRestriction* lpSourceRes)
	{
		ResCommentEditor MyResEditor(
			this,
			m_lpMapiObjects,
			lpSourceRes,
			m_lpAllocParent); // pass source res into editor
		if (MyResEditor.DisplayDialog())
		{
			m_lpOutputRes->rt = lpSourceRes->rt;

			// Since m_lpOutputRes->res.resComment.lpRes was owned by an m_lpAllocParent, we don't free it directly
			m_lpOutputRes->res.resComment.lpRes = MyResEditor.DetachModifiedSRestriction();

			// Since m_lpOutputRes->res.resComment.lpProp was owned by an m_lpAllocParent, we don't free it directly
			m_lpOutputRes->res.resComment.lpProp = MyResEditor.DetachModifiedSPropValue();
			if (m_lpOutputRes->res.resComment.lpProp)
			{
				m_lpOutputRes->res.resComment.cValues = MyResEditor.GetSPropValueCount();
			}

			return S_OK;
		}

		return S_FALSE;
	}
} // namespace dialog::editor