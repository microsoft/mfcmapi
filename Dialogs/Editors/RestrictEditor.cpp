#include "stdafx.h"
#include "RestrictEditor.h"
#include "PropertyEditor.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"
#include "MAPIFunctions.h"
#include "ImportProcs.h"
#include "ExtraPropTags.h"
#include <Controls/SortList/ResData.h>
#include <Controls/SortList/CommentData.h>
#include <Controls/SortList/BinaryData.h>

static wstring COMPCLASS = L"CResCompareEditor"; // STRING_OK
class CResCompareEditor : public CEditor
{
public:
	CResCompareEditor(
		_In_ CWnd* pParentWnd,
		ULONG ulRelop,
		ULONG ulPropTag1,
		ULONG ulPropTag2);
private:
	_Check_return_ ULONG HandleChange(UINT nID) override;
};

CResCompareEditor::CResCompareEditor(
	_In_ CWnd* pParentWnd,
	ULONG ulRelop,
	ULONG ulPropTag1,
	ULONG ulPropTag2) :
	CEditor(pParentWnd, IDS_RESED, IDS_RESEDCOMPPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(COMPCLASS);

	SetPromptPostFix(AllFlagsToString(flagRelop, false));
	InitPane(0, TextPane::CreateSingleLinePane(IDS_RELOP, false));
	SetHex(0, ulRelop);
	auto szFlags = InterpretFlags(flagRelop, ulRelop);
	InitPane(1, TextPane::CreateSingleLinePane(IDS_RELOP, szFlags, true));
	InitPane(2, TextPane::CreateSingleLinePane(IDS_ULPROPTAG1, false));
	SetHex(2, ulPropTag1);
	InitPane(3, TextPane::CreateSingleLinePane(IDS_ULPROPTAG1, TagToString(ulPropTag1, nullptr, false, true), true));

	InitPane(4, TextPane::CreateSingleLinePane(IDS_ULPROPTAG2, false));
	SetHex(4, ulPropTag2);
	InitPane(5, TextPane::CreateSingleLinePane(IDS_ULPROPTAG1, TagToString(ulPropTag2, nullptr, false, true), true));
}

_Check_return_ ULONG CResCompareEditor::HandleChange(UINT nID)
{
	auto i = CEditor::HandleChange(nID);

	if (0 == i)
	{
		SetStringW(1, InterpretFlags(flagRelop, GetHex(0)));
	}
	else if (2 == i)
	{
		SetStringW(3, TagToString(GetPropTag(2), nullptr, false, true));
	}
	else if (4 == i)
	{
		SetStringW(5, TagToString(GetPropTag(4), nullptr, false, true));
	}
	return i;
}

// This class is only invoked by CRestrictEditor. CRestrictEditor always passes an alloc parent.
// So all memory detached from this class is owned by a parent and must not be freed manually
static wstring CONTENTCLASS = L"CResCombinedEditor"; // STRING_OK
class CResCombinedEditor : public CEditor
{
public:
	CResCombinedEditor(
		_In_ CWnd* pParentWnd,
		ULONG ulResType,
		ULONG ulCompare,
		ULONG ulPropTag,
		_In_ LPSPropValue lpProp,
		_In_ LPVOID lpAllocParent);

	void OnEditAction1() override;
	_Check_return_ LPSPropValue DetachModifiedSPropValue();

private:
	_Check_return_ ULONG HandleChange(UINT nID) override;

	ULONG m_ulResType;
	LPVOID m_lpAllocParent;
	LPSPropValue m_lpOldProp;
	LPSPropValue m_lpNewProp;
};

CResCombinedEditor::CResCombinedEditor(
	_In_ CWnd* pParentWnd,
	ULONG ulResType,
	ULONG ulCompare,
	ULONG ulPropTag,
	_In_ LPSPropValue lpProp,
	_In_ LPVOID lpAllocParent) :
	CEditor(pParentWnd,
		IDS_RESED,
		ulResType == RES_CONTENT ? IDS_RESEDCONTPROMPT : // Content case
		ulResType == RES_PROPERTY ? IDS_RESEDPROPPROMPT : // Property case
		0, // else case
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL,
		IDS_ACTIONEDITPROP,
		NULL,
		NULL)
{
	TRACE_CONSTRUCTOR(CONTENTCLASS);

	m_ulResType = ulResType;
	m_lpOldProp = lpProp;
	m_lpNewProp = nullptr;
	m_lpAllocParent = lpAllocParent;

	wstring szFlags;

	if (RES_CONTENT == m_ulResType)
	{
		SetPromptPostFix(AllFlagsToString(flagFuzzyLevel, true));

		InitPane(0, TextPane::CreateSingleLinePane(IDS_ULFUZZYLEVEL, false));
		SetHex(0, ulCompare);
		szFlags = InterpretFlags(flagFuzzyLevel, ulCompare);
		InitPane(1, TextPane::CreateSingleLinePane(IDS_ULFUZZYLEVEL, szFlags, true));
	}
	else if (RES_PROPERTY == m_ulResType)
	{
		SetPromptPostFix(AllFlagsToString(flagRelop, false));
		InitPane(0, TextPane::CreateSingleLinePane(IDS_RELOP, false));
		SetHex(0, ulCompare);
		szFlags = InterpretFlags(flagRelop, ulCompare);
		InitPane(1, TextPane::CreateSingleLinePane(IDS_RELOP, szFlags, true));
	}

	InitPane(2, TextPane::CreateSingleLinePane(IDS_ULPROPTAG, false));
	SetHex(2, ulPropTag);
	InitPane(3, TextPane::CreateSingleLinePane(IDS_ULPROPTAG, TagToString(ulPropTag, nullptr, false, true), true));

	InitPane(4, TextPane::CreateSingleLinePane(IDS_LPPROPULPROPTAG, false));
	if (lpProp) SetHex(4, lpProp->ulPropTag);
	InitPane(5, TextPane::CreateSingleLinePane(IDS_LPPROPULPROPTAG, lpProp ? TagToString(lpProp->ulPropTag, nullptr, false, true) : emptystring, true));

	wstring szProp;
	wstring szAltProp;
	if (lpProp) InterpretProp(lpProp, &szProp, &szAltProp);
	InitPane(6, TextPane::CreateMultiLinePane(IDS_LPPROP, szProp, true));
	InitPane(7, TextPane::CreateMultiLinePane(IDS_LPPROPALTVIEW, szAltProp, true));
}

_Check_return_ ULONG CResCombinedEditor::HandleChange(UINT nID)
{
	auto i = CEditor::HandleChange(nID);

	wstring szFlags;
	if (0 == i)
	{
		if (RES_CONTENT == m_ulResType)
		{
			SetStringW(1, InterpretFlags(flagFuzzyLevel, GetHex(0)));
		}
		else if (RES_PROPERTY == m_ulResType)
		{
			SetStringW(1, InterpretFlags(flagRelop, GetHex(0)));
		}
	}
	else if (2 == i)
	{
		SetStringW(3, TagToString(GetPropTag(2), nullptr, false, true));
	}
	else if (4 == i)
	{
		SetStringW(5, TagToString(GetPropTag(4), nullptr, false, true));
		m_lpOldProp = nullptr;
		m_lpNewProp = nullptr;
		SetStringW(6, L"");
		SetStringW(7, L"");
	}
	return i;
}

_Check_return_ LPSPropValue CResCombinedEditor::DetachModifiedSPropValue()
{
	auto m_lpRet = m_lpNewProp;
	m_lpNewProp = nullptr;
	return m_lpRet;
}

void CResCombinedEditor::OnEditAction1()
{
	if (!m_lpAllocParent) return;

	auto hRes = S_OK;

	auto lpEditProp = m_lpOldProp;
	LPSPropValue lpOutProp = nullptr;
	if (m_lpNewProp) lpEditProp = m_lpNewProp;

	WC_H(DisplayPropertyEditor(
		this,
		IDS_PROPEDITOR,
		NULL,
		false,
		m_lpAllocParent,
		NULL,
		GetPropTag(4),
		false,
		lpEditProp,
		&lpOutProp));

	// Since m_lpNewProp was owned by an m_lpAllocParent, we don't free it directly
	if (S_OK == hRes && lpOutProp)
	{
		m_lpNewProp = lpOutProp;
		wstring szProp;
		wstring szAltProp;

		InterpretProp(m_lpNewProp, &szProp, &szAltProp);
		SetStringW(6, szProp);
		SetStringW(7, szAltProp);
	}
}

static wstring BITMASKCLASS = L"CResBitmaskEditor"; // STRING_OK
class CResBitmaskEditor : public CEditor
{
public:
	CResBitmaskEditor(
		_In_ CWnd* pParentWnd,
		ULONG relBMR,
		ULONG ulPropTag,
		ULONG ulMask);

private:
	_Check_return_ ULONG HandleChange(UINT nID) override;
};

CResBitmaskEditor::CResBitmaskEditor(
	_In_ CWnd* pParentWnd,
	ULONG relBMR,
	ULONG ulPropTag,
	ULONG ulMask) :
	CEditor(pParentWnd, IDS_RESED, IDS_RESEDBITPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(BITMASKCLASS);

	SetPromptPostFix(AllFlagsToString(flagBitmask, false));
	InitPane(0, TextPane::CreateSingleLinePane(IDS_RELBMR, false));
	SetHex(0, relBMR);
	auto szFlags = InterpretFlags(flagBitmask, relBMR);
	InitPane(1, TextPane::CreateSingleLinePane(IDS_RELBMR, szFlags, true));
	InitPane(2, TextPane::CreateSingleLinePane(IDS_ULPROPTAG, false));
	SetHex(2, ulPropTag);
	InitPane(3, TextPane::CreateSingleLinePane(IDS_ULPROPTAG, TagToString(ulPropTag, nullptr, false, true), true));

	InitPane(4, TextPane::CreateSingleLinePane(IDS_MASK, false));
	SetHex(4, ulMask);
}

_Check_return_ ULONG CResBitmaskEditor::HandleChange(UINT nID)
{
	auto i = CEditor::HandleChange(nID);

	if (0 == i)
	{
		SetStringW(1, InterpretFlags(flagBitmask, GetHex(0)));
	}
	else if (2 == i)
	{
		SetStringW(3, TagToString(GetPropTag(2), nullptr, false, true));
	}

	return i;
}

static wstring SIZECLASS = L"CResSizeEditor"; // STRING_OK
class CResSizeEditor : public CEditor
{
public:
	CResSizeEditor(
		_In_ CWnd* pParentWnd,
		ULONG relop,
		ULONG ulPropTag,
		ULONG cb);

private:
	_Check_return_ ULONG HandleChange(UINT nID) override;
};

CResSizeEditor::CResSizeEditor(
	_In_ CWnd* pParentWnd,
	ULONG relop,
	ULONG ulPropTag,
	ULONG cb) :
	CEditor(pParentWnd, IDS_RESED, IDS_RESEDSIZEPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(SIZECLASS);

	SetPromptPostFix(AllFlagsToString(flagRelop, false));
	InitPane(0, TextPane::CreateSingleLinePane(IDS_RELOP, false));
	SetHex(0, relop);
	auto szFlags = InterpretFlags(flagRelop, relop);
	InitPane(1, TextPane::CreateSingleLinePane(IDS_RELOP, szFlags, true));

	InitPane(2, TextPane::CreateSingleLinePane(IDS_ULPROPTAG, false));
	SetHex(2, ulPropTag);
	InitPane(3, TextPane::CreateSingleLinePane(IDS_ULPROPTAG, TagToString(ulPropTag, nullptr, false, true), true));

	InitPane(4, TextPane::CreateSingleLinePane(IDS_CB, false));
	SetHex(4, cb);
}

_Check_return_ ULONG CResSizeEditor::HandleChange(UINT nID)
{
	auto i = CEditor::HandleChange(nID);

	if (0 == i)
	{
		SetStringW(1, InterpretFlags(flagRelop, GetHex(0)));
	}
	else if (2 == i)
	{
		SetStringW(3, TagToString(GetPropTag(2), nullptr, false, true));
	}
	return i;
}

static wstring EXISTCLASS = L"CResExistEditor"; // STRING_OK
class CResExistEditor : public CEditor
{
public:
	CResExistEditor(
		_In_ CWnd* pParentWnd,
		ULONG ulPropTag);

private:
	_Check_return_ ULONG HandleChange(UINT nID) override;
};

CResExistEditor::CResExistEditor(
	_In_ CWnd* pParentWnd,
	ULONG ulPropTag) :
	CEditor(pParentWnd, IDS_RESED, IDS_RESEDEXISTPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(EXISTCLASS);

	InitPane(0, TextPane::CreateSingleLinePane(IDS_ULPROPTAG, false));
	SetHex(0, ulPropTag);
	InitPane(1, TextPane::CreateSingleLinePane(IDS_ULPROPTAG, TagToString(ulPropTag, nullptr, false, true), true));
}

_Check_return_ ULONG CResExistEditor::HandleChange(UINT nID)
{
	auto i = CEditor::HandleChange(nID);

	if (0 == i)
	{
		SetStringW(1, TagToString(GetPropTag(0), nullptr, false, true));
	}
	return i;
}

// This class is only invoked by CRestrictEditor. CRestrictEditor always passes an alloc parent.
// So all memory detached from this class is owned by a parent and must not be freed manually
static wstring SUBRESCLASS = L"CResSubResEditor"; // STRING_OK
class CResSubResEditor : public CEditor
{
public:
	CResSubResEditor(
		_In_ CWnd* pParentWnd,
		ULONG ulSubObject,
		_In_ LPSRestriction lpRes,
		_In_ LPVOID lpAllocParent);

	void OnEditAction1() override;
	_Check_return_ LPSRestriction DetachModifiedSRestriction();

private:
	_Check_return_ ULONG HandleChange(UINT nID) override;

	LPVOID m_lpAllocParent;
	LPSRestriction m_lpOldRes;
	LPSRestriction m_lpNewRes;
};

CResSubResEditor::CResSubResEditor(
	_In_ CWnd* pParentWnd,
	ULONG ulSubObject,
	_In_ LPSRestriction lpRes,
	_In_ LPVOID lpAllocParent) :
	CEditor(pParentWnd, IDS_SUBRESED, IDS_RESEDSUBPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL, IDS_ACTIONEDITRES, NULL, NULL)
{
	TRACE_CONSTRUCTOR(SUBRESCLASS);

	m_lpOldRes = lpRes;
	m_lpNewRes = nullptr;
	m_lpAllocParent = lpAllocParent;

	InitPane(0, TextPane::CreateSingleLinePane(IDS_ULSUBOBJECT, false));
	SetHex(0, ulSubObject);
	InitPane(1, TextPane::CreateSingleLinePane(IDS_ULSUBOBJECT, TagToString(ulSubObject, nullptr, false, true), true));

	InitPane(2, TextPane::CreateMultiLinePane(IDS_LPRES, RestrictionToString(lpRes, nullptr), true));
}

_Check_return_ ULONG CResSubResEditor::HandleChange(UINT nID)
{
	auto i = CEditor::HandleChange(nID);

	if (0 == i)
	{
		SetStringW(1, TagToString(GetPropTag(0), nullptr, false, true));
	}

	return i;
}

_Check_return_ LPSRestriction CResSubResEditor::DetachModifiedSRestriction()
{
	auto m_lpRet = m_lpNewRes;
	m_lpNewRes = nullptr;
	return m_lpRet;
}

void CResSubResEditor::OnEditAction1()
{
	auto hRes = S_OK;
	CRestrictEditor ResEdit(
		this,
		m_lpAllocParent,
		m_lpNewRes ? m_lpNewRes : m_lpOldRes);

	WC_H(ResEdit.DisplayDialog());

	if (S_OK == hRes)
	{
		// Since m_lpNewRes was owned by an m_lpAllocParent, we don't free it directly
		m_lpNewRes = ResEdit.DetachModifiedSRestriction();

		SetStringW(2, RestrictionToString(m_lpNewRes, nullptr));
	}
}

// This class is only invoked by CRestrictEditor. CRestrictEditor always passes an alloc parent.
// So all memory detached from this class is owned by a parent and must not be freed manually
static wstring ANDORCLASS = L"CResAndOrEditor"; // STRING_OK
class CResAndOrEditor : public CEditor
{
public:
	CResAndOrEditor(
		_In_ CWnd* pParentWnd,
		_In_ LPSRestriction lpRes,
		_In_ LPVOID lpAllocParent);

	_Check_return_ LPSRestriction DetachModifiedSRestrictionArray();
	_Check_return_ ULONG GetResCount() const;
	_Check_return_ bool DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData) override;

private:
	BOOL OnInitDialog() override;
	void InitListFromRestriction(ULONG ulListNum, _In_ LPSRestriction lpRes) const;
	void OnOK() override;

	LPVOID m_lpAllocParent;
	LPSRestriction m_lpRes;
	LPSRestriction m_lpNewResArray;
	ULONG m_ulNewResCount;
};

CResAndOrEditor::CResAndOrEditor(
	_In_ CWnd* pParentWnd,
	_In_ LPSRestriction lpRes,
	_In_ LPVOID lpAllocParent) :
	CEditor(pParentWnd, IDS_RESED, IDS_RESEDANDORPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(ANDORCLASS);
	m_lpRes = lpRes;
	m_lpNewResArray = nullptr;
	m_ulNewResCount = NULL;
	m_lpAllocParent = lpAllocParent;

	InitPane(0, ListPane::Create(IDS_SUBRESTRICTIONS, false, false, ListEditCallBack(this)));
}

// Used to call functions which need to be called AFTER controls are created
BOOL CResAndOrEditor::OnInitDialog()
{
	auto bRet = CEditor::OnInitDialog();

	InitListFromRestriction(0, m_lpRes);

	UpdateButtons();

	return bRet;
}

_Check_return_ LPSRestriction CResAndOrEditor::DetachModifiedSRestrictionArray()
{
	auto m_lpRet = m_lpNewResArray;
	m_lpNewResArray = nullptr;
	return m_lpRet;
}

_Check_return_ ULONG CResAndOrEditor::GetResCount() const
{
	return m_ulNewResCount;
}

void CResAndOrEditor::InitListFromRestriction(ULONG ulListNum, _In_ LPSRestriction lpRes) const
{
	if (!lpRes) return;

	ClearList(ulListNum);

	InsertColumn(ulListNum, 0, IDS_SHARP);
	SortListData* lpData = nullptr;
	switch (lpRes->rt)
	{
	case RES_AND:
		InsertColumn(ulListNum, 1, IDS_SUBRESTRICTION);

		for (ULONG i = 0; i < lpRes->res.resAnd.cRes; i++)
		{
			lpData = InsertListRow(ulListNum, i, std::to_wstring(i));
			if (lpData)
			{
				lpData->InitializeRes(&lpRes->res.resAnd.lpRes[i]);
				SetListString(ulListNum, i, 1, RestrictionToString(&lpRes->res.resAnd.lpRes[i], nullptr));
			}
		}
		break;
	case RES_OR:
		InsertColumn(ulListNum, 1, IDS_SUBRESTRICTION);

		for (ULONG i = 0; i < lpRes->res.resOr.cRes; i++)
		{
			lpData = InsertListRow(ulListNum, i, std::to_wstring(i));
			if (lpData)
			{
				lpData->InitializeRes(&lpRes->res.resOr.lpRes[i]);
				SetListString(ulListNum, i, 1, RestrictionToString(&lpRes->res.resOr.lpRes[i], nullptr));
			}
		}
		break;
	}
	ResizeList(ulListNum, false);
}

_Check_return_ bool CResAndOrEditor::DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData)
{
	if (!lpData || !lpData->Res()) return false;
	auto hRes = S_OK;

	auto lpSourceRes = lpData->Res()->m_lpNewRes;
	if (!lpSourceRes) lpSourceRes = lpData->Res()->m_lpOldRes;

	CRestrictEditor MyResEditor(
		this,
		m_lpAllocParent,
		lpSourceRes); // pass source res into editor
	WC_H(MyResEditor.DisplayDialog());

	if (S_OK == hRes)
	{
		// Since lpData->data.Res.lpNewRes was owned by an m_lpAllocParent, we don't free it directly
		lpData->Res()->m_lpNewRes = MyResEditor.DetachModifiedSRestriction();
		SetListString(ulListNum, iItem, 1, RestrictionToString(lpData->Res()->m_lpNewRes, nullptr));
		return true;
	}

	return false;
}

// Create our LPSRestriction array from the dialog here
void CResAndOrEditor::OnOK()
{
	CMyDialog::OnOK(); // don't need to call CEditor::OnOK

	LPSRestriction lpNewResArray = nullptr;
	auto ulNewResCount = GetListCount(0);

	if (ulNewResCount > ULONG_MAX / sizeof(SRestriction)) return;
	auto hRes = S_OK;
	EC_H(MAPIAllocateMore(
		sizeof(SRestriction)* ulNewResCount,
		m_lpAllocParent,
		reinterpret_cast<LPVOID*>(&lpNewResArray)));

	if (lpNewResArray)
	{
		for (ULONG i = 0; i < ulNewResCount; i++)
		{
			auto lpData = GetListRowData(0, i);
			if (lpData && lpData->Res())
			{
				if (lpData->Res()->m_lpNewRes)
				{
					memcpy(&lpNewResArray[i], lpData->Res()->m_lpNewRes, sizeof(SRestriction));
					memset(lpData->Res()->m_lpNewRes, 0, sizeof(SRestriction));
				}
				else
				{
					EC_H(HrCopyRestrictionArray(
						lpData->Res()->m_lpOldRes,
						m_lpAllocParent,
						1,
						&lpNewResArray[i]));
				}
			}
		}
		m_ulNewResCount = ulNewResCount;
		m_lpNewResArray = lpNewResArray;
	}
}

// This class is only invoked by CRestrictEditor. CRestrictEditor always passes an alloc parent.
// So all memory detached from this class is owned by a parent and must not be freed manually
static wstring COMMENTCLASS = L"CResCommentEditor"; // STRING_OK
class CResCommentEditor : public CEditor
{
public:
	CResCommentEditor(
		_In_ CWnd* pParentWnd,
		_In_ LPSRestriction lpRes,
		_In_ LPVOID lpAllocParent);

	_Check_return_ LPSRestriction DetachModifiedSRestriction();
	_Check_return_ LPSPropValue DetachModifiedSPropValue();
	_Check_return_ ULONG GetSPropValueCount() const;
	_Check_return_ bool DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData) override;

private:
	void OnEditAction1() override;
	BOOL OnInitDialog() override;
	void InitListFromPropArray(ULONG ulListNum, ULONG cProps, _In_count_(cProps) LPSPropValue lpProps) const;
	_Check_return_ LPSRestriction GetSourceRes() const;
	void OnOK() override;

	LPVOID m_lpAllocParent;
	LPSRestriction m_lpSourceRes;

	LPSRestriction m_lpNewCommentRes;
	LPSPropValue m_lpNewCommentProp;
	ULONG m_ulNewCommentProp;
};

CResCommentEditor::CResCommentEditor(
	_In_ CWnd* pParentWnd,
	_In_ LPSRestriction lpRes,
	_In_ LPVOID lpAllocParent) :
	CEditor(pParentWnd, IDS_COMMENTRESED, IDS_RESEDCOMMENTPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL, IDS_ACTIONEDITRES, NULL, NULL)
{
	TRACE_CONSTRUCTOR(COMMENTCLASS);

	m_lpSourceRes = lpRes;
	m_lpNewCommentRes = nullptr;
	m_lpNewCommentProp = nullptr;
	m_lpAllocParent = lpAllocParent;

	InitPane(0, CollapsibleListPane::Create(IDS_SUBRESTRICTION, false, false, ListEditCallBack(this)));
	InitPane(1, TextPane::CreateMultiLinePane(IDS_RESTRICTIONTEXT, RestrictionToString(m_lpSourceRes->res.resComment.lpRes, nullptr), true));
}

// Used to call functions which need to be called AFTER controls are created
BOOL CResCommentEditor::OnInitDialog()
{
	auto bRet = CEditor::OnInitDialog();

	InitListFromPropArray(0, m_lpSourceRes->res.resComment.cValues, m_lpSourceRes->res.resComment.lpProp);

	UpdateButtons();

	return bRet;
}

_Check_return_ LPSRestriction CResCommentEditor::GetSourceRes() const
{
	if (m_lpNewCommentRes) return m_lpNewCommentRes;
	if (m_lpSourceRes && m_lpSourceRes->res.resComment.lpRes) return m_lpSourceRes->res.resComment.lpRes;
	return nullptr;
}

_Check_return_ LPSRestriction CResCommentEditor::DetachModifiedSRestriction()
{
	auto m_lpRet = m_lpNewCommentRes;
	m_lpNewCommentRes = nullptr;
	return m_lpRet;
}

_Check_return_ LPSPropValue CResCommentEditor::DetachModifiedSPropValue()
{
	auto m_lpRet = m_lpNewCommentProp;
	m_lpNewCommentProp = nullptr;
	return m_lpRet;
}

_Check_return_ ULONG CResCommentEditor::GetSPropValueCount() const
{
	return m_ulNewCommentProp;
}

void CResCommentEditor::InitListFromPropArray(ULONG ulListNum, ULONG cProps, _In_count_(cProps) LPSPropValue lpProps) const
{
	ClearList(ulListNum);

	InsertColumn(ulListNum, 0, IDS_SHARP);
	wstring szProp;
	wstring szAltProp;
	InsertColumn(ulListNum, 1, IDS_PROPERTY);
	InsertColumn(ulListNum, 2, IDS_VALUE);
	InsertColumn(ulListNum, 3, IDS_ALTERNATEVIEW);

	for (ULONG i = 0; i < cProps; i++)
	{
		auto lpData = InsertListRow(ulListNum, i, std::to_wstring(i));
		if (lpData)
		{
			lpData->InitializeComment(&lpProps[i]);
			SetListString(ulListNum, i, 1, TagToString(lpProps[i].ulPropTag, nullptr, false, true));
			InterpretProp(&lpProps[i], &szProp, &szAltProp);
			SetListString(ulListNum, i, 2, szProp);
			SetListString(ulListNum, i, 3, szAltProp);
		}
	}

	ResizeList(ulListNum, false);
}

void CResCommentEditor::OnEditAction1()
{
	auto hRes = S_OK;

	auto lpSourceRes = GetSourceRes();

	CRestrictEditor MyResEditor(
		this,
		m_lpAllocParent,
		lpSourceRes); // pass source res into editor
	WC_H(MyResEditor.DisplayDialog());

	if (S_OK == hRes)
	{
		// Since m_lpNewCommentRes was owned by an m_lpAllocParent, we don't free it directly
		m_lpNewCommentRes = MyResEditor.DetachModifiedSRestriction();
		SetStringW(1, RestrictionToString(m_lpNewCommentRes, nullptr));
	}
}

_Check_return_ bool CResCommentEditor::DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData)
{
	if (!lpData || !lpData->Comment()) return false;
	if (!m_lpAllocParent) return false;
	auto hRes = S_OK;

	auto lpSourceProp = lpData->Comment()->m_lpNewProp;
	if (!lpSourceProp) lpSourceProp = lpData->Comment()->m_lpOldProp;

	SPropValue sProp = { 0 };

	if (!lpSourceProp)
	{
		CEditor MyTag(
			this,
			IDS_TAG,
			IDS_TAGPROMPT,
			true);

		MyTag.InitPane(0, TextPane::CreateSingleLinePane(IDS_TAG, false));

		WC_H(MyTag.DisplayDialog());
		if (S_OK != hRes) return false;
		sProp.ulPropTag = MyTag.GetHex(0);
		lpSourceProp = &sProp;
	}

	WC_H(DisplayPropertyEditor(
		this,
		IDS_PROPEDITOR,
		NULL,
		false,
		m_lpAllocParent,
		NULL,
		NULL,
		false,
		lpSourceProp,
		&lpData->Comment()->m_lpNewProp));

	// Since lpData->data.Comment.lpNewProp was owned by an m_lpAllocParent, we don't free it directly
	if (S_OK == hRes && lpData->Comment()->m_lpNewProp)
	{
		wstring szTmp;
		wstring szAltTmp;
		SetListString(ulListNum, iItem, 1, TagToString(lpData->Comment()->m_lpNewProp->ulPropTag, nullptr, false, true));
		InterpretProp(lpData->Comment()->m_lpNewProp, &szTmp, &szAltTmp);
		SetListString(ulListNum, iItem, 2, szTmp);
		SetListString(ulListNum, iItem, 3, szAltTmp);
		return true;
	}

	return false;
}

void CResCommentEditor::OnOK()
{
	CMyDialog::OnOK(); // don't need to call CEditor::OnOK

	LPSPropValue lpNewCommentProp = nullptr;
	auto ulNewCommentProp = GetListCount(0);

	auto hRes = S_OK;
	if (ulNewCommentProp && ulNewCommentProp < ULONG_MAX / sizeof(SPropValue))
	{
		EC_H(MAPIAllocateMore(
			sizeof(SPropValue)* ulNewCommentProp,
			m_lpAllocParent,
			reinterpret_cast<LPVOID*>(&lpNewCommentProp)));

		if (lpNewCommentProp)
		{
			for (ULONG i = 0; i < ulNewCommentProp; i++)
			{
				auto lpData = GetListRowData(0, i);
				if (lpData && lpData->Comment())
				{
					if (lpData->Comment()->m_lpNewProp)
					{
						EC_H(MyPropCopyMore(
							&lpNewCommentProp[i],
							lpData->Comment()->m_lpNewProp,
							MAPIAllocateMore,
							m_lpAllocParent));
					}
					else
					{
						EC_H(MyPropCopyMore(
							&lpNewCommentProp[i],
							lpData->Comment()->m_lpOldProp,
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
		EC_H(HrCopyRestriction(m_lpSourceRes->res.resComment.lpRes, m_lpAllocParent, &m_lpNewCommentRes));
	}
}

// Note that an alloc parent is passed in to CRestrictEditor. If a parent isn't passed, we allocate one ourselves.
// All other memory allocated in CRestrictEditor is owned by the parent and must not be freed manually
// If we return (detach) memory to a caller, they must MAPIFreeBuffer only if they did not pass in a parent
static wstring CLASS = L"CRestrictEditor"; // STRING_OK
// Create an editor for a restriction
// Takes LPSRestriction lpRes as input
CRestrictEditor::CRestrictEditor(
	_In_ CWnd* pParentWnd,
	_In_opt_ LPVOID lpAllocParent,
	_In_opt_ LPSRestriction lpRes) :
	CEditor(pParentWnd, IDS_RESED, IDS_RESEDPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL, IDS_ACTIONEDITRES, NULL, NULL)
{
	TRACE_CONSTRUCTOR(CLASS);
	auto hRes = S_OK;

	// Not copying the source restriction, but since we're modal it won't matter
	m_lpRes = lpRes;
	m_lpOutputRes = nullptr;
	m_bModified = false;
	if (!lpRes) m_bModified = true; // if we don't have a source, make sure we never look at it

	m_lpAllocParent = lpAllocParent;

	// Set up our output restriction and determine m_lpAllocParent- this will be the basis for all other allocations
	// Allocate base memory:
	if (m_lpAllocParent)
	{
		EC_H(MAPIAllocateMore(
			sizeof(SRestriction),
			m_lpAllocParent,
			reinterpret_cast<LPVOID*>(&m_lpOutputRes)));
	}
	else
	{
		EC_H(MAPIAllocateBuffer(
			sizeof(SRestriction),
			reinterpret_cast<LPVOID*>(&m_lpOutputRes)));

		m_lpAllocParent = m_lpOutputRes;
	}

	if (m_lpOutputRes)
	{
		memset(m_lpOutputRes, 0, sizeof(SRestriction));
		if (m_lpRes) m_lpOutputRes->rt = m_lpRes->rt;
	}

	SetPromptPostFix(AllFlagsToString(flagRestrictionType, true));
	InitPane(0, TextPane::CreateSingleLinePane(IDS_RESTRICTIONTYPE, false)); // type as a number
	InitPane(1, TextPane::CreateSingleLinePane(IDS_RESTRICTIONTYPE, true)); // type as a string (flagRestrictionType)
	InitPane(2, TextPane::CreateMultiLinePane(IDS_RESTRICTIONTEXT, RestrictionToString(GetSourceRes(), nullptr), true));
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
	auto bRet = CEditor::OnInitDialog();

	auto lpSourceRes = GetSourceRes();

	if (lpSourceRes)
	{
		SetHex(0, lpSourceRes->rt);
		SetStringW(1, InterpretFlags(flagRestrictionType, lpSourceRes->rt));
	}

	return bRet;
}

_Check_return_ LPSRestriction CRestrictEditor::GetSourceRes() const
{
	if (m_lpRes && !m_bModified) return m_lpRes;
	return m_lpOutputRes;
}

// Create our LPSRestriction from the dialog here
void CRestrictEditor::OnOK()
{
	CMyDialog::OnOK(); // don't need to call CEditor::OnOK
}

_Check_return_ LPSRestriction CRestrictEditor::DetachModifiedSRestriction()
{
	if (!m_bModified) return nullptr;
	auto m_lpRet = m_lpOutputRes;
	m_lpOutputRes = nullptr;
	return m_lpRet;
}

// CEditor::HandleChange will return the control number of what changed
// so I can react to it if I need to
_Check_return_ ULONG CRestrictEditor::HandleChange(UINT nID)
{
	auto i = CEditor::HandleChange(nID);

	// If the restriction type changed
	if (0 == i)
	{
		if (!m_lpOutputRes) return i;
		auto hRes = S_OK;
		auto ulOldResType = m_lpOutputRes->rt;
		auto ulNewResType = GetHex(i);

		if (ulOldResType == ulNewResType) return i;

		SetStringW(1, InterpretFlags(flagRestrictionType, ulNewResType));

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
				EC_H(MAPIAllocateBuffer(
					sizeof(SRestriction),
					reinterpret_cast<LPVOID*>(&m_lpOutputRes)));

				m_lpAllocParent = m_lpOutputRes;
			}
			else
			{
				// If the pointers are different, m_lpOutputRes was allocated with MAPIAllocateMore
				// Since m_lpOutputRes is owned by m_lpAllocParent, we don't free it directly
				EC_H(MAPIAllocateMore(
					sizeof(SRestriction),
					m_lpAllocParent,
					reinterpret_cast<LPVOID*>(&m_lpOutputRes)));
			}
			memset(m_lpOutputRes, 0, sizeof(SRestriction));
			m_lpOutputRes->rt = ulNewResType;
		}

		SetStringW(2, RestrictionToString(m_lpOutputRes, nullptr));
	}
	return i;
}

void CRestrictEditor::OnEditAction1()
{
	auto hRes = S_OK;
	auto lpSourceRes = GetSourceRes();
	if (!lpSourceRes || !m_lpOutputRes) return;

	switch (lpSourceRes->rt)
	{
	case RES_COMPAREPROPS:
		WC_H(EditCompare(lpSourceRes));
		break;
		// Structures for these two types are identical
	case RES_OR:
	case RES_AND:
		WC_H(EditAndOr(lpSourceRes));
		break;
		// Structures for these two types are identical
	case RES_NOT:
	case RES_COUNT:
		WC_H(EditRestrict(lpSourceRes));
		break;
		// Structures for these two types are identical
	case RES_PROPERTY:
	case RES_CONTENT:
		WC_H(EditCombined(lpSourceRes));
		break;
	case RES_BITMASK:
		WC_H(EditBitmask(lpSourceRes));
		break;
	case RES_SIZE:
		WC_H(EditSize(lpSourceRes));
		break;
	case RES_EXIST:
		WC_H(EditExist(lpSourceRes));
		break;
	case RES_SUBRESTRICTION:
		WC_H(EditSubrestriction(lpSourceRes));
		break;
		// Structures for these two types are identical
	case RES_COMMENT:
	case RES_ANNOTATION:
		WC_H(EditComment(lpSourceRes));
		break;
	}

	if (S_OK == hRes)
	{
		m_bModified = true;
		SetStringW(2, RestrictionToString(m_lpOutputRes, nullptr));
	}
}

HRESULT CRestrictEditor::EditCompare(LPSRestriction lpSourceRes)
{
	auto hRes = S_OK;
	CResCompareEditor MyEditor(
		this,
		lpSourceRes->res.resCompareProps.relop,
		lpSourceRes->res.resCompareProps.ulPropTag1,
		lpSourceRes->res.resCompareProps.ulPropTag2);
	WC_H(MyEditor.DisplayDialog());
	if (S_OK == hRes)
	{
		m_lpOutputRes->rt = lpSourceRes->rt;
		m_lpOutputRes->res.resCompareProps.relop = MyEditor.GetHex(0);
		m_lpOutputRes->res.resCompareProps.ulPropTag1 = MyEditor.GetPropTag(2);
		m_lpOutputRes->res.resCompareProps.ulPropTag2 = MyEditor.GetPropTag(4);
	}

	return hRes;
}

HRESULT CRestrictEditor::EditAndOr(LPSRestriction lpSourceRes)
{
	auto hRes = S_OK;
	CResAndOrEditor MyResEditor(
		this,
		lpSourceRes,
		m_lpAllocParent); // pass source res into editor
	WC_H(MyResEditor.DisplayDialog());

	if (S_OK == hRes)
	{
		m_lpOutputRes->rt = lpSourceRes->rt;
		m_lpOutputRes->res.resAnd.cRes = MyResEditor.GetResCount();

		auto lpNewResArray = MyResEditor.DetachModifiedSRestrictionArray();
		if (lpNewResArray)
		{
			// I can do this because our new memory was allocated from a common parent
			MAPIFreeBuffer(m_lpOutputRes->res.resAnd.lpRes);
			m_lpOutputRes->res.resAnd.lpRes = lpNewResArray;
		}
	}

	return hRes;
}

HRESULT CRestrictEditor::EditRestrict(LPSRestriction lpSourceRes)
{
	auto hRes = S_OK;
	CRestrictEditor MyResEditor(
		this,
		m_lpAllocParent,
		lpSourceRes->res.resNot.lpRes); // pass source res into editor
	WC_H(MyResEditor.DisplayDialog());

	if (S_OK == hRes)
	{
		m_lpOutputRes->rt = lpSourceRes->rt;
		// Since m_lpOutputRes->res.resNot.lpRes was owned by an m_lpAllocParent, we don't free it directly
		m_lpOutputRes->res.resNot.lpRes = MyResEditor.DetachModifiedSRestriction();
	}

	return hRes;
}

HRESULT CRestrictEditor::EditCombined(LPSRestriction lpSourceRes)
{
	auto hRes = S_OK;
	CResCombinedEditor MyEditor(
		this,
		lpSourceRes->rt,
		lpSourceRes->res.resContent.ulFuzzyLevel,
		lpSourceRes->res.resContent.ulPropTag,
		lpSourceRes->res.resContent.lpProp,
		m_lpAllocParent);
	WC_H(MyEditor.DisplayDialog());
	if (S_OK == hRes)
	{
		m_lpOutputRes->rt = lpSourceRes->rt;
		m_lpOutputRes->res.resContent.ulFuzzyLevel = MyEditor.GetHex(0);
		m_lpOutputRes->res.resContent.ulPropTag = MyEditor.GetPropTag(2);

		// Since m_lpOutputRes->res.resContent.lpProp was owned by an m_lpAllocParent, we don't free it directly
		m_lpOutputRes->res.resContent.lpProp = MyEditor.DetachModifiedSPropValue();

		if (!m_lpOutputRes->res.resContent.lpProp)
		{
			// Got a problem here - the relop or fuzzy level was changed, but not the property
			// Need to copy the property from the source Res to the output Res
			EC_H(MAPIAllocateMore(
				sizeof(SPropValue),
				m_lpAllocParent,
				reinterpret_cast<LPVOID*>(&m_lpOutputRes->res.resContent.lpProp)));
			EC_H(MyPropCopyMore(
				m_lpOutputRes->res.resContent.lpProp,
				lpSourceRes->res.resContent.lpProp,
				MAPIAllocateMore,
				m_lpAllocParent));
		}
	}

	return hRes;
}

HRESULT CRestrictEditor::EditBitmask(LPSRestriction lpSourceRes)
{
	auto hRes = S_OK;
	CResBitmaskEditor MyEditor(
		this,
		lpSourceRes->res.resBitMask.relBMR,
		lpSourceRes->res.resBitMask.ulPropTag,
		lpSourceRes->res.resBitMask.ulMask);
	WC_H(MyEditor.DisplayDialog());
	if (S_OK == hRes)
	{
		m_lpOutputRes->rt = lpSourceRes->rt;
		m_lpOutputRes->res.resBitMask.relBMR = MyEditor.GetHex(0);
		m_lpOutputRes->res.resBitMask.ulPropTag = MyEditor.GetPropTag(2);
		m_lpOutputRes->res.resBitMask.ulMask = MyEditor.GetHex(4);
	}

	return hRes;
}

HRESULT CRestrictEditor::EditSize(LPSRestriction lpSourceRes)
{
	auto hRes = S_OK;
	CResSizeEditor MyEditor(
		this,
		lpSourceRes->res.resSize.relop,
		lpSourceRes->res.resSize.ulPropTag,
		lpSourceRes->res.resSize.cb);
	WC_H(MyEditor.DisplayDialog());
	if (S_OK == hRes)
	{
		m_lpOutputRes->rt = lpSourceRes->rt;
		m_lpOutputRes->res.resSize.relop = MyEditor.GetHex(0);
		m_lpOutputRes->res.resSize.ulPropTag = MyEditor.GetPropTag(2);
		m_lpOutputRes->res.resSize.cb = MyEditor.GetHex(4);
	}

	return hRes;
}

HRESULT CRestrictEditor::EditExist(LPSRestriction lpSourceRes)
{
	auto hRes = S_OK;
	CResExistEditor MyEditor(
		this,
		lpSourceRes->res.resExist.ulPropTag);
	WC_H(MyEditor.DisplayDialog());
	if (S_OK == hRes)
	{
		m_lpOutputRes->rt = lpSourceRes->rt;
		m_lpOutputRes->res.resExist.ulPropTag = MyEditor.GetPropTag(0);
		m_lpOutputRes->res.resExist.ulReserved1 = 0;
		m_lpOutputRes->res.resExist.ulReserved2 = 0;
	}

	return hRes;
}

HRESULT CRestrictEditor::EditSubrestriction(LPSRestriction lpSourceRes)
{
	auto hRes = S_OK;
	CResSubResEditor MyEditor(
		this,
		lpSourceRes->res.resSub.ulSubObject,
		lpSourceRes->res.resSub.lpRes,
		m_lpAllocParent);
	WC_H(MyEditor.DisplayDialog());
	if (S_OK == hRes)
	{
		m_lpOutputRes->rt = lpSourceRes->rt;
		m_lpOutputRes->res.resSub.ulSubObject = MyEditor.GetHex(1);

		// Since m_lpOutputRes->res.resSub.lpRes was owned by an m_lpAllocParent, we don't free it directly
		m_lpOutputRes->res.resSub.lpRes = MyEditor.DetachModifiedSRestriction();
	}

	return hRes;
}

HRESULT CRestrictEditor::EditComment(LPSRestriction lpSourceRes)
{
	auto hRes = S_OK;
	CResCommentEditor MyResEditor(
		this,
		lpSourceRes,
		m_lpAllocParent); // pass source res into editor
	WC_H(MyResEditor.DisplayDialog());

	if (S_OK == hRes)
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
	}

	return hRes;
}

// Note that no alloc parent is passed in to CCriteriaEditor. So we're completely responsible for freeing any memory we allocate.
// If we return (detach) memory to a caller, they must MAPIFreeBuffer
static wstring CRITERIACLASS = L"CCriteriaEditor"; // STRING_OK
#define LISTNUM 4
CCriteriaEditor::CCriteriaEditor(
	_In_ CWnd* pParentWnd,
	_In_ LPSRestriction lpRes,
	_In_ LPENTRYLIST lpEntryList,
	ULONG ulSearchState) :
	CEditor(pParentWnd, IDS_CRITERIAEDITOR, IDS_CRITERIAEDITORPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL, IDS_ACTIONEDITRES, NULL, NULL)
{
	TRACE_CONSTRUCTOR(CRITERIACLASS);

	auto hRes = S_OK;

	m_lpSourceRes = lpRes;
	m_lpNewRes = nullptr;
	m_lpSourceEntryList = lpEntryList;

	EC_H(MAPIAllocateBuffer(
		sizeof(SBinaryArray),
		reinterpret_cast<LPVOID*>(&m_lpNewEntryList)));

	m_ulNewSearchFlags = NULL;

	SetPromptPostFix(AllFlagsToString(flagSearchFlag, true));
	InitPane(0, TextPane::CreateSingleLinePane(IDS_SEARCHSTATE, true));
	SetHex(0, ulSearchState);
	auto szFlags = InterpretFlags(flagSearchState, ulSearchState);
	InitPane(1, TextPane::CreateSingleLinePane(IDS_SEARCHSTATE, szFlags, true));
	InitPane(2, TextPane::CreateSingleLinePane(IDS_SEARCHFLAGS, false));
	SetHex(2, 0);
	InitPane(3, TextPane::CreateSingleLinePane(IDS_SEARCHFLAGS, true));
	InitPane(4, CollapsibleListPane::Create(IDS_EIDLIST, false, false, ListEditCallBack(this)));
	InitPane(5, TextPane::CreateMultiLinePane(IDS_RESTRICTIONTEXT, RestrictionToString(m_lpSourceRes, nullptr), true));
}

CCriteriaEditor::~CCriteriaEditor()
{
	// If these structures weren't detached, we need to free them
	MAPIFreeBuffer(m_lpNewEntryList);
	MAPIFreeBuffer(m_lpNewRes);
}

// Used to call functions which need to be called AFTER controls are created
BOOL CCriteriaEditor::OnInitDialog()
{
	auto bRet = CEditor::OnInitDialog();

	InitListFromEntryList(LISTNUM, m_lpSourceEntryList);

	UpdateButtons();

	return bRet;
}

_Check_return_ LPSRestriction CCriteriaEditor::GetSourceRes() const
{
	if (m_lpNewRes) return m_lpNewRes;
	return m_lpSourceRes;
}

_Check_return_ ULONG CCriteriaEditor::HandleChange(UINT nID)
{
	auto i = CEditor::HandleChange(nID);

	if (2 == i)
	{
		SetStringW(3, InterpretFlags(flagSearchFlag, GetHex(i)));
	}

	return i;
}

// Whoever gets this MUST MAPIFreeBuffer
_Check_return_ LPSRestriction CCriteriaEditor::DetachModifiedSRestriction()
{
	auto m_lpRet = m_lpNewRes;
	m_lpNewRes = nullptr;
	return m_lpRet;
}

// Whoever gets this MUST MAPIFreeBuffer
_Check_return_ LPENTRYLIST CCriteriaEditor::DetachModifiedEntryList()
{
	auto m_lpRet = m_lpNewEntryList;
	m_lpNewEntryList = nullptr;
	return m_lpRet;
}

_Check_return_ ULONG CCriteriaEditor::GetSearchFlags() const
{
	return m_ulNewSearchFlags;
}

void CCriteriaEditor::InitListFromEntryList(ULONG ulListNum, _In_ LPENTRYLIST lpEntryList) const
{
	ClearList(ulListNum);

	InsertColumn(ulListNum, 0, IDS_SHARP);
	InsertColumn(ulListNum, 1, IDS_CB);
	InsertColumn(ulListNum, 2, IDS_BINARY);
	InsertColumn(ulListNum, 3, IDS_TEXTVIEW);

	if (lpEntryList)
	{
		for (ULONG i = 0; i < lpEntryList->cValues; i++)
		{
			auto lpData = InsertListRow(ulListNum, i, std::to_wstring(i));
			if (lpData)
			{
				lpData->InitializeBinary(&lpEntryList->lpbin[i]);
			}

			SetListString(ulListNum, i, 1, std::to_wstring(lpEntryList->lpbin[i].cb));
			SetListString(ulListNum, i, 2, BinToHexString(&lpEntryList->lpbin[i], false));
			SetListString(ulListNum, i, 3, BinToTextString(&lpEntryList->lpbin[i], true));
			if (lpData) lpData->bItemFullyLoaded = true;
		}
	}

	ResizeList(ulListNum, false);
}

void CCriteriaEditor::OnEditAction1()
{
	auto hRes = S_OK;

	auto lpSourceRes = GetSourceRes();

	CRestrictEditor MyResEditor(
		this,
		nullptr,
		lpSourceRes); // pass source res into editor
	WC_H(MyResEditor.DisplayDialog());

	if (S_OK == hRes)
	{
		auto lpModRes = MyResEditor.DetachModifiedSRestriction();
		if (lpModRes)
		{
			// We didn't pass an alloc parent to CRestrictEditor, so we must free what came back
			MAPIFreeBuffer(m_lpNewRes);
			m_lpNewRes = lpModRes;
			SetStringW(5, RestrictionToString(m_lpNewRes, nullptr));
		}
	}
}

_Check_return_ bool CCriteriaEditor::DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData)
{
	if (!lpData && !lpData->Binary()) return false;
	auto hRes = S_OK;
	wstring szTmp;

	CEditor BinEdit(
		this,
		IDS_EIDEDITOR,
		IDS_EIDEDITORPROMPT,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	LPSBinary lpSourcebin = nullptr;
	if (lpData->Binary()->m_OldBin.lpb)
	{
		lpSourcebin = &lpData->Binary()->m_OldBin;
	}
	else
	{
		lpSourcebin = &lpData->Binary()->m_NewBin;
	}

	BinEdit.InitPane(0, TextPane::CreateSingleLinePane(IDS_EID, BinToHexString(lpSourcebin, false), false));

	WC_H(BinEdit.DisplayDialog());
	if (S_OK == hRes)
	{
		auto bin = HexStringToBin(BinEdit.GetStringW(0));
		lpData->Binary()->m_NewBin.lpb = ByteVectorToMAPI(bin, m_lpNewEntryList);
		if (lpData->Binary()->m_NewBin.lpb)
		{
			lpData->Binary()->m_NewBin.cb = static_cast<ULONG>(bin.size());
			szTmp = std::to_wstring(lpData->Binary()->m_NewBin.cb);
			SetListString(ulListNum, iItem, 1, szTmp);
			SetListString(ulListNum, iItem, 2, BinToHexString(&lpData->Binary()->m_NewBin, false));
			SetListString(ulListNum, iItem, 3, BinToTextString(&lpData->Binary()->m_NewBin, true));
			return true;
		}
	}

	return false;
}

void CCriteriaEditor::OnOK()
{
	CMyDialog::OnOK(); // don't need to call CEditor::OnOK
	auto ulValues = GetListCount(LISTNUM);

	auto hRes = S_OK;

	if (m_lpNewEntryList && ulValues < ULONG_MAX / sizeof(SBinary))
	{
		m_lpNewEntryList->cValues = ulValues;
		EC_H(MAPIAllocateMore(
			m_lpNewEntryList->cValues * sizeof(SBinary),
			m_lpNewEntryList,
			reinterpret_cast<LPVOID*>(&m_lpNewEntryList->lpbin)));

		for (ULONG i = 0; i < m_lpNewEntryList->cValues; i++)
		{
			auto lpData = GetListRowData(LISTNUM, i);
			if (lpData && lpData->Binary())
			{
				if (lpData->Binary()->m_NewBin.lpb)
				{
					m_lpNewEntryList->lpbin[i].cb = lpData->Binary()->m_NewBin.cb;
					m_lpNewEntryList->lpbin[i].lpb = lpData->Binary()->m_NewBin.lpb;
					// clean out the source
					lpData->Binary()->m_OldBin.lpb = nullptr;
				}
				else
				{
					m_lpNewEntryList->lpbin[i].cb = lpData->Binary()->m_OldBin.cb;
					EC_H(MAPIAllocateMore(
						m_lpNewEntryList->lpbin[i].cb,
						m_lpNewEntryList,
						reinterpret_cast<LPVOID*>(&m_lpNewEntryList->lpbin[i].lpb)));

					memcpy(m_lpNewEntryList->lpbin[i].lpb, lpData->Binary()->m_OldBin.lpb, m_lpNewEntryList->lpbin[i].cb);
				}
			}
		}
	}

	if (!m_lpNewRes && m_lpSourceRes)
	{
		EC_H(HrCopyRestriction(m_lpSourceRes, NULL, &m_lpNewRes))
	}

	m_ulNewSearchFlags = GetHex(2);
}