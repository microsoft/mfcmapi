// RestrictEditor.cpp : implementation file
//

#include "stdafx.h"
#include "RestrictEditor.h"
#include "PropertyEditor.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"
#include "MAPIFunctions.h"
#include "ImportProcs.h"
#include "ExtraPropTags.h"

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
	_Check_return_ ULONG HandleChange(UINT nID);
};

CResCompareEditor::CResCompareEditor(
	_In_ CWnd* pParentWnd,
	ULONG ulRelop,
	ULONG ulPropTag1,
	ULONG ulPropTag2) :
	CEditor(pParentWnd, IDS_RESED, IDS_RESEDCOMPPROMPT, 6, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(COMPCLASS);

	SetPromptPostFix(AllFlagsToString(flagRelop, false));
	InitPane(0, CreateSingleLinePane(IDS_RELOP, NULL, false));
	SetHex(0, ulRelop);
	wstring szFlags = InterpretFlags(flagRelop, ulRelop);
	InitPane(1, CreateSingleLinePaneW(IDS_RELOP, szFlags.c_str(), true));
	InitPane(2, CreateSingleLinePane(IDS_ULPROPTAG1, NULL, false));
	SetHex(2, ulPropTag1);
	InitPane(3, CreateSingleLinePaneW(IDS_ULPROPTAG1, TagToString(ulPropTag1, NULL, false, true).c_str(), true));

	InitPane(4, CreateSingleLinePane(IDS_ULPROPTAG2, NULL, false));
	SetHex(4, ulPropTag2);
	InitPane(5, CreateSingleLinePaneW(IDS_ULPROPTAG1, TagToString(ulPropTag2, NULL, false, true).c_str(), true));
} // CResCompareEditor::CResCompareEditor

_Check_return_ ULONG CResCompareEditor::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	if (0 == i)
	{
		wstring szFlags = InterpretFlags(flagRelop, GetHexUseControl(0));
		SetStringW(1, szFlags.c_str());
	}
	else if (2 == i)
	{
		SetStringW(3, TagToString(GetPropTagUseControl(2), NULL, false, true).c_str());
	}
	else if (4 == i)
	{
		SetStringW(5, TagToString(GetPropTagUseControl(4), NULL, false, true).c_str());
	}
	return i;
} // CResCompareEditor::HandleChange

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

	void OnEditAction1();
	_Check_return_ LPSPropValue DetachModifiedSPropValue();

private:
	_Check_return_ ULONG HandleChange(UINT nID);

	ULONG			m_ulResType;
	LPVOID			m_lpAllocParent;
	LPSPropValue	m_lpOldProp;
	LPSPropValue	m_lpNewProp;
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
		8,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL,
		IDS_ACTIONEDITPROP,
		NULL,
		NULL)
{
	TRACE_CONSTRUCTOR(CONTENTCLASS);

	m_ulResType = ulResType;
	m_lpOldProp = lpProp;
	m_lpNewProp = NULL;
	m_lpAllocParent = lpAllocParent;

	wstring szFlags;

	if (RES_CONTENT == m_ulResType)
	{
		SetPromptPostFix(AllFlagsToString(flagFuzzyLevel, true));

		InitPane(0, CreateSingleLinePane(IDS_ULFUZZYLEVEL, NULL, false));
		SetHex(0, ulCompare);
		szFlags = InterpretFlags(flagFuzzyLevel, ulCompare);
		InitPane(1, CreateSingleLinePaneW(IDS_ULFUZZYLEVEL, szFlags.c_str(), true));
	}
	else if (RES_PROPERTY == m_ulResType)
	{
		SetPromptPostFix(AllFlagsToString(flagRelop, false));
		InitPane(0, CreateSingleLinePane(IDS_RELOP, NULL, false));
		SetHex(0, ulCompare);
		szFlags = InterpretFlags(flagRelop, ulCompare);
		InitPane(1, CreateSingleLinePaneW(IDS_RELOP, szFlags.c_str(), true));
	}

	InitPane(2, CreateSingleLinePane(IDS_ULPROPTAG, NULL, false));
	SetHex(2, ulPropTag);
	InitPane(3, CreateSingleLinePaneW(IDS_ULPROPTAG, TagToString(ulPropTag, NULL, false, true).c_str(), true));

	InitPane(4, CreateSingleLinePane(IDS_LPPROPULPROPTAG, NULL, false));
	if (lpProp) SetHex(4, lpProp->ulPropTag);
	InitPane(5, CreateSingleLinePaneW(IDS_LPPROPULPROPTAG, lpProp ? TagToString(lpProp->ulPropTag, NULL, false, true).c_str() : NULL, true));

	wstring szProp;
	wstring szAltProp;
	if (lpProp) InterpretProp(lpProp, &szProp, &szAltProp);
	InitPane(6, CreateMultiLinePaneW(IDS_LPPROP, szProp.c_str(), true));
	InitPane(7, CreateMultiLinePaneW(IDS_LPPROPALTVIEW, szAltProp.c_str(), true));
}

_Check_return_ ULONG CResCombinedEditor::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	wstring szFlags;
	if (0 == i)
	{
		if (RES_CONTENT == m_ulResType)
		{
			szFlags = InterpretFlags(flagFuzzyLevel, GetHexUseControl(0));
			SetStringW(1, szFlags.c_str());
		}
		else if (RES_PROPERTY == m_ulResType)
		{
			szFlags = InterpretFlags(flagRelop, GetHexUseControl(0));
			SetStringW(1, szFlags.c_str());
		}
	}
	else if (2 == i)
	{
		SetStringW(3, TagToString(GetPropTagUseControl(2), NULL, false, true).c_str());
	}
	else if (4 == i)
	{
		SetStringW(5, TagToString(GetPropTagUseControl(4), NULL, false, true).c_str());
		m_lpOldProp = NULL;
		m_lpNewProp = NULL;
		SetString(6, NULL);
		SetString(7, NULL);
	}
	return i;
} // CResCombinedEditor::HandleChange

_Check_return_ LPSPropValue CResCombinedEditor::DetachModifiedSPropValue()
{
	LPSPropValue m_lpRet = m_lpNewProp;
	m_lpNewProp = NULL;
	return m_lpRet;
} // CResCombinedEditor::DetachModifiedSPropValue

void CResCombinedEditor::OnEditAction1()
{
	if (!m_lpAllocParent) return;

	HRESULT hRes = S_OK;

	LPSPropValue lpEditProp = m_lpOldProp;
	LPSPropValue lpOutProp = NULL;
	if (m_lpNewProp) lpEditProp = m_lpNewProp;

	WC_H(DisplayPropertyEditor(
		this,
		IDS_PROPEDITOR,
		NULL,
		false,
		m_lpAllocParent,
		NULL,
		GetPropTagUseControl(4),
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
		SetStringW(6, szProp.c_str());
		SetStringW(7, szAltProp.c_str());
	}
} // CResCombinedEditor::OnEditAction1

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
	_Check_return_ ULONG HandleChange(UINT nID);
};

CResBitmaskEditor::CResBitmaskEditor(
	_In_ CWnd* pParentWnd,
	ULONG relBMR,
	ULONG ulPropTag,
	ULONG ulMask) :
	CEditor(pParentWnd, IDS_RESED, IDS_RESEDBITPROMPT, 5, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(BITMASKCLASS);

	SetPromptPostFix(AllFlagsToString(flagBitmask, false));
	InitPane(0, CreateSingleLinePane(IDS_RELBMR, NULL, false));
	SetHex(0, relBMR);
	wstring szFlags = InterpretFlags(flagBitmask, relBMR);
	InitPane(1, CreateSingleLinePaneW(IDS_RELBMR, szFlags.c_str(), true));
	InitPane(2, CreateSingleLinePane(IDS_ULPROPTAG, NULL, false));
	SetHex(2, ulPropTag);
	InitPane(3, CreateSingleLinePaneW(IDS_ULPROPTAG, TagToString(ulPropTag, NULL, false, true).c_str(), true));

	InitPane(4, CreateSingleLinePane(IDS_MASK, NULL, false));
	SetHex(4, ulMask);
} // CResBitmaskEditor::CResBitmaskEditor

_Check_return_ ULONG CResBitmaskEditor::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	if (0 == i)
	{
		wstring szFlags = InterpretFlags(flagBitmask, GetHexUseControl(0));
		SetStringW(1, szFlags.c_str());
	}
	else if (2 == i)
	{
		SetStringW(3, TagToString(GetPropTagUseControl(2), NULL, false, true).c_str());
	}
	return i;
} // CResBitmaskEditor::HandleChange

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
	_Check_return_ ULONG HandleChange(UINT nID);
};

CResSizeEditor::CResSizeEditor(
	_In_ CWnd* pParentWnd,
	ULONG relop,
	ULONG ulPropTag,
	ULONG cb) :
	CEditor(pParentWnd, IDS_RESED, IDS_RESEDSIZEPROMPT, 5, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(SIZECLASS);

	SetPromptPostFix(AllFlagsToString(flagRelop, false));
	InitPane(0, CreateSingleLinePane(IDS_RELOP, NULL, false));
	SetHex(0, relop);
	wstring szFlags = InterpretFlags(flagRelop, relop);
	InitPane(1, CreateSingleLinePaneW(IDS_RELOP, szFlags.c_str(), true));

	InitPane(2, CreateSingleLinePane(IDS_ULPROPTAG, NULL, false));
	SetHex(2, ulPropTag);
	InitPane(3, CreateSingleLinePaneW(IDS_ULPROPTAG, TagToString(ulPropTag, NULL, false, true).c_str(), true));

	InitPane(4, CreateSingleLinePane(IDS_CB, NULL, false));
	SetHex(4, cb);
} // CResSizeEditor::CResSizeEditor

_Check_return_ ULONG CResSizeEditor::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	if (0 == i)
	{
		wstring szFlags = InterpretFlags(flagRelop, GetHexUseControl(0));
		SetStringW(1, szFlags.c_str());
	}
	else if (2 == i)
	{
		SetStringW(3, TagToString(GetPropTagUseControl(2), NULL, false, true).c_str());
	}
	return i;
} // CResSizeEditor::HandleChange

static wstring EXISTCLASS = L"CResExistEditor"; // STRING_OK
class CResExistEditor : public CEditor
{
public:
	CResExistEditor(
		_In_ CWnd* pParentWnd,
		ULONG ulPropTag);

private:
	_Check_return_ ULONG HandleChange(UINT nID);
};

CResExistEditor::CResExistEditor(
	_In_ CWnd* pParentWnd,
	ULONG ulPropTag) :
	CEditor(pParentWnd, IDS_RESED, IDS_RESEDEXISTPROMPT, 2, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(EXISTCLASS);

	InitPane(0, CreateSingleLinePane(IDS_ULPROPTAG, NULL, false));
	SetHex(0, ulPropTag);
	InitPane(1, CreateSingleLinePaneW(IDS_ULPROPTAG, TagToString(ulPropTag, NULL, false, true).c_str(), true));
} // CResExistEditor::CResExistEditor

_Check_return_ ULONG CResExistEditor::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	if (0 == i)
	{
		SetStringW(1, TagToString(GetPropTagUseControl(0), NULL, false, true).c_str());
	}
	return i;
} // CResExistEditor::HandleChange

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

	void OnEditAction1();
	_Check_return_ LPSRestriction DetachModifiedSRestriction();

private:
	_Check_return_ ULONG HandleChange(UINT nID);

	LPVOID			m_lpAllocParent;
	LPSRestriction	m_lpOldRes;
	LPSRestriction	m_lpNewRes;
};

CResSubResEditor::CResSubResEditor(
	_In_ CWnd* pParentWnd,
	ULONG ulSubObject,
	_In_ LPSRestriction lpRes,
	_In_ LPVOID lpAllocParent) :
	CEditor(pParentWnd, IDS_SUBRESED, IDS_RESEDSUBPROMPT, 3, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL, IDS_ACTIONEDITRES, NULL, NULL)
{
	TRACE_CONSTRUCTOR(SUBRESCLASS);

	m_lpOldRes = lpRes;
	m_lpNewRes = NULL;
	m_lpAllocParent = lpAllocParent;

	InitPane(0, CreateSingleLinePane(IDS_ULSUBOBJECT, NULL, false));
	SetHex(0, ulSubObject);
	InitPane(1, CreateSingleLinePaneW(IDS_ULSUBOBJECT, TagToString(ulSubObject, NULL, false, true).c_str(), true));

	InitPane(2, CreateMultiLinePaneW(IDS_LPRES, RestrictionToString(lpRes, NULL).c_str(), true));
} // CResSubResEditor::CResSubResEditor

_Check_return_ ULONG CResSubResEditor::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	if (0 == i)
	{
		SetStringW(1, TagToString(GetPropTagUseControl(0), NULL, false, true).c_str());
	}
	return i;
} // CResSubResEditor::HandleChange

_Check_return_ LPSRestriction CResSubResEditor::DetachModifiedSRestriction()
{
	LPSRestriction m_lpRet = m_lpNewRes;
	m_lpNewRes = NULL;
	return m_lpRet;
} // CResSubResEditor::DetachModifiedSRestriction

void CResSubResEditor::OnEditAction1()
{
	HRESULT hRes = S_OK;
	CRestrictEditor ResEdit(
		this,
		m_lpAllocParent,
		m_lpNewRes ? m_lpNewRes : m_lpOldRes);

	WC_H(ResEdit.DisplayDialog());

	if (S_OK == hRes)
	{
		// Since m_lpNewRes was owned by an m_lpAllocParent, we don't free it directly
		m_lpNewRes = ResEdit.DetachModifiedSRestriction();

		SetStringW(2, RestrictionToString(m_lpNewRes, NULL).c_str());
	}
} // CResSubResEditor::OnEditAction1

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

	_Check_return_ LPSRestriction  DetachModifiedSRestrictionArray();
	_Check_return_ ULONG			GetResCount();

private:
	_Check_return_ bool DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData);
	BOOL OnInitDialog();
	void InitListFromRestriction(ULONG ulListNum, _In_ LPSRestriction lpRes);
	void OnOK();

	LPVOID			m_lpAllocParent;
	LPSRestriction	m_lpRes;
	LPSRestriction	m_lpNewResArray;
	ULONG			m_ulNewResCount;
};

CResAndOrEditor::CResAndOrEditor(
	_In_ CWnd* pParentWnd,
	_In_ LPSRestriction lpRes,
	_In_ LPVOID lpAllocParent) :
	CEditor(pParentWnd, IDS_RESED, IDS_RESEDANDORPROMPT, 1, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(ANDORCLASS);
	m_lpRes = lpRes;
	m_lpNewResArray = NULL;
	m_ulNewResCount = NULL;
	m_lpAllocParent = lpAllocParent;

	InitPane(0, CreateListPane(IDS_SUBRESTRICTIONS, false, false, this));
} // CResAndOrEditor::CResAndOrEditor

// Used to call functions which need to be called AFTER controls are created
BOOL CResAndOrEditor::OnInitDialog()
{
	BOOL bRet = CEditor::OnInitDialog();

	InitListFromRestriction(0, m_lpRes);

	UpdateListButtons();

	return bRet;
} // CResAndOrEditor::OnInitDialog

_Check_return_ LPSRestriction CResAndOrEditor::DetachModifiedSRestrictionArray()
{
	LPSRestriction m_lpRet = m_lpNewResArray;
	m_lpNewResArray = NULL;
	return m_lpRet;
} // CResAndOrEditor::DetachModifiedSRestrictionArray

_Check_return_ ULONG CResAndOrEditor::GetResCount()
{
	return m_ulNewResCount;
} // CResAndOrEditor::GetResCount

void CResAndOrEditor::InitListFromRestriction(ULONG ulListNum, _In_ LPSRestriction lpRes)
{
	if (!IsValidList(ulListNum)) return;
	if (!lpRes) return;

	ClearList(ulListNum);

	InsertColumn(ulListNum, 0, IDS_SHARP);
	CString szTmp;
	CString szAltTmp;
	SortListData* lpData = NULL;
	ULONG i = 0;
	switch (lpRes->rt)
	{
	case RES_AND:
		InsertColumn(ulListNum, 1, IDS_SUBRESTRICTION);

		for (i = 0; i < lpRes->res.resAnd.cRes; i++)
		{
			szTmp.Format(_T("%u"), i); // STRING_OK
			lpData = InsertListRow(ulListNum, i, szTmp);
			if (lpData)
			{
				lpData->ulSortDataType = SORTLIST_RES;
				lpData->data.Res.lpOldRes = &lpRes->res.resAnd.lpRes[i]; // save off address of source restriction
				SetListStringW(ulListNum, i, 1, RestrictionToString(lpData->data.Res.lpOldRes, NULL).c_str());
				lpData->bItemFullyLoaded = true;
			}
		}
		break;
	case RES_OR:
		InsertColumn(ulListNum, 1, IDS_SUBRESTRICTION);

		for (i = 0; i < lpRes->res.resOr.cRes; i++)
		{
			szTmp.Format(_T("%u"), i); // STRING_OK
			lpData = InsertListRow(ulListNum, i, szTmp);
			if (lpData)
			{
				lpData->ulSortDataType = SORTLIST_RES;
				lpData->data.Res.lpOldRes = &lpRes->res.resOr.lpRes[i]; // save off address of source restriction
				SetListStringW(ulListNum, i, 1, RestrictionToString(lpData->data.Res.lpOldRes, NULL).c_str());
				lpData->bItemFullyLoaded = true;
			}
		}
		break;
	}
	ResizeList(ulListNum, false);
} // CResAndOrEditor::InitListFromRestriction

_Check_return_ bool CResAndOrEditor::DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData)
{
	if (!lpData) return false;
	if (!IsValidList(ulListNum)) return false;
	HRESULT hRes = S_OK;

	LPSRestriction lpSourceRes = NULL;

	lpSourceRes = lpData->data.Res.lpNewRes;
	if (!lpSourceRes) lpSourceRes = lpData->data.Res.lpOldRes;

	CRestrictEditor MyResEditor(
		this,
		m_lpAllocParent,
		lpSourceRes); // pass source res into editor
	WC_H(MyResEditor.DisplayDialog());

	if (S_OK == hRes)
	{
		// Since lpData->data.Res.lpNewRes was owned by an m_lpAllocParent, we don't free it directly
		lpData->data.Res.lpNewRes = MyResEditor.DetachModifiedSRestriction();
		SetListStringW(ulListNum, iItem, 1, RestrictionToString(lpData->data.Res.lpNewRes, NULL).c_str());
		return true;
	}
	return false;
} // CResAndOrEditor::DoListEdit

// Create our LPSRestriction array from the dialog here
void CResAndOrEditor::OnOK()
{
	CMyDialog::OnOK(); // don't need to call CEditor::OnOK

	if (!IsValidList(0)) return;
	LPSRestriction  lpNewResArray = NULL;
	ULONG ulNewResCount = GetListCount(0);

	if (ulNewResCount > ULONG_MAX / sizeof(SRestriction)) return;
	HRESULT hRes = S_OK;
	EC_H(MAPIAllocateMore(
		sizeof(SRestriction)* ulNewResCount,
		m_lpAllocParent,
		(LPVOID*)&lpNewResArray));

	if (lpNewResArray)
	{
		ULONG i = 0;
		for (i = 0; i < ulNewResCount; i++)
		{
			SortListData* lpData = GetListRowData(0, i);
			if (lpData)
			{
				if (lpData->data.Res.lpNewRes)
				{
					memcpy(&lpNewResArray[i], lpData->data.Res.lpNewRes, sizeof(SRestriction));
					memset(lpData->data.Res.lpNewRes, 0, sizeof(SRestriction));
				}
				else
				{
					EC_H(HrCopyRestrictionArray(
						lpData->data.Res.lpOldRes,
						m_lpAllocParent,
						1,
						&lpNewResArray[i]));
				}
			}
		}
		m_ulNewResCount = ulNewResCount;
		m_lpNewResArray = lpNewResArray;
	}
} // CResAndOrEditor::OnOK

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
	_Check_return_ LPSPropValue   DetachModifiedSPropValue();
	_Check_return_ ULONG          GetSPropValueCount();

private:
	void           OnEditAction1();
	_Check_return_ bool           DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData);
	BOOL           OnInitDialog();
	void           InitListFromPropArray(ULONG ulListNum, ULONG cProps, _In_count_(cProps) LPSPropValue lpProps);
	_Check_return_ LPSRestriction GetSourceRes();
	void           OnOK();

	LPVOID         m_lpAllocParent;
	LPSRestriction m_lpSourceRes;

	LPSRestriction m_lpNewCommentRes;
	LPSPropValue   m_lpNewCommentProp;
	ULONG          m_ulNewCommentProp;
};

CResCommentEditor::CResCommentEditor(
	_In_ CWnd* pParentWnd,
	_In_ LPSRestriction lpRes,
	_In_ LPVOID lpAllocParent) :
	CEditor(pParentWnd, IDS_COMMENTRESED, IDS_RESEDCOMMENTPROMPT, 2, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL, IDS_ACTIONEDITRES, NULL, NULL)
{
	TRACE_CONSTRUCTOR(COMMENTCLASS);

	m_lpSourceRes = lpRes;
	m_lpNewCommentRes = NULL;
	m_lpNewCommentProp = NULL;
	m_lpAllocParent = lpAllocParent;

	InitPane(0, CreateListPane(IDS_SUBRESTRICTION, false, false, this));
	InitPane(1, CreateMultiLinePaneW(IDS_RESTRICTIONTEXT, RestrictionToString(m_lpSourceRes->res.resComment.lpRes, NULL).c_str(), true));
} // CResCommentEditor::CResCommentEditor

// Used to call functions which need to be called AFTER controls are created
BOOL CResCommentEditor::OnInitDialog()
{
	BOOL bRet = CEditor::OnInitDialog();

	InitListFromPropArray(0, m_lpSourceRes->res.resComment.cValues, m_lpSourceRes->res.resComment.lpProp);

	UpdateListButtons();

	return bRet;
} // CResCommentEditor::OnInitDialog

_Check_return_ LPSRestriction CResCommentEditor::GetSourceRes()
{
	if (m_lpNewCommentRes) return m_lpNewCommentRes;
	if (m_lpSourceRes && m_lpSourceRes->res.resComment.lpRes) return m_lpSourceRes->res.resComment.lpRes;
	return NULL;
} // CResCommentEditor::GetSourceRe

_Check_return_ LPSRestriction CResCommentEditor::DetachModifiedSRestriction()
{
	LPSRestriction m_lpRet = m_lpNewCommentRes;
	m_lpNewCommentRes = NULL;
	return m_lpRet;
} // CResCommentEditor::DetachModifiedSRestriction

_Check_return_ LPSPropValue CResCommentEditor::DetachModifiedSPropValue()
{
	LPSPropValue m_lpRet = m_lpNewCommentProp;
	m_lpNewCommentProp = NULL;
	return m_lpRet;
} // CResCommentEditor::DetachModifiedSPropValue

_Check_return_ ULONG CResCommentEditor::GetSPropValueCount()
{
	return m_ulNewCommentProp;
} // CResCommentEditor::GetSPropValueCount

void CResCommentEditor::InitListFromPropArray(ULONG ulListNum, ULONG cProps, _In_count_(cProps) LPSPropValue lpProps)
{
	if (!IsValidList(ulListNum)) return;

	ClearList(ulListNum);

	InsertColumn(ulListNum, 0, IDS_SHARP);
	CString szTmp;
	wstring szProp;
	wstring szAltProp;
	SortListData* lpData = NULL;
	ULONG i = 0;
	InsertColumn(ulListNum, 1, IDS_PROPERTY);
	InsertColumn(ulListNum, 2, IDS_VALUE);
	InsertColumn(ulListNum, 3, IDS_ALTERNATEVIEW);

	for (i = 0; i < cProps; i++)
	{
		szTmp.Format(_T("%u"), i); // STRING_OK
		lpData = InsertListRow(ulListNum, i, szTmp);
		if (lpData)
		{
			lpData->ulSortDataType = SORTLIST_COMMENT;
			lpData->data.Comment.lpOldProp = &lpProps[i];
			SetListStringW(ulListNum, i, 1, TagToString(lpProps[i].ulPropTag, NULL, false, true).c_str());
			InterpretProp(&lpProps[i], &szProp, &szAltProp);
			SetListStringW(ulListNum, i, 2, szProp.c_str());
			SetListStringW(ulListNum, i, 3, szAltProp.c_str());
			lpData->bItemFullyLoaded = true;
		}
	}
	ResizeList(ulListNum, false);
} // CResCommentEditor::InitListFromPropArray

void CResCommentEditor::OnEditAction1()
{
	HRESULT hRes = S_OK;

	LPSRestriction lpSourceRes = GetSourceRes();

	CRestrictEditor MyResEditor(
		this,
		m_lpAllocParent,
		lpSourceRes); // pass source res into editor
	WC_H(MyResEditor.DisplayDialog());

	if (S_OK == hRes)
	{
		// Since m_lpNewCommentRes was owned by an m_lpAllocParent, we don't free it directly
		m_lpNewCommentRes = MyResEditor.DetachModifiedSRestriction();
		SetStringW(1, RestrictionToString(m_lpNewCommentRes, NULL).c_str());
	}
} // CResCommentEditor::OnEditAction1

_Check_return_ bool CResCommentEditor::DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData)
{
	if (!lpData) return false;
	if (!IsValidList(ulListNum)) return false;
	if (!m_lpAllocParent) return false;
	HRESULT hRes = S_OK;

	LPSPropValue lpSourceProp = lpData->data.Comment.lpNewProp;
	if (!lpSourceProp) lpSourceProp = lpData->data.Comment.lpOldProp;

	SPropValue sProp = { 0 };

	if (!lpSourceProp)
	{
		CEditor MyTag(
			this,
			IDS_TAG,
			IDS_TAGPROMPT,
			1,
			true);

		MyTag.InitPane(0, CreateSingleLinePane(IDS_TAG, NULL, false));

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
		&lpData->data.Comment.lpNewProp));

	// Since lpData->data.Comment.lpNewProp was owned by an m_lpAllocParent, we don't free it directly
	if (S_OK == hRes && lpData->data.Comment.lpNewProp)
	{
		wstring szTmp;
		wstring szAltTmp;
		SetListStringW(ulListNum, iItem, 1, TagToString(lpData->data.Comment.lpNewProp->ulPropTag, NULL, false, true).c_str());
		InterpretProp(lpData->data.Comment.lpNewProp, &szTmp, &szAltTmp);
		SetListStringW(ulListNum, iItem, 2, szTmp.c_str());
		SetListStringW(ulListNum, iItem, 3, szAltTmp.c_str());
		return true;
	}

	return false;
} // CResCommentEditor::DoListEdit

void CResCommentEditor::OnOK()
{
	CMyDialog::OnOK(); // don't need to call CEditor::OnOK
	if (!IsValidList(0)) return;

	LPSPropValue lpNewCommentProp = NULL;
	ULONG ulNewCommentProp = GetListCount(0);

	HRESULT hRes = S_OK;
	if (ulNewCommentProp && ulNewCommentProp < ULONG_MAX / sizeof(SPropValue))
	{
		EC_H(MAPIAllocateMore(
			sizeof(SPropValue)* ulNewCommentProp,
			m_lpAllocParent,
			(LPVOID*)&lpNewCommentProp));

		if (lpNewCommentProp)
		{
			ULONG i = 0;
			for (i = 0; i < ulNewCommentProp; i++)
			{
				SortListData* lpData = GetListRowData(0, i);
				if (lpData)
				{
					if (lpData->data.Comment.lpNewProp)
					{
						EC_H(MyPropCopyMore(
							&lpNewCommentProp[i],
							lpData->data.Comment.lpNewProp,
							MAPIAllocateMore,
							m_lpAllocParent));
					}
					else
					{
						EC_H(MyPropCopyMore(
							&lpNewCommentProp[i],
							lpData->data.Comment.lpOldProp,
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
} // CResCommentEditor::OnOK

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
	CEditor(pParentWnd, IDS_RESED, IDS_RESEDPROMPT, 3, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL, IDS_ACTIONEDITRES, NULL, NULL)
{
	TRACE_CONSTRUCTOR(CLASS);
	HRESULT hRes = S_OK;

	// Not copying the source restriction, but since we're modal it won't matter
	m_lpRes = lpRes;
	m_lpOutputRes = NULL;
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
			(LPVOID*)&m_lpOutputRes));
	}
	else
	{
		EC_H(MAPIAllocateBuffer(
			sizeof(SRestriction),
			(LPVOID*)&m_lpOutputRes));

		m_lpAllocParent = m_lpOutputRes;
	}

	if (m_lpOutputRes)
	{
		memset(m_lpOutputRes, 0, sizeof(SRestriction));
		if (m_lpRes) m_lpOutputRes->rt = m_lpRes->rt;
	}

	SetPromptPostFix(AllFlagsToString(flagRestrictionType, true));
	InitPane(0, CreateSingleLinePane(IDS_RESTRICTIONTYPE, NULL, false)); // type as a number
	InitPane(1, CreateSingleLinePane(IDS_RESTRICTIONTYPE, NULL, true)); // type as a string (flagRestrictionType)
	InitPane(2, CreateMultiLinePaneW(IDS_RESTRICTIONTEXT, RestrictionToString(GetSourceRes(), NULL).c_str(), true));
} // CRestrictEditor::CRestrictEditor

CRestrictEditor::~CRestrictEditor()
{
	TRACE_DESTRUCTOR(CLASS);
	// Only if we self allocated m_lpOutputRes and did not detach it can we free it
	if (m_lpAllocParent == m_lpOutputRes)
	{
		MAPIFreeBuffer(m_lpOutputRes);
	}
} // CRestrictEditor::~CRestrictEditor

// Used to call functions which need to be called AFTER controls are created
BOOL CRestrictEditor::OnInitDialog()
{
	BOOL bRet = CEditor::OnInitDialog();

	LPSRestriction lpSourceRes = GetSourceRes();

	if (lpSourceRes)
	{
		SetHex(0, lpSourceRes->rt);
		wstring szFlags = InterpretFlags(flagRestrictionType, lpSourceRes->rt);
		SetStringW(1, szFlags.c_str());
	}

	return bRet;
} // CRestrictEditor::OnInitDialog

_Check_return_ LPSRestriction CRestrictEditor::GetSourceRes()
{
	if (m_lpRes && !m_bModified) return m_lpRes;
	return m_lpOutputRes;
} // CRestrictEditor::GetSourceRes

// Create our LPSRestriction from the dialog here
void CRestrictEditor::OnOK()
{
	CMyDialog::OnOK(); // don't need to call CEditor::OnOK
} // CRestrictEditor::OnOK

_Check_return_ LPSRestriction CRestrictEditor::DetachModifiedSRestriction()
{
	if (!m_bModified) return NULL;
	LPSRestriction m_lpRet = m_lpOutputRes;
	m_lpOutputRes = NULL;
	return m_lpRet;
} // CRestrictEditor::DetachModifiedSRestriction

// CEditor::HandleChange will return the control number of what changed
// so I can react to it if I need to
_Check_return_ ULONG CRestrictEditor::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	// If the restriction type changed
	if (0 == i)
	{
		if (!m_lpOutputRes) return i;
		HRESULT hRes = S_OK;
		ULONG ulOldResType = m_lpOutputRes->rt;
		ULONG ulNewResType = GetHexUseControl(i);

		if (ulOldResType == ulNewResType) return i;

		wstring szFlags = InterpretFlags(flagRestrictionType, ulNewResType);
		SetStringW(1, szFlags.c_str());

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
					(LPVOID*)&m_lpOutputRes));

				m_lpAllocParent = m_lpOutputRes;
			}
			else
			{
				// If the pointers are different, m_lpOutputRes was allocated with MAPIAllocateMore
				// Since m_lpOutputRes is owned by m_lpAllocParent, we don't free it directly
				EC_H(MAPIAllocateMore(
					sizeof(SRestriction),
					m_lpAllocParent,
					(LPVOID*)&m_lpOutputRes));
			}
			memset(m_lpOutputRes, 0, sizeof(SRestriction));
			m_lpOutputRes->rt = ulNewResType;
		}

		SetStringW(2, RestrictionToString(m_lpOutputRes, NULL).c_str());
	}
	return i;
} // CRestrictEditor::HandleChange

void CRestrictEditor::OnEditAction1()
{
	HRESULT hRes = S_OK;
	LPSRestriction lpSourceRes = GetSourceRes();
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
		SetStringW(2, RestrictionToString(m_lpOutputRes, NULL).c_str());
	}
}

HRESULT CRestrictEditor::EditCompare(LPSRestriction lpSourceRes)
{
	HRESULT hRes = S_OK;
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
	HRESULT hRes = S_OK;
	CResAndOrEditor MyResEditor(
		this,
		lpSourceRes,
		m_lpAllocParent); // pass source res into editor
	WC_H(MyResEditor.DisplayDialog());

	if (S_OK == hRes)
	{
		m_lpOutputRes->rt = lpSourceRes->rt;
		m_lpOutputRes->res.resAnd.cRes = MyResEditor.GetResCount();

		LPSRestriction lpNewResArray = MyResEditor.DetachModifiedSRestrictionArray();
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
	HRESULT hRes = S_OK;
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
	HRESULT hRes = S_OK;
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
				(LPVOID*)&m_lpOutputRes->res.resContent.lpProp));
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
	HRESULT hRes = S_OK;
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
	HRESULT hRes = S_OK;
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
	HRESULT hRes = S_OK;
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
	HRESULT hRes = S_OK;
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
	HRESULT hRes = S_OK;
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
	CEditor(pParentWnd, IDS_CRITERIAEDITOR, IDS_CRITERIAEDITORPROMPT, 6, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_CANCEL, IDS_ACTIONEDITRES, NULL, NULL)
{
	TRACE_CONSTRUCTOR(CRITERIACLASS);

	HRESULT hRes = S_OK;

	m_lpSourceRes = lpRes;
	m_lpNewRes = NULL;
	m_lpSourceEntryList = lpEntryList;

	EC_H(MAPIAllocateBuffer(
		sizeof(SBinaryArray),
		(LPVOID*)&m_lpNewEntryList));

	m_ulNewSearchFlags = NULL;

	SetPromptPostFix(AllFlagsToString(flagSearchFlag, true));
	InitPane(0, CreateSingleLinePane(IDS_SEARCHSTATE, NULL, true));
	SetHex(0, ulSearchState);
	wstring szFlags = InterpretFlags(flagSearchState, ulSearchState);
	InitPane(1, CreateSingleLinePaneW(IDS_SEARCHSTATE, szFlags.c_str(), true));
	InitPane(2, CreateSingleLinePane(IDS_SEARCHFLAGS, NULL, false));
	SetHex(2, 0);
	InitPane(3, CreateSingleLinePane(IDS_SEARCHFLAGS, NULL, true));
	InitPane(4, CreateListPane(IDS_EIDLIST, false, false, this));
	InitPane(5, CreateMultiLinePaneW(IDS_RESTRICTIONTEXT, RestrictionToString(m_lpSourceRes, NULL).c_str(), true));
} // CCriteriaEditor::CCriteriaEditor

CCriteriaEditor::~CCriteriaEditor()
{
	// If these structures weren't detached, we need to free them
	MAPIFreeBuffer(m_lpNewEntryList);
	MAPIFreeBuffer(m_lpNewRes);
} // CCriteriaEditor::~CCriteriaEditor

// Used to call functions which need to be called AFTER controls are created
BOOL CCriteriaEditor::OnInitDialog()
{
	BOOL bRet = CEditor::OnInitDialog();

	InitListFromEntryList(LISTNUM, m_lpSourceEntryList);

	UpdateListButtons();

	return bRet;
} // CCriteriaEditor::OnInitDialog

_Check_return_ LPSRestriction CCriteriaEditor::GetSourceRes()
{
	if (m_lpNewRes) return m_lpNewRes;
	return m_lpSourceRes;
} // CCriteriaEditor::GetSourceRes

_Check_return_ ULONG CCriteriaEditor::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	if (2 == i)
	{
		wstring szFlags = InterpretFlags(flagSearchFlag, GetHexUseControl(i));
		SetStringW(3, szFlags.c_str());
	}

	return i;
} // CCriteriaEditor::HandleChange

// Whoever gets this MUST MAPIFreeBuffer
_Check_return_ LPSRestriction CCriteriaEditor::DetachModifiedSRestriction()
{
	LPSRestriction m_lpRet = m_lpNewRes;
	m_lpNewRes = NULL;
	return m_lpRet;
} // CCriteriaEditor::DetachModifiedSRestriction

// Whoever gets this MUST MAPIFreeBuffer
_Check_return_ LPENTRYLIST CCriteriaEditor::DetachModifiedEntryList()
{
	LPENTRYLIST m_lpRet = m_lpNewEntryList;
	m_lpNewEntryList = NULL;
	return m_lpRet;
} // CCriteriaEditor::DetachModifiedEntryList

_Check_return_ ULONG CCriteriaEditor::GetSearchFlags()
{
	return m_ulNewSearchFlags;
} // CCriteriaEditor::GetSearchFlags

void CCriteriaEditor::InitListFromEntryList(ULONG ulListNum, _In_ LPENTRYLIST lpEntryList)
{
	if (!IsValidList(ulListNum)) return;

	ClearList(ulListNum);

	InsertColumn(ulListNum, 0, IDS_SHARP);
	CString szTmp;
	SortListData* lpData = NULL;
	ULONG i = 0;
	InsertColumn(ulListNum, 1, IDS_CB);
	InsertColumn(ulListNum, 2, IDS_BINARY);
	InsertColumn(ulListNum, 3, IDS_TEXTVIEW);

	if (lpEntryList)
	{
		for (i = 0; i < lpEntryList->cValues; i++)
		{
			szTmp.Format(_T("%u"), i); // STRING_OK
			lpData = InsertListRow(ulListNum, i, szTmp);
			if (lpData)
			{
				lpData->ulSortDataType = SORTLIST_BINARY;
				lpData->data.Binary.OldBin.cb = lpEntryList->lpbin[i].cb;
				lpData->data.Binary.OldBin.lpb = lpEntryList->lpbin[i].lpb;
				lpData->data.Binary.NewBin.cb = NULL;
				lpData->data.Binary.NewBin.lpb = NULL;
			}

			szTmp.Format(_T("%u"), lpEntryList->lpbin[i].cb); // STRING_OK
			SetListString(ulListNum, i, 1, szTmp);
			SetListStringW(ulListNum, i, 2, BinToHexString(&lpEntryList->lpbin[i], false).c_str());
			SetListStringW(ulListNum, i, 3, BinToTextString(&lpEntryList->lpbin[i], true).c_str());
			if (lpData) lpData->bItemFullyLoaded = true;
		}
	}
	ResizeList(ulListNum, false);
} // CCriteriaEditor::InitListFromEntryList

void CCriteriaEditor::OnEditAction1()
{
	HRESULT hRes = S_OK;

	LPSRestriction lpSourceRes = GetSourceRes();

	CRestrictEditor MyResEditor(
		this,
		NULL,
		lpSourceRes); // pass source res into editor
	WC_H(MyResEditor.DisplayDialog());

	if (S_OK == hRes)
	{
		LPSRestriction lpModRes = MyResEditor.DetachModifiedSRestriction();
		if (lpModRes)
		{
			// We didn't pass an alloc parent to CRestrictEditor, so we must free what came back
			MAPIFreeBuffer(m_lpNewRes);
			m_lpNewRes = lpModRes;
			SetStringW(5, RestrictionToString(m_lpNewRes, NULL).c_str());
		}
	}
} // CCriteriaEditor::OnEditAction1

_Check_return_ bool CCriteriaEditor::DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData)
{
	if (!lpData) return false;
	if (!IsValidList(ulListNum)) return false;
	HRESULT hRes = S_OK;
	CString szTmp;
	CString szAltTmp;

	CEditor BinEdit(
		this,
		IDS_EIDEDITOR,
		IDS_EIDEDITORPROMPT,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	LPSBinary lpSourcebin = NULL;
	if (lpData->data.Binary.OldBin.lpb)
	{
		lpSourcebin = &lpData->data.Binary.OldBin;
	}
	else
	{
		lpSourcebin = &lpData->data.Binary.NewBin;
	}

	BinEdit.InitPane(0, CreateSingleLinePaneW(IDS_EID, BinToHexString(lpSourcebin, false).c_str(), false));

	WC_H(BinEdit.DisplayDialog());
	if (S_OK == hRes)
	{
		vector<BYTE> bin = HexStringToBin(BinEdit.GetStringW(0));
		lpData->data.Binary.NewBin.lpb = ByteVectorToMAPI(bin, m_lpNewEntryList);
		if (lpData->data.Binary.NewBin.lpb)
		{
			lpData->data.Binary.NewBin.cb = bin.size();
			szTmp.Format(_T("%u"), lpData->data.Binary.NewBin.cb); // STRING_OK
			SetListString(ulListNum, iItem, 1, szTmp);
			SetListStringW(ulListNum, iItem, 2, BinToHexString(&lpData->data.Binary.NewBin, false).c_str());
			SetListStringW(ulListNum, iItem, 3, BinToTextString(&lpData->data.Binary.NewBin, true).c_str());
			return true;
		}
	}

	return false;
} // CCriteriaEditor::DoListEdit

void CCriteriaEditor::OnOK()
{
	CMyDialog::OnOK(); // don't need to call CEditor::OnOK
	if (!IsValidList(LISTNUM)) return;
	ULONG ulValues = GetListCount(LISTNUM);

	HRESULT hRes = S_OK;

	if (m_lpNewEntryList && ulValues < ULONG_MAX / sizeof(SBinary))
	{
		m_lpNewEntryList->cValues = ulValues;
		EC_H(MAPIAllocateMore(
			m_lpNewEntryList->cValues * sizeof(SBinary),
			m_lpNewEntryList,
			(LPVOID*)&m_lpNewEntryList->lpbin));

		ULONG i = 0;
		for (i = 0; i < m_lpNewEntryList->cValues; i++)
		{
			SortListData* lpData = GetListRowData(LISTNUM, i);
			if (lpData)
			{
				if (lpData->data.Binary.NewBin.lpb)
				{
					m_lpNewEntryList->lpbin[i].cb = lpData->data.Binary.NewBin.cb;
					m_lpNewEntryList->lpbin[i].lpb = lpData->data.Binary.NewBin.lpb;
					// clean out the source
					lpData->data.Binary.OldBin.lpb = NULL;
				}
				else
				{
					m_lpNewEntryList->lpbin[i].cb = lpData->data.Binary.OldBin.cb;
					EC_H(MAPIAllocateMore(
						m_lpNewEntryList->lpbin[i].cb,
						m_lpNewEntryList,
						(LPVOID*)&m_lpNewEntryList->lpbin[i].lpb));

					memcpy(m_lpNewEntryList->lpbin[i].lpb, lpData->data.Binary.OldBin.lpb, m_lpNewEntryList->lpbin[i].cb);
				}
			}
		}
	}
	if (!m_lpNewRes && m_lpSourceRes)
	{
		EC_H(HrCopyRestriction(m_lpSourceRes, NULL, &m_lpNewRes))
	}
	m_ulNewSearchFlags = GetHexUseControl(2);
} // CCriteriaEditor::OnOK