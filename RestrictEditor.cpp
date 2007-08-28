// RestrictEditor.cpp : implementation file
//

#include "stdafx.h"
#include "Error.h"

#include "RestrictEditor.h"
#include "PropertyEditor.h"

#include "InterpretProp.h"
#include "InterpretProp2.h"
#include "MAPIFunctions.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

static TCHAR* COMPCLASS = _T("CResCompareEditor");// STRING_OK
class CResCompareEditor : public CEditor
{
public:
	CResCompareEditor(
		CWnd* pParentWnd,
		ULONG ulRelop,
		ULONG ulPropTag1,
		ULONG ulPropTag2);
private:
	ULONG	HandleChange(UINT nID);
};

CResCompareEditor::CResCompareEditor(
								 CWnd* pParentWnd,
								 ULONG ulRelop,
								 ULONG ulPropTag1,
								 ULONG ulPropTag2):
CEditor(pParentWnd,IDS_RESED,IDS_RESEDCOMPPROMPT,6,CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(COMPCLASS);

	SetPromptPostFix(AllFlagsToString(flagRelop,false));
	InitSingleLine(0,IDS_RELOP,NULL,false);
	SetHex(0,ulRelop);
	HRESULT hRes = S_OK;
	LPTSTR szFlags = NULL;
	EC_H(InterpretFlags(flagRelop, ulRelop, &szFlags));
	InitSingleLineSz(1,IDS_RELOP,szFlags,true);
	MAPIFreeBuffer(szFlags);
	szFlags = NULL;

	InitSingleLine(2,IDS_ULPROPTAG1,NULL,false);
	SetHex(2,ulPropTag1);
	InitSingleLineSz(3,IDS_ULPROPTAG1,TagToString(ulPropTag1,NULL,false,true),true);

	InitSingleLine(4,IDS_ULPROPTAG2,NULL,false);
	SetHex(4,ulPropTag2);
	InitSingleLineSz(5,IDS_ULPROPTAG1,TagToString(ulPropTag2,NULL,false,true),true);
}

ULONG CResCompareEditor::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	if (0 == i)
	{
		HRESULT hRes = S_OK;
		LPTSTR szFlags = NULL;
		EC_H(InterpretFlags(flagRelop, GetHexUseControl(0), &szFlags));
		SetString(1,szFlags);
		MAPIFreeBuffer(szFlags);
		szFlags = NULL;
	}
	else if (2 == i)
	{
		SetString(3,TagToString(GetHexUseControl(2),NULL,false,true));
	}
	else if (4 == i)
	{
		SetString(5,TagToString(GetHexUseControl(4),NULL,false,true));
	}
	return i;
}

static TCHAR* CONTENTCLASS = _T("CResCombinedEditor");// STRING_OK
class CResCombinedEditor : public CEditor
{
public:
	CResCombinedEditor(
		CWnd* pParentWnd,
		ULONG ulResType,
		ULONG ulCompare,
		ULONG ulPropTag,
		LPSPropValue lpProp,
		LPVOID lpAllocParent);
	~CResCombinedEditor();
	virtual void OnEditAction1();
	LPSPropValue DetachModifiedSPropValue();

private:
	ULONG	HandleChange(UINT nID);
	ULONG			m_ulResType;
	LPVOID			m_lpAllocParent;
	LPSPropValue	m_lpOldProp;
	LPSPropValue	m_lpNewProp;
};

CResCombinedEditor::CResCombinedEditor(
									 CWnd* pParentWnd,
									 ULONG ulResType,
									 ULONG ulCompare,
									 ULONG ulPropTag,
									 LPSPropValue lpProp,
									 LPVOID lpAllocParent):
CEditor(pParentWnd,IDS_RESED,NULL,8,CEDITOR_BUTTON_OK|CEDITOR_BUTTON_ACTION1|CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(CONTENTCLASS);

	m_uidActionButtonText1 = IDS_ACTIONEDITPROP;
	m_ulResType = ulResType;
	m_lpOldProp = lpProp;
	m_lpNewProp = NULL;
	m_lpAllocParent = lpAllocParent;

	HRESULT hRes = S_OK;
	LPTSTR szFlags = NULL;

	if (RES_CONTENT == m_ulResType)
	{
		m_uidPrompt = IDS_RESEDCONTPROMPT;
		SetPromptPostFix(AllFlagsToString(flagFuzzyLevel,true));

		InitSingleLine(0,IDS_ULFUZZYLEVEL,NULL,false);
		SetHex(0,ulCompare);
		EC_H(InterpretFlags(flagFuzzyLevel, ulCompare, &szFlags));
		InitSingleLineSz(1,IDS_ULFUZZYLEVEL,szFlags,true);
	}
	else if (RES_PROPERTY == m_ulResType)
	{
		m_uidPrompt = IDS_RESEDPROPPROMPT;
		SetPromptPostFix(AllFlagsToString(flagRelop,false));
		InitSingleLine(0,IDS_RELOP,NULL,false);
		SetHex(0,ulCompare);
		EC_H(InterpretFlags(flagRelop, ulCompare, &szFlags));
		InitSingleLineSz(1,IDS_RELOP,szFlags,true);
	}
	MAPIFreeBuffer(szFlags);
	szFlags = NULL;

	InitSingleLine(2,IDS_ULPROPTAG,NULL,false);
	SetHex(2,ulPropTag);
	InitSingleLineSz(3,IDS_ULPROPTAG,TagToString(ulPropTag,NULL,false,true),true);

	InitSingleLine(4,IDS_LPPROPULPROPTAG,NULL,false);
	if (lpProp) SetHex(4,lpProp->ulPropTag);
	InitSingleLineSz(5,IDS_LPPROPULPROPTAG,lpProp?(LPCTSTR)TagToString(lpProp->ulPropTag,NULL,false,true):NULL,true);

	CString szProp;
	CString szAltProp;
	InterpretProp(lpProp,&szProp,&szAltProp);
	InitMultiLine(6,IDS_LPPROP,szProp,true);
	InitMultiLine(7,IDS_LPPROPALTVIEW,szAltProp,true);
}

CResCombinedEditor::~CResCombinedEditor()
{
	MAPIFreeBuffer(m_lpNewProp);
}

ULONG CResCombinedEditor::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	HRESULT hRes = S_OK;
	LPTSTR szFlags = NULL;
	if (0 == i)
	{
		if (RES_CONTENT == m_ulResType)
		{
			EC_H(InterpretFlags(flagFuzzyLevel, GetHexUseControl(0), &szFlags));
			SetString(1,szFlags);
		}
		else if (RES_PROPERTY == m_ulResType)
		{
			EC_H(InterpretFlags(flagRelop, GetHexUseControl(0), &szFlags));
			SetString(1,szFlags);
		}
		MAPIFreeBuffer(szFlags);
		szFlags = NULL;
	}
	else if (2 == i)
	{
		SetString(3,TagToString(GetHexUseControl(2),NULL,false,true));
	}
	else if (4 == i)
	{
		ULONG ulNewPropTag = GetHexUseControl(4);
		SetString(5,TagToString(ulNewPropTag,NULL,false,true));
		MAPIFreeBuffer(m_lpOldProp);
		MAPIFreeBuffer(m_lpNewProp);
		m_lpOldProp = NULL;
		m_lpNewProp = NULL;
		SetString(6,NULL);
		SetString(7,NULL);
	}
	return i;
}

LPSPropValue CResCombinedEditor::DetachModifiedSPropValue()
{
	LPSPropValue m_lpRet = m_lpNewProp;
	m_lpNewProp = NULL;
	return m_lpRet;
}

void CResCombinedEditor::OnEditAction1()
{
	HRESULT hRes = S_OK;
	CPropertyEditor PropEdit(
		this,
		IDS_PROPEDITOR,
		IDS_PROPEDITORPROMPT,
		false,
		m_lpAllocParent);


	LPSPropValue lpEditProp = m_lpOldProp;
	if (m_lpNewProp) lpEditProp = m_lpNewProp;

	if (lpEditProp)
	{
		PropEdit.InitPropValue(lpEditProp);
	}
	else
	{
		PropEdit.InitPropValue(NULL,GetHexUseControl(4));
	}

	WC_H(PropEdit.DisplayDialog());

	if (S_OK == hRes)
	{
		MAPIFreeBuffer(m_lpNewProp);
		m_lpNewProp = PropEdit.DetachModifiedSPropValue();
		CString szProp;
		CString szAltProp;
		if (m_lpNewProp)
		{
			InterpretProp(m_lpNewProp,&szProp,&szAltProp);
			SetString(6,szProp);
			SetString(7,szAltProp);
		}
	}
}

static TCHAR* BITMASKCLASS = _T("CResBitmaskEditor");// STRING_OK
class CResBitmaskEditor : public CEditor
{
public:
	CResBitmaskEditor(
		CWnd* pParentWnd,
		ULONG relBMR,
		ULONG ulPropTag,
		ULONG ulMask);
private:
	ULONG	HandleChange(UINT nID);
};

CResBitmaskEditor::CResBitmaskEditor(
									 CWnd* pParentWnd,
									 ULONG relBMR,
									 ULONG ulPropTag,
									 ULONG ulMask):
CEditor(pParentWnd,IDS_RESED,IDS_RESEDBITPROMPT,5,CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(BITMASKCLASS);

	SetPromptPostFix(AllFlagsToString(flagBitmask,false));
	InitSingleLine(0,IDS_RELBMR,NULL,false);
	SetHex(0,relBMR);
	HRESULT hRes = S_OK;
	LPTSTR szFlags = NULL;
	EC_H(InterpretFlags(flagBitmask, relBMR, &szFlags));
	InitSingleLineSz(1,IDS_RELBMR,szFlags,true);
	MAPIFreeBuffer(szFlags);
	szFlags = NULL;

	InitSingleLine(2,IDS_ULPROPTAG,NULL,false);
	SetHex(2,ulPropTag);
	InitSingleLineSz(3,IDS_ULPROPTAG,TagToString(ulPropTag,NULL,false,true),true);

	InitSingleLine(4,IDS_MASK,NULL,false);
	SetHex(4,ulMask);
}

ULONG CResBitmaskEditor::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	if (0 == i)
	{
		HRESULT hRes = S_OK;
		LPTSTR szFlags = NULL;
		EC_H(InterpretFlags(flagBitmask, GetHexUseControl(0), &szFlags));
		SetString(1,szFlags);
		MAPIFreeBuffer(szFlags);
		szFlags = NULL;
	}
	else if (2 == i)
	{
		SetString(3,TagToString(GetHexUseControl(2),NULL,false,true));
	}
	return i;
}

static TCHAR* SIZECLASS = _T("CResSizeEditor");// STRING_OK
class CResSizeEditor : public CEditor
{
public:
	CResSizeEditor(
		CWnd* pParentWnd,
		ULONG relop,
		ULONG ulPropTag,
		ULONG cb);
private:
	ULONG	HandleChange(UINT nID);
};

CResSizeEditor::CResSizeEditor(
									 CWnd* pParentWnd,
									 ULONG relop,
									 ULONG ulPropTag,
									 ULONG cb):
CEditor(pParentWnd,IDS_RESED,IDS_RESEDSIZEPROMPT,5,CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(SIZECLASS);

	SetPromptPostFix(AllFlagsToString(flagRelop,false));
	InitSingleLine(0,IDS_RELOP,NULL,false);
	SetHex(0,relop);
	HRESULT hRes = S_OK;
	LPTSTR szFlags = NULL;
	EC_H(InterpretFlags(flagRelop, relop, &szFlags));
	InitSingleLineSz(1,IDS_RELOP,szFlags,true);
	MAPIFreeBuffer(szFlags);
	szFlags = NULL;

	InitSingleLine(2,IDS_ULPROPTAG,NULL,false);
	SetHex(2,ulPropTag);
	InitSingleLineSz(3,IDS_ULPROPTAG,TagToString(ulPropTag,NULL,false,true),true);

	InitSingleLine(4,IDS_CB,NULL,false);
	SetHex(4,cb);
}

ULONG CResSizeEditor::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	if (0 == i)
	{
		HRESULT hRes = S_OK;
		LPTSTR szFlags = NULL;
		EC_H(InterpretFlags(flagRelop, GetHexUseControl(0), &szFlags));
		SetString(1,szFlags);
		MAPIFreeBuffer(szFlags);
		szFlags = NULL;
	}
	else if (2 == i)
	{
		SetString(3,TagToString(GetHexUseControl(2),NULL,false,true));
	}
	return i;
}

static TCHAR* EXISTCLASS = _T("CResExistEditor");// STRING_OK
class CResExistEditor : public CEditor
{
public:
	CResExistEditor(
		CWnd* pParentWnd,
		ULONG ulPropTag);
private:
	ULONG	HandleChange(UINT nID);
};

CResExistEditor::CResExistEditor(
									 CWnd* pParentWnd,
									 ULONG ulPropTag):
CEditor(pParentWnd,IDS_RESED,IDS_RESEDEXISTPROMPT,2,CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(EXISTCLASS);

	InitSingleLine(0,IDS_ULPROPTAG,NULL,false);
	SetHex(0,ulPropTag);
	InitSingleLineSz(1,IDS_ULPROPTAG,TagToString(ulPropTag,NULL,false,true),true);
}

ULONG CResExistEditor::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	if (0 == i)
	{
		SetString(1,TagToString(GetHexUseControl(0),NULL,false,true));
	}
	return i;
}

static TCHAR* SUBRESCLASS = _T("CResSubResEditor");// STRING_OK
class CResSubResEditor : public CEditor
{
public:
	CResSubResEditor(
		CWnd* pParentWnd,
		ULONG ulSubObject,
		LPSRestriction lpRes,
		LPVOID lpAllocParent);
	virtual void OnEditAction1();
	LPSRestriction DetachModifiedSRestriction();
private:
	ULONG	HandleChange(UINT nID);
	LPVOID			m_lpAllocParent;
	LPSRestriction	m_lpOldRes;
	LPSRestriction	m_lpNewRes;
};

CResSubResEditor::CResSubResEditor(
									 CWnd* pParentWnd,
									 ULONG ulSubObject,
									 LPSRestriction lpRes,
									 LPVOID lpAllocParent):
CEditor(pParentWnd,IDS_SUBRESED,IDS_RESEDSUBPROMPT,3,CEDITOR_BUTTON_OK|CEDITOR_BUTTON_ACTION1|CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(SUBRESCLASS);

	m_uidActionButtonText1 = IDS_ACTIONEDITRES;
	m_lpOldRes = lpRes;
	m_lpNewRes = NULL;
	m_lpAllocParent = lpAllocParent;

	InitSingleLine(0,IDS_ULSUBOBJECT,NULL,false);
	SetHex(0,ulSubObject);
	InitSingleLineSz(1,IDS_ULSUBOBJECT,TagToString(ulSubObject,NULL,false,true),true);

	InitMultiLine(2,IDS_LPRES,RestrictionToString(lpRes,NULL),true);
}

ULONG CResSubResEditor::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	if (0 == i)
	{
		SetString(1,TagToString(GetHexUseControl(0),NULL,false,true));
	}
	return i;
}

LPSRestriction CResSubResEditor::DetachModifiedSRestriction()
{
	LPSRestriction m_lpRet = m_lpNewRes;
	m_lpNewRes = NULL;
	return m_lpRet;
}

void CResSubResEditor::OnEditAction1()
{
	HRESULT hRes = S_OK;
	CRestrictEditor ResEdit(
		this,
		m_lpAllocParent,
		m_lpNewRes?m_lpNewRes:m_lpOldRes);

	WC_H(ResEdit.DisplayDialog());

	if (S_OK == hRes)
	{
		MAPIFreeBuffer(m_lpNewRes);
		m_lpNewRes = ResEdit.DetachModifiedSRestriction();
		InitMultiLine(2,IDS_LPRES,RestrictionToString(m_lpNewRes,NULL),true);
	}
}

static TCHAR* ANDORCLASS = _T("CResAndOrEditor");// STRING_OK
class CResAndOrEditor : public CEditor
{
public:
	CResAndOrEditor(
		CWnd* pParentWnd,
		LPSRestriction lpRes,
		LPVOID lpAllocParent);
	LPSRestriction  DetachModifiedSRestrictionArray();
	ULONG			GetResCount();
protected:
	//Use this function to implement list editing
	virtual BOOL	DoListEdit(ULONG ulListNum, int iItem, SortListData* lpData);

private:
	BOOL	OnInitDialog();
	void	InitListFromRestriction(ULONG ulListNum,LPSRestriction lpRes);
	void	OnOK();

	LPVOID			m_lpAllocParent;
	LPSRestriction	m_lpRes;
	LPSRestriction	m_lpNewResArray;
	ULONG			m_ulNewResCount;
};

CResAndOrEditor::CResAndOrEditor(
									 CWnd* pParentWnd,
									 LPSRestriction lpRes,
									 LPVOID lpAllocParent):
CEditor(pParentWnd,IDS_RESED,IDS_RESEDANDORPROMPT,1,CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(ANDORCLASS);
	m_lpRes = lpRes;
	m_lpNewResArray = NULL;
	m_ulNewResCount = NULL;
	m_lpAllocParent = lpAllocParent;

	InitList(0,IDS_SUBRESTRICTIONS,false,false);
}

//Used to call functions which need to be called AFTER controls are created
BOOL CResAndOrEditor::OnInitDialog()
{
	HRESULT hRes = S_OK;
	EC_B(CEditor::OnInitDialog());

	InitListFromRestriction(0,m_lpRes);

	UpdateListButtons();

	return HRES_TO_BOOL(hRes);
}

LPSRestriction CResAndOrEditor::DetachModifiedSRestrictionArray()
{
	LPSRestriction m_lpRet = m_lpNewResArray;
	m_lpNewResArray = NULL;
	return m_lpRet;
}

ULONG CResAndOrEditor::GetResCount()
{
	return m_ulNewResCount;
}

void CResAndOrEditor::InitListFromRestriction(ULONG ulListNum,LPSRestriction lpRes)
{
	if (!m_lpControls) return;
	if (!IsValidList(ulListNum)) return;
	if (!lpRes) return;

	HRESULT hRes = S_OK;
	m_lpControls[ulListNum].UI.lpList->List.DeleteAllColumns();
	EC_B(m_lpControls[ulListNum].UI.lpList->List.DeleteAllItems());

	InsertColumn(ulListNum,0,IDS_SHARP);
	CString szTmp;
	CString szAltTmp;
	SortListData* lpData = NULL;
	ULONG i = 0;
	switch(lpRes->rt)
	{
	case RES_AND:
		InsertColumn(ulListNum,1,IDS_SUBRESTRICTION);

		for (i = 0;i< lpRes->res.resAnd.cRes;i++)
		{
			szTmp.Format(_T("%d"),i);// STRING_OK
			lpData = m_lpControls[ulListNum].UI.lpList->List.InsertRow(i,(LPTSTR)(LPCTSTR)szTmp);
			lpData->ulSortDataType = SORTLIST_RES;
			lpData->data.Res.lpOldRes = &lpRes->res.resAnd.lpRes[i];//save off address of source restriction
			WC_B(m_lpControls[ulListNum].UI.lpList->List.SetItemText(i,1,(LPCTSTR)RestrictionToString(lpData->data.Res.lpOldRes,NULL)));
			lpData->bItemFullyLoaded = true;
		}
		break;
	case RES_OR:
		InsertColumn(ulListNum,1,IDS_SUBRESTRICTION);

		for (i = 0;i< lpRes->res.resOr.cRes;i++)
		{
			szTmp.Format(_T("%d"),i);// STRING_OK
			lpData = m_lpControls[ulListNum].UI.lpList->List.InsertRow(i,(LPTSTR)(LPCTSTR)szTmp);
			lpData->ulSortDataType = SORTLIST_RES;
			lpData->data.Res.lpOldRes = &lpRes->res.resOr.lpRes[i];//save off address of source restriction
			WC_B(m_lpControls[ulListNum].UI.lpList->List.SetItemText(i,1,(LPCTSTR)RestrictionToString(lpData->data.Res.lpOldRes,NULL)));
			lpData->bItemFullyLoaded = true;
		}
		break;
	}
	m_lpControls[ulListNum].UI.lpList->List.AutoSizeColumns();
}//CResAndOrEditor::InitListFromRestriction

BOOL CResAndOrEditor::DoListEdit(ULONG ulListNum, int iItem, SortListData* lpData)
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
		lpSourceRes);//pass source res into editor
	WC_H(MyResEditor.DisplayDialog());

	if (S_OK == hRes)
	{
		MAPIFreeBuffer(lpData->data.Res.lpNewRes);
		lpData->data.Res.lpNewRes = MyResEditor.DetachModifiedSRestriction();
		WC_B(m_lpControls[ulListNum].UI.lpList->List.SetItemText(iItem,1,(LPCTSTR)RestrictionToString(lpData->data.Res.lpNewRes,NULL)));
		return true;
	}
	return false;
}

//Create our LPSRestriction array from the dialog here
void CResAndOrEditor::OnOK()
{
	CDialog::OnOK();//don't need to call CEditor::OnOK

	if (!IsValidList(0)) return;
	LPSRestriction  lpNewResArray = NULL;
	ULONG ulNewResCount = m_lpControls[0].UI.lpList->List.GetItemCount();

	if (ulNewResCount > ULONG_MAX/sizeof(SRestriction)) return;
	HRESULT hRes = S_OK;
	EC_H(MAPIAllocateMore(
		sizeof(SRestriction) * ulNewResCount,
		m_lpAllocParent,
		(LPVOID*)&lpNewResArray));

	if (lpNewResArray)
	{
		ULONG i = 0;
		for (i=0;i<ulNewResCount;i++)
		{
			SortListData* lpData = (SortListData*) m_lpControls[0].UI.lpList->List.GetItemData(i);
			if (lpData->data.Res.lpNewRes)
			{
				memcpy(&lpNewResArray[i],lpData->data.Res.lpNewRes,sizeof(SRestriction));
				memset(lpData->data.Res.lpNewRes,0,sizeof(SRestriction));
			}
			else
			{
				EC_H(CopyRestriction(&lpNewResArray[i],lpData->data.Res.lpOldRes,m_lpAllocParent));
			}
		}
		m_ulNewResCount = ulNewResCount;
		m_lpNewResArray = lpNewResArray;
	}
}

static TCHAR* COMMENTCLASS = _T("CResCommentEditor");// STRING_OK
class CResCommentEditor : public CEditor
{
public:
	CResCommentEditor(
		CWnd* pParentWnd,
		LPSRestriction lpRes,
		LPVOID lpAllocParent);
	virtual void OnEditAction1();
	LPSRestriction DetachModifiedSRestriction();
	LPSPropValue DetachModifiedSPropValue();
	ULONG GetSPropValueCount();
protected:
	//Use this function to implement list editing
	virtual BOOL	DoListEdit(ULONG ulListNum, int iItem, SortListData* lpData);

private:
	BOOL	OnInitDialog();
	void	InitListFromPropArray(ULONG ulListNum, ULONG cProps, LPSPropValue lpProps);
	LPSRestriction GetSourceRes();
	void	OnOK();

	LPVOID			m_lpAllocParent;
	LPSRestriction	m_lpSourceRes;

	LPSRestriction	m_lpNewCommentRes;
	LPSPropValue	m_lpNewCommentProp;
	ULONG			m_ulNewCommentProp;
};

CResCommentEditor::CResCommentEditor(
									 CWnd* pParentWnd,
									 LPSRestriction lpRes,
									 LPVOID lpAllocParent):
CEditor(pParentWnd,IDS_COMMENTRESED,IDS_RESEDCOMMENTPROMPT,2,CEDITOR_BUTTON_OK|CEDITOR_BUTTON_ACTION1|CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(COMMENTCLASS);

	m_uidActionButtonText1 = IDS_ACTIONEDITRES;

	m_lpSourceRes = lpRes;
	m_lpNewCommentRes = NULL;
	m_lpNewCommentProp = NULL;
	m_lpAllocParent = lpAllocParent;

	InitList(0,IDS_SUBRESTRICTION,false,false);
	InitMultiLine(1,IDS_RESTRICTIONTEXT,RestrictionToString(m_lpSourceRes->res.resComment.lpRes,NULL),true);
}

//Used to call functions which need to be called AFTER controls are created
BOOL CResCommentEditor::OnInitDialog()
{
	HRESULT hRes = S_OK;
	EC_B(CEditor::OnInitDialog());

	InitListFromPropArray(0,m_lpSourceRes->res.resComment.cValues,m_lpSourceRes->res.resComment.lpProp);

	UpdateListButtons();

	return HRES_TO_BOOL(hRes);
}

LPSRestriction CResCommentEditor::GetSourceRes()
{
	if (m_lpNewCommentRes) return m_lpNewCommentRes;
	if (m_lpSourceRes && m_lpSourceRes->res.resComment.lpRes) return m_lpSourceRes->res.resComment.lpRes;
	return NULL;
}

LPSRestriction CResCommentEditor::DetachModifiedSRestriction()
{
	LPSRestriction m_lpRet = m_lpNewCommentRes;
	m_lpNewCommentRes = NULL;
	return m_lpRet;
}

LPSPropValue CResCommentEditor::DetachModifiedSPropValue()
{
	LPSPropValue m_lpRet = m_lpNewCommentProp;
	m_lpNewCommentProp = NULL;
	return m_lpRet;
}

ULONG CResCommentEditor::GetSPropValueCount()
{
	return m_ulNewCommentProp;
}


void CResCommentEditor::InitListFromPropArray(ULONG ulListNum, ULONG cProps, LPSPropValue lpProps)
{
	if (!m_lpControls) return;
	if (!IsValidList(ulListNum)) return;

	HRESULT hRes = S_OK;
	m_lpControls[ulListNum].UI.lpList->List.DeleteAllColumns();
	EC_B(m_lpControls[ulListNum].UI.lpList->List.DeleteAllItems());

	InsertColumn(ulListNum,0,IDS_SHARP);
	CString szTmp;
	CString szAltTmp;
	SortListData* lpData = NULL;
	ULONG i = 0;
	InsertColumn(ulListNum,1,IDS_PROPERTY);
	InsertColumn(ulListNum,2,IDS_VALUE);
	InsertColumn(ulListNum,3,IDS_ALTERNATEVIEW);

	for (i = 0;i< cProps;i++)
	{
		szTmp.Format(_T("%d"),i);// STRING_OK
		lpData = m_lpControls[ulListNum].UI.lpList->List.InsertRow(i,(LPTSTR)(LPCTSTR)szTmp);
		lpData->ulSortDataType = SORTLIST_COMMENT;
		lpData->data.Comment.lpOldProp = &lpProps[i];
		WC_B(m_lpControls[ulListNum].UI.lpList->List.SetItemText(i,1,
			(LPCTSTR)TagToString(lpProps[i].ulPropTag,NULL,false,true)));
		InterpretProp(&lpProps[i],&szTmp,&szAltTmp);
		WC_B(m_lpControls[ulListNum].UI.lpList->List.SetItemText(i,2,(LPCTSTR)szTmp));
		WC_B(m_lpControls[ulListNum].UI.lpList->List.SetItemText(i,3,(LPCTSTR)szAltTmp));
		lpData->bItemFullyLoaded = true;
	}
	m_lpControls[ulListNum].UI.lpList->List.AutoSizeColumns();
}//CResCommentEditor::InitListFromPropArray

void CResCommentEditor::OnEditAction1()
{
	HRESULT hRes = S_OK;

	LPSRestriction lpSourceRes = GetSourceRes();

	CRestrictEditor MyResEditor(
		this,
		m_lpAllocParent,
		lpSourceRes);//pass source res into editor
	WC_H(MyResEditor.DisplayDialog());

	if (S_OK == hRes)
	{
		MAPIFreeBuffer(m_lpNewCommentRes);
		m_lpNewCommentRes = MyResEditor.DetachModifiedSRestriction();
		SetString(1,RestrictionToString(m_lpNewCommentRes,NULL));
	}
}

BOOL CResCommentEditor::DoListEdit(ULONG ulListNum, int iItem, SortListData* lpData)
{
	if (!lpData) return false;
	if (!IsValidList(ulListNum)) return false;
	HRESULT hRes = S_OK;

	CPropertyEditor PropEdit(
		this,
		IDS_PROPEDITOR,
		IDS_PROPEDITORPROMPT,
		false,
		m_lpAllocParent);

	LPSPropValue lpSourceProp = lpData->data.Comment.lpNewProp;
	if (!lpSourceProp) lpSourceProp = lpData->data.Comment.lpOldProp;

	SPropValue sProp = {0};

	if (!lpSourceProp)
	{
		CEditor MyTag(
			this,
			IDS_TAG,
			IDS_TAGPROMPT,
			1,
			true);

		MyTag.InitSingleLine(0,IDS_TAG,NULL,false);

		WC_H(MyTag.DisplayDialog());
		if (S_OK != hRes) return false;
		sProp.ulPropTag = MyTag.GetHex(0);
		lpSourceProp = &sProp;
	}

	PropEdit.InitPropValue(lpSourceProp);

	WC_H(PropEdit.DisplayDialog());
	if (S_OK == hRes)
	{
		MAPIFreeBuffer(lpData->data.Comment.lpNewProp);
		lpData->data.Comment.lpNewProp = PropEdit.DetachModifiedSPropValue();

		if (lpData->data.Comment.lpNewProp)
		{
			CString szTmp;
			CString szAltTmp;
			WC_B(m_lpControls[ulListNum].UI.lpList->List.SetItemText(iItem,1,
				(LPCTSTR)TagToString(lpData->data.Comment.lpNewProp->ulPropTag,NULL,false,true)));
			InterpretProp(lpData->data.Comment.lpNewProp,&szTmp,&szAltTmp);
			WC_B(m_lpControls[ulListNum].UI.lpList->List.SetItemText(iItem,2,(LPCTSTR)szTmp));
			WC_B(m_lpControls[ulListNum].UI.lpList->List.SetItemText(iItem,3,(LPCTSTR)szAltTmp));
			return true;
		}
	}

	return false;
}

void CResCommentEditor::OnOK()
{
	CDialog::OnOK();//don't need to call CEditor::OnOK
	if (!IsValidList(0)) return;

	LPSPropValue lpNewCommentProp = NULL;
	ULONG ulNewCommentProp = m_lpControls[0].UI.lpList->List.GetItemCount();

	HRESULT hRes = S_OK;
	if (ulNewCommentProp && ulNewCommentProp < ULONG_MAX/sizeof(SPropValue))
	{
		EC_H(MAPIAllocateMore(
			sizeof(SPropValue) * ulNewCommentProp,
			m_lpAllocParent,
			(LPVOID*)&lpNewCommentProp));

		if (lpNewCommentProp)
		{
			ULONG i = 0;
			for (i=0;i<ulNewCommentProp;i++)
			{
				SortListData* lpData = (SortListData*) m_lpControls[0].UI.lpList->List.GetItemData(i);
				if (lpData->data.Comment.lpNewProp)
				{
					EC_H(PropCopyMore(
						&lpNewCommentProp[i],
						lpData->data.Comment.lpNewProp,
						MAPIAllocateMore,
						m_lpAllocParent));
				}
				else
				{
					EC_H(PropCopyMore(
						&lpNewCommentProp[i],
						lpData->data.Comment.lpOldProp,
						MAPIAllocateMore,
						m_lpAllocParent));
				}
			}
			m_ulNewCommentProp = ulNewCommentProp;
			m_lpNewCommentProp = lpNewCommentProp;
		}
	}
	if (!m_lpNewCommentRes && m_lpSourceRes && m_lpSourceRes->res.resComment.lpRes)
	{
		EC_H(CopyRestriction(&m_lpNewCommentRes,m_lpSourceRes->res.resComment.lpRes,m_lpAllocParent));
	}
}

static TCHAR* CLASS = _T("CRestrictEditor");// STRING_OK

//Create an editor for a restriction
//Takes LPSRestriction lpRestriction as input
CRestrictEditor::CRestrictEditor(
								 CWnd* pParentWnd,
								 LPVOID lpAllocParent,
								 LPSRestriction lpRes):
CEditor(pParentWnd,IDS_RESED,IDS_RESEDPROMPT,3,CEDITOR_BUTTON_OK|CEDITOR_BUTTON_ACTION1|CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(CLASS);
	HRESULT hRes = S_OK;

	m_uidActionButtonText1 = IDS_ACTIONEDITRES;

	//Not copying the source restriction, but since we're modal it won't matter
	m_lpRes = lpRes;
	m_lpOutputRes = NULL;
	m_bModified = false;
	if (!lpRes) m_bModified = true;//if we don't have a source, make sure we never look at it

	m_lpAllocParent = lpAllocParent;

	//Set up our output restriction and determine m_lpAllocParent- this will be the basis for all other allocations
	//Allocate base memory:
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
		memset(m_lpOutputRes,0,sizeof(SRestriction));
		if (m_lpRes) m_lpOutputRes->rt = m_lpRes->rt;
	}

	SetPromptPostFix(AllFlagsToString(flagRestrictionType,true));
	InitSingleLine(0,IDS_RESTRICTIONTYPE,NULL,false);//type as a number
	InitSingleLine(1,IDS_RESTRICTIONTYPE,NULL,true);//type as a string (flagRestrictionType)
	InitMultiLine(2,IDS_RESTRICTIONTEXT,RestrictionToString(GetSourceRes(),NULL),true);
}

CRestrictEditor::~CRestrictEditor()
{
	TRACE_DESTRUCTOR(CLASS);
	MAPIFreeBuffer(m_lpOutputRes);
}

BEGIN_MESSAGE_MAP(CRestrictEditor, CEditor)
//{{AFX_MSG_MAP(CRestrictEditor)
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

//Used to call functions which need to be called AFTER controls are created
BOOL CRestrictEditor::OnInitDialog()
{
	HRESULT hRes = S_OK;
	EC_B(CEditor::OnInitDialog());

	LPSRestriction lpSourceRes = GetSourceRes();

	if (lpSourceRes)
	{
		SetHex(0,lpSourceRes->rt);
		LPTSTR szFlags = NULL;
		EC_H(InterpretFlags(flagRestrictionType, lpSourceRes->rt, &szFlags));
		SetString(1,szFlags);
		MAPIFreeBuffer(szFlags);
		szFlags = NULL;
	}
	return HRES_TO_BOOL(hRes);
}

LPSRestriction CRestrictEditor::GetSourceRes()
{
	if (m_lpRes && !m_bModified) return m_lpRes;
	return m_lpOutputRes;
}

//Create our LPSRestriction from the dialog here
void CRestrictEditor::OnOK()
{
	CDialog::OnOK();//don't need to call CEditor::OnOK
}

LPSRestriction CRestrictEditor::DetachModifiedSRestriction()
{
	if (!m_bModified) return NULL;
	LPSRestriction m_lpRet = m_lpOutputRes;
	m_lpOutputRes = NULL;
	return m_lpRet;
}

//CEditor::HandleChange will return the control number of what changed
//so I can react to it if I need to
ULONG CRestrictEditor::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	if (0 == i)
	{
		if (!m_lpOutputRes) return i;
		HRESULT hRes = S_OK;
		ULONG ulOldResType = m_lpOutputRes->rt;
		ULONG ulNewResType = GetHexUseControl(i);

		if (ulOldResType == ulNewResType) return i;

		LPTSTR szFlags = NULL;
		EC_H(InterpretFlags(flagRestrictionType, ulNewResType, &szFlags));
		SetString(1,szFlags);
		MAPIFreeBuffer(szFlags);
		szFlags = NULL;

		m_bModified = true;
		if ((ulOldResType == RES_AND || ulOldResType == RES_OR) &&
			(ulNewResType == RES_AND || ulNewResType == RES_OR))
		{
			m_lpOutputRes->rt = ulNewResType;//just need to adjust our type and refresh
		}
		else
		{
			//Need to wipe out our current output restriction and rebuild it
			if (m_lpAllocParent == m_lpOutputRes)
			{
				MAPIFreeBuffer(m_lpOutputRes);
				EC_H(MAPIAllocateBuffer(
					sizeof(SRestriction),
					(LPVOID*)&m_lpOutputRes));

				m_lpAllocParent = m_lpOutputRes;
			}
			else
			{
				MAPIFreeBuffer(m_lpOutputRes);
				EC_H(MAPIAllocateMore(
					sizeof(SRestriction),
					m_lpAllocParent,
					(LPVOID*)&m_lpOutputRes));
			}
			memset(m_lpOutputRes,0,sizeof(SRestriction));
			m_lpOutputRes->rt = ulNewResType;
		}

		SetString(2,RestrictionToString(m_lpOutputRes,NULL));
	}
	return i;
}

void CRestrictEditor::OnEditAction1()
{
	HRESULT hRes = S_OK;

	LPSRestriction lpSourceRes = GetSourceRes();

	if (!lpSourceRes || !m_lpOutputRes) return;

	//Do we need to set this again?
	switch(lpSourceRes->rt)
	{
	case RES_COMPAREPROPS:
		{
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
				m_lpOutputRes->res.resCompareProps.ulPropTag1 = MyEditor.GetHex(2);
				m_lpOutputRes->res.resCompareProps.ulPropTag2 = MyEditor.GetHex(4);
			}
		}
		break;
//These cases are essentially the same - resAnd and resOr are the same structure
	case RES_OR:
	case RES_AND:
		{
			CResAndOrEditor MyResEditor(
				this,
				lpSourceRes,
				m_lpAllocParent);//pass source res into editor
			WC_H(MyResEditor.DisplayDialog());

			if (S_OK == hRes)
			{
				m_lpOutputRes->rt = lpSourceRes->rt;
				m_lpOutputRes->res.resAnd.cRes = MyResEditor.GetResCount();

				LPSRestriction lpNewResArray = MyResEditor.DetachModifiedSRestrictionArray();
				if (lpNewResArray)
				{
					//I can do this because our new memory was allocated from a common parent
					MAPIFreeBuffer(m_lpOutputRes->res.resAnd.lpRes);
					m_lpOutputRes->res.resAnd.lpRes = lpNewResArray;
				}
			}
		}
		break;
	case RES_NOT:
		{
			CRestrictEditor MyResEditor(
				this,
				m_lpAllocParent,
				lpSourceRes->res.resNot.lpRes);//pass source res into editor
			WC_H(MyResEditor.DisplayDialog());

			if (S_OK == hRes)
			{
				m_lpOutputRes->rt = lpSourceRes->rt;
				LPSRestriction lpNewRes = MyResEditor.DetachModifiedSRestriction();
				if (lpNewRes)
				{
					//I can do this because our new memory was allocated from a common parent
					MAPIFreeBuffer(m_lpOutputRes->res.resNot.lpRes);
					m_lpOutputRes->res.resNot.lpRes = lpNewRes;
				}
			}
		}
		break;
//Structures for these two types are identical
	case RES_PROPERTY:
	case RES_CONTENT:
		{
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
				m_lpOutputRes->res.resContent.ulPropTag = MyEditor.GetHex(2);

				LPSPropValue lpNewProp = MyEditor.DetachModifiedSPropValue();

				if (lpNewProp)
				{
					//I can do this because our new memory was allocated from a common parent
					MAPIFreeBuffer(m_lpOutputRes->res.resContent.lpProp);
					m_lpOutputRes->res.resContent.lpProp = lpNewProp;
				}
				if (!m_lpOutputRes->res.resContent.lpProp)
				{
					//Got a problem here - the relop or fuzzy level was changed, but not the property
					//Need to copy the property from the source Res to the output Res
					EC_H(MAPIAllocateMore(
						sizeof(SPropValue),
						m_lpAllocParent,
						(LPVOID*)&m_lpOutputRes->res.resContent.lpProp));
					EC_H(PropCopyMore(
						m_lpOutputRes->res.resContent.lpProp,
						lpSourceRes->res.resContent.lpProp,
						MAPIAllocateMore,
						m_lpAllocParent));
				}
			}
		}
		break;
	case RES_BITMASK:
		{
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
				m_lpOutputRes->res.resBitMask.ulPropTag = MyEditor.GetHex(2);
				m_lpOutputRes->res.resBitMask.ulMask = MyEditor.GetHex(4);
			}
		}
		break;
	case RES_SIZE:
		{
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
				m_lpOutputRes->res.resSize.ulPropTag = MyEditor.GetHex(2);
				m_lpOutputRes->res.resSize.cb = MyEditor.GetHex(4);
			}
		}
		break;
	case RES_EXIST:
		{
			CResExistEditor MyEditor(
				this,
				lpSourceRes->res.resExist.ulPropTag);
			WC_H(MyEditor.DisplayDialog());
			if (S_OK == hRes)
			{
				m_lpOutputRes->rt = lpSourceRes->rt;
				m_lpOutputRes->res.resExist.ulPropTag = MyEditor.GetHex(0);
				m_lpOutputRes->res.resExist.ulReserved1 = 0;
				m_lpOutputRes->res.resExist.ulReserved2 = 0;
			}
		}
		break;
	case RES_SUBRESTRICTION:
		{
			CResSubResEditor MyEditor(
				this,
				lpSourceRes->res.resSub.ulSubObject,
				lpSourceRes->res.resSub.lpRes,
				m_lpAllocParent);
			WC_H(MyEditor.DisplayDialog());
			if (S_OK == hRes)
			{
				m_lpOutputRes->rt = lpSourceRes->rt;
				m_lpOutputRes->res.resSub.ulSubObject = MyEditor.GetHex(0);

				LPSRestriction lpNewRes = MyEditor.DetachModifiedSRestriction();
				if (lpNewRes)
				{
					//I can do this because our new memory was allocated from a common parent
					MAPIFreeBuffer(m_lpOutputRes->res.resSub.lpRes);
					m_lpOutputRes->res.resSub.lpRes = lpNewRes;
				}
			}
		}
		break;
	case RES_COMMENT:
		{
			CResCommentEditor MyResEditor(
				this,
				lpSourceRes,
				m_lpAllocParent);//pass source res into editor
			WC_H(MyResEditor.DisplayDialog());

			if (S_OK == hRes)
			{
				m_lpOutputRes->rt = lpSourceRes->rt;

				LPSRestriction lpNewRes = MyResEditor.DetachModifiedSRestriction();
				if (lpNewRes)
				{
					//I can do this because our new memory was allocated from a common parent
					MAPIFreeBuffer(m_lpOutputRes->res.resComment.lpRes);
					m_lpOutputRes->res.resComment.lpRes = lpNewRes;
				}

				LPSPropValue lpNewProps = MyResEditor.DetachModifiedSPropValue();
				if (lpNewProps)
				{
					//I can do this because our new memory was allocated from a common parent
					MAPIFreeBuffer(m_lpOutputRes->res.resComment.lpProp);
					m_lpOutputRes->res.resComment.lpProp = lpNewProps;
					m_lpOutputRes->res.resComment.cValues = MyResEditor.GetSPropValueCount();
				}
			}
		}
		break;
	}
	if (S_OK == hRes)
	{
		m_bModified = true;
		SetString(2,RestrictionToString(m_lpOutputRes,NULL));
	}
	return;
}


static TCHAR* CRITERIACLASS = _T("CCriteriaEditor");// STRING_OK
#define LISTNUM 4
CCriteriaEditor::CCriteriaEditor(
								 CWnd* pParentWnd,
								 LPSRestriction lpRes,
								 LPENTRYLIST lpEntryList,
								 ULONG ulSearchState):
CEditor(pParentWnd,IDS_CRITERIAEDITOR,IDS_CRITERIAEDITORPROMPT,6,CEDITOR_BUTTON_OK|CEDITOR_BUTTON_ACTION1|CEDITOR_BUTTON_CANCEL)
{
	TRACE_CONSTRUCTOR(CRITERIACLASS);

	HRESULT hRes = S_OK;

	m_uidActionButtonText1 = IDS_ACTIONEDITRES;
	m_lpSourceRes = lpRes;
	m_lpNewRes = NULL;
	m_lpSourceEntryList = lpEntryList;

	EC_H(MAPIAllocateBuffer(
		sizeof(SBinaryArray),
		(LPVOID*)&m_lpNewEntryList));

	m_ulNewSearchFlags = NULL;

	SetPromptPostFix(AllFlagsToString(flagSearchFlag,true));
	InitSingleLine(0,IDS_SEARCHSTATE,NULL,true);
	SetHex(0,ulSearchState);
	LPTSTR szFlags = NULL;
	EC_H(InterpretFlags(flagSearchState, ulSearchState, &szFlags));
	InitSingleLineSz(1,IDS_SEARCHSTATE,szFlags,true);
	MAPIFreeBuffer(szFlags);
	szFlags = NULL;
	InitSingleLine(2,IDS_SEARCHFLAGS,NULL,false);
	SetHex(2,0);
	InitSingleLine(3,IDS_SEARCHFLAGS,NULL,true);
	InitList(4,IDS_EIDLIST,false,false);
	InitMultiLine(5,IDS_RESTRICTIONTEXT,RestrictionToString(m_lpSourceRes,NULL),true);
}

//Used to call functions which need to be called AFTER controls are created
BOOL CCriteriaEditor::OnInitDialog()
{
	HRESULT hRes = S_OK;
	EC_B(CEditor::OnInitDialog());

	InitListFromEntryList(LISTNUM,m_lpSourceEntryList);

	UpdateListButtons();

	return HRES_TO_BOOL(hRes);
}

LPSRestriction CCriteriaEditor::GetSourceRes()
{
	if (m_lpNewRes) return m_lpNewRes;
	if (m_lpSourceRes) return m_lpSourceRes;
	return NULL;
}

ULONG CCriteriaEditor::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	if (2 == i)
	{
		HRESULT hRes = S_OK;
		LPTSTR szFlags = NULL;
		EC_H(InterpretFlags(flagSearchFlag, GetHexUseControl(i), &szFlags));
		SetString(3,szFlags);
		MAPIFreeBuffer(szFlags);
		szFlags = NULL;
	}
	return i;
}

LPSRestriction CCriteriaEditor::DetachModifiedSRestriction()
{
	LPSRestriction m_lpRet = m_lpNewRes;
	m_lpNewRes = NULL;
	return m_lpRet;
}

LPENTRYLIST CCriteriaEditor::DetachModifiedEntryList()
{
	LPENTRYLIST m_lpRet = m_lpNewEntryList;
	m_lpNewEntryList = NULL;
	return m_lpRet;
}

ULONG CCriteriaEditor::GetSearchFlags()
{
	return m_ulNewSearchFlags;
}


void CCriteriaEditor::InitListFromEntryList(ULONG ulListNum, LPENTRYLIST lpEntryList)
{
	if (!IsValidList(ulListNum)) return;

	HRESULT hRes = S_OK;
	m_lpControls[ulListNum].UI.lpList->List.DeleteAllColumns();
	EC_B(m_lpControls[ulListNum].UI.lpList->List.DeleteAllItems());

	InsertColumn(ulListNum,0,IDS_SHARP);
	CString szTmp;
	SortListData* lpData = NULL;
	ULONG i = 0;
	InsertColumn(ulListNum,1,IDS_CB);
	InsertColumn(ulListNum,2,IDS_BINARY);
	InsertColumn(ulListNum,3,IDS_TEXTVIEW);

	if (lpEntryList)
	{
		for (i = 0;i< lpEntryList->cValues;i++)
		{
			szTmp.Format(_T("%d"),i);// STRING_OK
			lpData = m_lpControls[ulListNum].UI.lpList->List.InsertRow(i,(LPTSTR)(LPCTSTR)szTmp);
			lpData->ulSortDataType = SORTLIST_BINARY;
			lpData->data.Binary.OldBin.cb = lpEntryList->lpbin[i].cb;
			lpData->data.Binary.OldBin.lpb = lpEntryList->lpbin[i].lpb;
			lpData->data.Binary.NewBin.cb = NULL;
			lpData->data.Binary.NewBin.lpb = NULL;

			szTmp.Format(_T("%d"),lpEntryList->lpbin[i].cb);// STRING_OK
			WC_B(m_lpControls[ulListNum].UI.lpList->List.SetItemText(i,1,(LPCTSTR)szTmp));
			WC_B(m_lpControls[ulListNum].UI.lpList->List.SetItemText(i,2,(LPCTSTR)BinToHexString(&lpEntryList->lpbin[i],false)));
			WC_B(m_lpControls[ulListNum].UI.lpList->List.SetItemText(i,3,(LPCTSTR)BinToTextString(&lpEntryList->lpbin[i],true)));
			lpData->bItemFullyLoaded = true;
		}
	}
	m_lpControls[ulListNum].UI.lpList->List.AutoSizeColumns();
}//CCriteriaEditor::InitListFromEntryList

void CCriteriaEditor::OnEditAction1()
{
	HRESULT hRes = S_OK;

	LPSRestriction lpSourceRes = GetSourceRes();

	CRestrictEditor MyResEditor(
		this,
		NULL,//m_lpAllocParent,
		lpSourceRes);//pass source res into editor
	WC_H(MyResEditor.DisplayDialog());

	if (S_OK == hRes)
	{
		MAPIFreeBuffer(m_lpNewRes);
		m_lpNewRes = MyResEditor.DetachModifiedSRestriction();
		SetString(5,RestrictionToString(m_lpNewRes,NULL));
	}
}

BOOL CCriteriaEditor::DoListEdit(ULONG ulListNum, int iItem, SortListData* lpData)
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
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);//m_lpAllocParent);

	LPSBinary lpSourcebin = NULL;
	if (lpData->data.Binary.OldBin.lpb)
	{
		lpSourcebin = &lpData->data.Binary.OldBin;
	}
	else
	{
		lpSourcebin = &lpData->data.Binary.NewBin;
	}

	BinEdit.InitSingleLineSz(0,IDS_EID,BinToHexString(lpSourcebin,false),false);

	WC_H(BinEdit.DisplayDialog());
	if (S_OK == hRes)
	{
		szTmp = BinEdit.GetString(0);
		ULONG ulStrLen = szTmp.GetLength();

		if (!(ulStrLen & 1))//can't use an odd length string
		{
			MAPIFreeBuffer(lpData->data.Binary.NewBin.lpb);

			lpData->data.Binary.NewBin.cb = ulStrLen / 2;
			EC_H(MAPIAllocateMore(
				lpData->data.Binary.NewBin.cb,
				m_lpNewEntryList,
				(LPVOID*)&lpData->data.Binary.NewBin.lpb));
			if (lpData->data.Binary.NewBin.lpb)
			{
				MyBinFromHex(
					(LPCTSTR) szTmp,
					lpData->data.Binary.NewBin.lpb,
					lpData->data.Binary.NewBin.cb);

				szTmp.Format(_T("%d"),lpData->data.Binary.NewBin.cb);// STRING_OK
				WC_B(m_lpControls[ulListNum].UI.lpList->List.SetItemText(iItem,1,(LPCTSTR)szTmp));
				WC_B(m_lpControls[ulListNum].UI.lpList->List.SetItemText(iItem,2,(LPCTSTR)BinToHexString(&lpData->data.Binary.NewBin,false)));
				WC_B(m_lpControls[ulListNum].UI.lpList->List.SetItemText(iItem,3,(LPCTSTR)BinToTextString(&lpData->data.Binary.NewBin,true)));
				return true;
			}
		}
	}

	return false;
}

void CCriteriaEditor::OnOK()
{
	CDialog::OnOK();//don't need to call CEditor::OnOK
	if (!IsValidList(LISTNUM)) return;
	ULONG ulValues = m_lpControls[LISTNUM].UI.lpList->List.GetItemCount();

	HRESULT hRes = S_OK;

	if (m_lpNewEntryList && ulValues < ULONG_MAX/sizeof(SBinary))
	{
		m_lpNewEntryList->cValues = ulValues;
		EC_H(MAPIAllocateMore(
			m_lpNewEntryList->cValues * sizeof(SBinary),
			m_lpNewEntryList,
			(LPVOID*)&m_lpNewEntryList->lpbin));

		ULONG i = 0;
		for (i=0;i<m_lpNewEntryList->cValues;i++)
		{
			SortListData* lpData = (SortListData*) m_lpControls[LISTNUM].UI.lpList->List.GetItemData(i);
			if (lpData->data.Binary.NewBin.lpb)
			{
				m_lpNewEntryList->lpbin[i].cb = lpData->data.Binary.NewBin.cb;
				m_lpNewEntryList->lpbin[i].lpb = lpData->data.Binary.NewBin.lpb;
				//clean out the source
				lpData->data.Binary.OldBin.lpb = NULL;
			}
			else
			{
				m_lpNewEntryList->lpbin[i].cb = lpData->data.Binary.OldBin.cb;
				EC_H(MAPIAllocateMore(
					m_lpNewEntryList->lpbin[i].cb,
					m_lpNewEntryList,
					(LPVOID*)&m_lpNewEntryList->lpbin[i].lpb));

				memcpy(m_lpNewEntryList->lpbin[i].lpb,lpData->data.Binary.OldBin.lpb,m_lpNewEntryList->lpbin[i].cb);
			}
		}
	}
	if (!m_lpNewRes && m_lpSourceRes)
	{
		EC_H(CopyRestriction(&m_lpNewRes,m_lpSourceRes,NULL));//m_lpAllocParent));
	}
	m_ulNewSearchFlags = GetHexUseControl(2);
}
