// ContentsTableDlg.cpp : implementation file
//

#include "stdafx.h"
#include "ContentsTableDlg.h"
#include "ContentsTableListCtrl.h"
#include "FakeSplitter.h"
#include "FileDialogEx.h"
#include "SingleMAPIPropListCtrl.h"
#include "MAPIFunctions.h"
#include "MAPIObjects.h"
#include "MFCUtilityFunctions.h"
#include "Editor.h"
#include "TagArrayEditor.h"
#include "InterpretProp2.h"
#include "RestrictEditor.h"
#include "PropertyTagEditor.h"
#include "PropTagArray.h"
#include "AttachmentsDlg.h"
#include "RecipientsDlg.h"

static TCHAR* CLASS = _T("CContentsTableDlg");

/////////////////////////////////////////////////////////////////////////////
// CContentsTableDlg dialog


CContentsTableDlg::CContentsTableDlg(
									 _In_ CParentWnd* pParentWnd,
									 _In_ CMapiObjects* lpMapiObjects,
									 UINT uidTitle,
									 __mfcmapiCreateDialogEnum bCreateDialog,
									 _In_opt_ LPMAPITABLE lpContentsTable,
									 _In_ LPSPropTagArray sptExtraColumnTags,
									 ULONG iNumExtraDisplayColumns,
									 _In_count_(iNumExtraDisplayColumns) TagNames* lpExtraDisplayColumns,
									 ULONG nIDContextMenu,
									 ULONG ulAddInContext
									 ):
CBaseDialog(
			pParentWnd,
			lpMapiObjects,
			ulAddInContext)
{
	TRACE_CONSTRUCTOR(CLASS);
	HRESULT hRes = S_OK;
	if (NULL != uidTitle)
	{
		EC_B(m_szTitle.LoadString(uidTitle));
	}
	else
	{
		EC_B(m_szTitle.LoadString(IDS_TABLEASCONTENTS));
	}

	m_lpContentsTableListCtrl = NULL;
	m_lpContainer = NULL;
	m_nIDContextMenu = nIDContextMenu;

	m_ulDisplayFlags = dfNormal;

	m_lpContentsTable = lpContentsTable;
	if (m_lpContentsTable) m_lpContentsTable->AddRef();

	m_sptExtraColumnTags = sptExtraColumnTags;
	m_iNumExtraDisplayColumns = iNumExtraDisplayColumns;
	m_lpExtraDisplayColumns = lpExtraDisplayColumns;

	if (mfcmapiCALL_CREATE_DIALOG == bCreateDialog)
	{
		CreateDialogAndMenu(NULL);
	}
} // CContentsTableDlg::CContentsTableDlg

CContentsTableDlg::~CContentsTableDlg()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_lpContentsTable) m_lpContentsTable->Release();
	m_lpContentsTable = NULL;
} // CContentsTableDlg::~CContentsTableDlg

_Check_return_ BOOL CContentsTableDlg::HandleMenu(WORD wMenuSelect)
{
	DebugPrint(DBGMenu,_T("CContentsTableDlg::HandleMenu wMenuSelect = 0x%X = %d\n"),wMenuSelect,wMenuSelect);
	switch (wMenuSelect)
	{
	case ID_APPLYFINDROW: SetRestrictionType(mfcmapiFINDROW_RESTRICTION); return true;
	case ID_APPLYRESTRICTION: SetRestrictionType(mfcmapiNORMAL_RESTRICTION); return true;
	case ID_CLEARRESTRICTION: SetRestrictionType(mfcmapiNO_RESTRICTION); return true;
	}

	return CBaseDialog::HandleMenu(wMenuSelect);
} // CContentsTableDlg::HandleMenu

_Check_return_ BOOL CContentsTableDlg::OnInitDialog()
{
	HRESULT	hRes = S_OK;
	BOOL	bRet = CBaseDialog::OnInitDialog();
	LPSPropValue	lpProp = NULL;
	ULONG			ulFlags = NULL;

	m_lpContentsTableListCtrl = new CContentsTableListCtrl(
		m_lpFakeSplitter,
		m_lpMapiObjects,
		m_sptExtraColumnTags,
		m_iNumExtraDisplayColumns,
		m_lpExtraDisplayColumns,
		m_nIDContextMenu,
		m_bIsAB,
		this);

	if (m_lpContentsTableListCtrl && m_lpFakeSplitter)
	{
		m_lpFakeSplitter->SetPaneOne(m_lpContentsTableListCtrl);
		m_lpFakeSplitter->SetPercent((FLOAT) 0.40);
		m_lpFakeSplitter->SetSplitType(SplitVertical);
	}

	if (m_lpContainer)
	{
		// Get a property for the title bar
		WC_H(HrGetOneProp(
			m_lpContainer,
			PR_DISPLAY_NAME,
			&lpProp));

		if (lpProp)
		{
			if (CheckStringProp(lpProp,PT_TSTRING))
			{
				m_szTitle = lpProp->Value.LPSZ;
			}
			else
			{
				EC_B(m_szTitle.LoadString(IDS_DISPLAYNAMENOTFOUND));
			}
			MAPIFreeBuffer(lpProp);
		}

		if (m_lpContentsTable) m_lpContentsTable->Release();
		m_lpContentsTable = NULL;

		ulFlags =
			(m_ulDisplayFlags & dfAssoc?MAPI_ASSOCIATED:NULL) |
			(m_ulDisplayFlags & dfDeleted?SHOW_SOFT_DELETES:NULL) |
			fMapiUnicode;

		hRes = S_OK;
		// Get the table of contents of the IMAPIContainer!!!
		EC_H(m_lpContainer->GetContentsTable(
			ulFlags,
			&m_lpContentsTable));
		if (hRes == MAPI_E_NO_ACCESS)
		{
			WARNHRESMSG(hRes,IDS_GETCONTENTSNOACCESS);
			m_lpContentsTable = NULL;
		}
		else if (hRes == MAPI_E_CALL_FAILED)
		{
			WARNHRESMSG(hRes,IDS_GETCONTENTSFAILED);
			m_lpContentsTable = NULL;
		}
		else if (hRes == MAPI_E_NO_SUPPORT)
		{
			WARNHRESMSG(hRes,IDS_GETCONTENTSNOTSUPPORTED);
			m_lpContentsTable = NULL;
		}
		else if (hRes == MAPI_E_UNKNOWN_FLAGS)
		{
			WARNHRESMSG(hRes,IDS_GETCONTENTSBADFLAG);
			m_lpContentsTable = NULL;
		}
		else if (hRes == MAPI_E_INVALID_PARAMETER)
		{
			WARNHRESMSG(hRes,IDS_GETCONTENTSBADPARAM);
			m_lpContentsTable = NULL;
		}
	}

	UpdateTitleBarText(NULL);

	return bRet;
} // CContentsTableDlg::OnInitDialog

void CContentsTableDlg::CreateDialogAndMenu(UINT nIDMenuResource)
{
	HRESULT hRes = S_OK;

	DebugPrintEx(DBGCreateDialog,CLASS,_T("CreateDialogAndMenu"),_T("id = 0x%X\n"),nIDMenuResource);
	CBaseDialog::CreateDialogAndMenu(nIDMenuResource);

	AddMenu(IDR_MENU_TABLE,IDS_TABLEMENU,(UINT)-1);

	if (m_lpContentsTableListCtrl && m_lpContentsTable)
	{
		ULONG ulPropType = NULL;

		ulPropType = GetMAPIObjectType(m_lpContainer);

		// Pass the contents table to the list control, but don't render yet - call BuildUIForContentsTable from CreateDialogAndMenu for that
		WC_H(m_lpContentsTableListCtrl->SetContentsTable(
			m_lpContentsTable,
			m_ulDisplayFlags,
			ulPropType));
	}
} // CContentsTableDlg::CreateDialogAndMenu

BEGIN_MESSAGE_MAP(CContentsTableDlg, CBaseDialog)
	ON_COMMAND(ID_DISPLAYSELECTEDITEM, OnDisplayItem)
	ON_COMMAND(ID_CANCELTABLELOAD,OnEscHit)
	ON_COMMAND(ID_REFRESHVIEW, OnRefreshView)
	ON_COMMAND(ID_CREATEMESSAGERESTRICTION, OnCreateMessageRestriction)
	ON_COMMAND(ID_CREATEPROPERTYSTRINGRESTRICTION, OnCreatePropertyStringRestriction)
	ON_COMMAND(ID_CREATERANGERESTRICTION, OnCreateRangeRestriction)
	ON_COMMAND(ID_EDITRESTRICTION, OnEditRestriction)
	ON_COMMAND(ID_GETSTATUS, OnGetStatus)
	ON_COMMAND(ID_OUTPUTTABLE,OnOutputTable)
	ON_COMMAND(ID_SETCOLUMNS,OnSetColumns)
	ON_COMMAND(ID_SORTTABLE, OnSortTable)
	ON_COMMAND(ID_TABLENOTIFICATIONON, OnNotificationOn)
	ON_COMMAND(ID_TABLENOTIFICATIONOFF, OnNotificationOff)
END_MESSAGE_MAP()

void CContentsTableDlg::OnInitMenu(_In_opt_ CMenu* pMenu)
{
	if (m_lpContentsTableListCtrl && pMenu)
	{
		int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();

		pMenu->EnableMenuItem(ID_CANCELTABLELOAD,DIM(m_lpContentsTableListCtrl->IsLoading()));
		pMenu->EnableMenuItem(ID_CREATEMESSAGERESTRICTION,DIM(1 == iNumSel && MAPI_FOLDER == m_lpContentsTableListCtrl->GetContainerType()));

		pMenu->EnableMenuItem(ID_DISPLAYSELECTEDITEM,DIMMSOK(iNumSel));

		__mfcmapiRestrictionTypeEnum RestrictionType = m_lpContentsTableListCtrl->GetRestrictionType();
		pMenu->CheckMenuItem(ID_APPLYFINDROW,CHECK(mfcmapiFINDROW_RESTRICTION == RestrictionType));
		pMenu->CheckMenuItem(ID_APPLYRESTRICTION,CHECK(mfcmapiNORMAL_RESTRICTION == RestrictionType));
		pMenu->CheckMenuItem(ID_CLEARRESTRICTION,CHECK(mfcmapiNO_RESTRICTION == RestrictionType));
		pMenu->EnableMenuItem(ID_TABLENOTIFICATIONON,DIM(m_lpContentsTableListCtrl->IsContentsTableSet() && !m_lpContentsTableListCtrl->IsAdviseSet()));
		pMenu->CheckMenuItem(ID_TABLENOTIFICATIONON,CHECK(m_lpContentsTableListCtrl->IsAdviseSet()));
		pMenu->EnableMenuItem(ID_TABLENOTIFICATIONOFF,DIM(m_lpContentsTableListCtrl->IsAdviseSet()));
		pMenu->EnableMenuItem(ID_OUTPUTTABLE,DIM(m_lpContentsTableListCtrl->IsContentsTableSet() && !m_lpContentsTableListCtrl->IsLoading()));

		pMenu->EnableMenuItem(ID_SETCOLUMNS,DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
		pMenu->EnableMenuItem(ID_SORTTABLE,DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
		pMenu->EnableMenuItem(ID_GETSTATUS,DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
		pMenu->EnableMenuItem(ID_CREATEMESSAGERESTRICTION,DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
		pMenu->EnableMenuItem(ID_CREATEPROPERTYSTRINGRESTRICTION,DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
		pMenu->EnableMenuItem(ID_CREATERANGERESTRICTION,DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
		pMenu->EnableMenuItem(ID_EDITRESTRICTION,DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
		pMenu->EnableMenuItem(ID_APPLYFINDROW,DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
		pMenu->EnableMenuItem(ID_APPLYRESTRICTION,DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
		pMenu->EnableMenuItem(ID_CLEARRESTRICTION,DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
		pMenu->EnableMenuItem(ID_REFRESHVIEW,DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));

		ULONG ulMenu = ID_ADDINMENU;
		for (ulMenu = ID_ADDINMENU; ulMenu < ID_ADDINMENU+m_ulAddInMenuItems ; ulMenu++)
		{
			LPMENUITEM lpAddInMenu = GetAddinMenuItem(m_hWnd,ulMenu);
			if (!lpAddInMenu) continue;

			ULONG ulFlags = lpAddInMenu->ulFlags;
			UINT uiEnable = MF_ENABLED;

			if ((ulFlags & MENU_FLAGS_SINGLESELECT) && iNumSel != 1) uiEnable = MF_GRAYED;
			if ((ulFlags & MENU_FLAGS_MULTISELECT) && !iNumSel) uiEnable = MF_GRAYED;
			EnableAddInMenus(pMenu, ulMenu, lpAddInMenu, uiEnable);
		}
	}
	CBaseDialog::OnInitMenu(pMenu);
} // CContentsTableDlg::OnInitMenu

void CContentsTableDlg::OnCancel()
{
	DebugPrintEx(DBGGeneric,CLASS,_T("OnCancel"),_T("\n"));
	// get rid of the window before we start our cleanup
	ShowWindow(SW_HIDE);

	if (m_lpContentsTableListCtrl) m_lpContentsTableListCtrl->OnCancelTableLoad();

	if (m_lpContentsTableListCtrl) m_lpContentsTableListCtrl->Release();
	m_lpContentsTableListCtrl = NULL;
	CBaseDialog::OnCancel();
} // CContentsTableDlg::OnCancel

void CContentsTableDlg::OnEscHit()
{
	DebugPrintEx(DBGGeneric,CLASS,_T("OnEscHit"),_T("\n"));
	if (m_lpContentsTableListCtrl)
	{
		m_lpContentsTableListCtrl->OnCancelTableLoad();
	}
} // CContentsTableDlg::OnEscHit

void CContentsTableDlg::SetRestrictionType(__mfcmapiRestrictionTypeEnum RestrictionType)
{
	if (m_lpContentsTableListCtrl)
		m_lpContentsTableListCtrl->SetRestrictionType(RestrictionType);
	OnRefreshView();
} // CContentsTableDlg::OnApplyFindRow

void CContentsTableDlg::OnDisplayItem()
{
	HRESULT			hRes = S_OK;
	LPMAPIPROP		lpMAPIProp = NULL;
	int				iItem = -1;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	do
	{
		hRes = S_OK;
		EC_H(m_lpContentsTableListCtrl->OpenNextSelectedItemProp(
			&iItem,
			mfcmapiREQUEST_MODIFY,
			&lpMAPIProp));

		if (lpMAPIProp)
		{
			EC_H(DisplayObject(
				lpMAPIProp,
				NULL,
				otHierarchy,
				this));
			lpMAPIProp->Release();
			lpMAPIProp = NULL;
		}
	}
	while (iItem != -1);
} // CContentsTableDlg::OnDisplayItem

// Clear the current list and get a new one with whatever code we've got in LoadMAPIPropList
void CContentsTableDlg::OnRefreshView()
{
	HRESULT hRes = S_OK;
	if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;
	DebugPrintEx(DBGGeneric,CLASS,_T("OnRefreshView"),_T("\n"));
	if (m_lpContentsTableListCtrl->IsLoading()) m_lpContentsTableListCtrl->OnCancelTableLoad();
	EC_H(m_lpContentsTableListCtrl->RefreshTable());
} // CContentsTableDlg::OnRefreshView

void CContentsTableDlg::OnNotificationOn()
{
	if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;
	HRESULT hRes = S_OK;
	EC_H(m_lpContentsTableListCtrl->NotificationOn());
} // CContentsTableDlg::OnNotificationOn

void CContentsTableDlg::OnNotificationOff()
{
	if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;
	m_lpContentsTableListCtrl->NotificationOff();
} // CContentsTableDlg::OnNotificationOff

// Read properties of the current message and save them for a restriction
// Designed to work with messages, but could be made to work with anything
void CContentsTableDlg::OnCreateMessageRestriction()
{
	if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;
	HRESULT			hRes = S_OK;
	ULONG			cVals = 0;
	LPSPropValue	lpProps = NULL;
	LPMAPIPROP		lpMAPIProp = NULL;

	LPSRestriction	lpRes = NULL;
	LPSRestriction	lpResLevel1 = NULL;
	LPSRestriction	lpResLevel2 = NULL;

	LPSPropValue	lpspvSubject = NULL;
	LPSPropValue	lpspvSubmitTime = NULL;
	LPSPropValue	lpspvDeliveryTime = NULL;

	// These are the properties we're going to copy off of the current message and store
	// in some object level variables
	enum{frPR_SUBJECT,
		frPR_CLIENT_SUBMIT_TIME,
		frPR_MESSAGE_DELIVERY_TIME,
		frNUMCOLS};
	SizedSPropTagArray(frNUMCOLS,sptFRCols) = {frNUMCOLS,
		PR_SUBJECT,
		PR_CLIENT_SUBMIT_TIME,
		PR_MESSAGE_DELIVERY_TIME};

	if (!m_lpContentsTableListCtrl) return;

	EC_H(m_lpContentsTableListCtrl->OpenNextSelectedItemProp(
		NULL,
		mfcmapiREQUEST_MODIFY,
		&lpMAPIProp));

	if (lpMAPIProp)
	{
		EC_H_GETPROPS(lpMAPIProp->GetProps(
			(LPSPropTagArray) &sptFRCols,
			fMapiUnicode,
			&cVals,
			&lpProps));
		if (lpProps)
		{
			// Allocate and create our SRestriction
			// Allocate base memory:
			EC_H(MAPIAllocateBuffer(
				sizeof(SRestriction),
				(LPVOID*)&lpRes));

			EC_H(MAPIAllocateMore(
				sizeof(SRestriction)*2,
				lpRes,
				(LPVOID*)&lpResLevel1));

			EC_H(MAPIAllocateMore(
				sizeof(SRestriction)*2,
				lpRes,
				(LPVOID*)&lpResLevel2));

			EC_H(MAPIAllocateMore(
				sizeof(SPropValue),
				lpRes,
				(LPVOID*)&lpspvSubject));

			EC_H(MAPIAllocateMore(
				sizeof(SPropValue),
				lpRes,
				(LPVOID*)&lpspvDeliveryTime));

			EC_H(MAPIAllocateMore(
				sizeof(SPropValue),
				lpRes,
				(LPVOID*)&lpspvSubmitTime));

			// Check that all our allocations were good before going on
			if (!FAILED(hRes))
			{
				// Zero out allocated memory.
				ZeroMemory(lpRes, sizeof(SRestriction));
				ZeroMemory(lpResLevel1, sizeof(SRestriction)*2);
				ZeroMemory(lpResLevel2, sizeof(SRestriction)*2);

				ZeroMemory(lpspvSubject, sizeof(SPropValue));
				ZeroMemory(lpspvSubmitTime, sizeof(SPropValue));
				ZeroMemory(lpspvDeliveryTime, sizeof(SPropValue));

				// Root Node
				lpRes->rt = RES_AND;				// We're doing an AND...
				lpRes->res.resAnd.cRes = 2;		// ...of two criteria...
				lpRes->res.resAnd.lpRes = lpResLevel1; // ...described here

				lpResLevel1[0].rt = RES_PROPERTY;
				lpResLevel1[0].res.resProperty.relop = RELOP_EQ;
				lpResLevel1[0].res.resProperty.ulPropTag = PR_SUBJECT;
				lpResLevel1[0].res.resProperty.lpProp = lpspvSubject;

				lpResLevel1[1].rt = RES_OR;
				lpResLevel1[1].res.resOr.cRes = 2;
				lpResLevel1[1].res.resOr.lpRes = lpResLevel2;

				lpResLevel2[0].rt = RES_PROPERTY;
				lpResLevel2[0].res.resProperty.relop = RELOP_EQ;
				lpResLevel2[0].res.resProperty.ulPropTag = PR_CLIENT_SUBMIT_TIME;
				lpResLevel2[0].res.resProperty.lpProp = lpspvSubmitTime;

				lpResLevel2[1].rt = RES_PROPERTY;
				lpResLevel2[1].res.resProperty.relop = RELOP_EQ;
				lpResLevel2[1].res.resProperty.ulPropTag = PR_MESSAGE_DELIVERY_TIME;
				lpResLevel2[1].res.resProperty.lpProp = lpspvDeliveryTime;

				// Allocate and fill out properties:
				lpspvSubject->ulPropTag = PR_SUBJECT;

				if (CheckStringProp(&lpProps[frPR_SUBJECT],PT_TSTRING))
				{
					EC_H(CopyString(
						&lpspvSubject->Value.LPSZ,
						lpProps[frPR_SUBJECT].Value.LPSZ,
						lpRes));
				}
				else lpspvSubject->Value.LPSZ = NULL;

				lpspvSubmitTime->ulPropTag = PR_CLIENT_SUBMIT_TIME;
				if (PR_CLIENT_SUBMIT_TIME == lpProps[frPR_CLIENT_SUBMIT_TIME].ulPropTag)
				{
					lpspvSubmitTime->Value.ft.dwLowDateTime = lpProps[frPR_CLIENT_SUBMIT_TIME].Value.ft.dwLowDateTime;
					lpspvSubmitTime->Value.ft.dwHighDateTime = lpProps[frPR_CLIENT_SUBMIT_TIME].Value.ft.dwHighDateTime;
				}
				else
				{
					lpspvSubmitTime->Value.ft.dwLowDateTime = 0x0;
					lpspvSubmitTime->Value.ft.dwHighDateTime = 0x0;
				}

				lpspvDeliveryTime->ulPropTag = PR_MESSAGE_DELIVERY_TIME;
				if (PR_MESSAGE_DELIVERY_TIME == lpProps[frPR_MESSAGE_DELIVERY_TIME].ulPropTag)
				{
					lpspvDeliveryTime->Value.ft.dwLowDateTime = lpProps[frPR_MESSAGE_DELIVERY_TIME].Value.ft.dwLowDateTime;
					lpspvDeliveryTime->Value.ft.dwHighDateTime = lpProps[frPR_MESSAGE_DELIVERY_TIME].Value.ft.dwHighDateTime;
				}
				else
				{
					lpspvDeliveryTime->Value.ft.dwLowDateTime = 0x0;
					lpspvDeliveryTime->Value.ft.dwHighDateTime = 0x0;
				}

				DebugPrintEx(DBGGeneric,CLASS,_T("OnCreateMessageRestriction"),_T("built restriction:\n"));
				DebugPrintRestriction(DBGGeneric,lpRes,lpMAPIProp);
			}
			else
			{
				// We failed in our allocations - clean up what we got before we apply
				// Everything was allocated off of lpRes, so cleaning up that will suffice
				if (lpRes) MAPIFreeBuffer(lpRes);
				lpRes = NULL;
			}
			m_lpContentsTableListCtrl->SetRestriction(lpRes);

			SetRestrictionType(mfcmapiNORMAL_RESTRICTION);
			MAPIFreeBuffer(lpProps);
		}
		lpMAPIProp->Release();
	}
} // CContentsTableDlg::OnCreateMessageRestriction

void CContentsTableDlg::OnCreatePropertyStringRestriction()
{
	HRESULT			hRes = S_OK;
	LPSRestriction	lpRes = NULL;

	if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;

	CPropertyTagEditor MyPropertyTag(
		IDS_CREATEPROPRES,
		NULL, // prompt
		PR_SUBJECT,
		m_bIsAB,
		m_lpContainer,
		this);

	WC_H(MyPropertyTag.DisplayDialog());
	if (S_OK == hRes)
	{
		CEditor MyData(
			this,
			IDS_SEARCHCRITERIA,
			IDS_CONTSEARCHCRITERIAPROMPT,
			3,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
		MyData.SetPromptPostFix(AllFlagsToString(flagFuzzyLevel,true));

		MyData.InitSingleLine(0,IDS_PROPVALUE,NULL,false);
		MyData.InitSingleLine(1,IDS_ULFUZZYLEVEL,NULL,false);
		MyData.SetHex(1,FL_IGNORECASE | FL_PREFIX);
		MyData.InitCheck(2,IDS_APPLYUSINGFINDROW,false,false);

		WC_H(MyData.DisplayDialog());
		if (S_OK != hRes) return;

		// Allocate and create our SRestriction
		EC_H(CreatePropertyStringRestriction(
			CHANGE_PROP_TYPE(MyPropertyTag.GetPropertyTag(),PT_TSTRING),
			MyData.GetString(0),
			MyData.GetHex(1),
			NULL,
			&lpRes));
		if (S_OK != hRes)
		{
			MAPIFreeBuffer(lpRes);
			lpRes = NULL;
			hRes = S_OK;
		}

		m_lpContentsTableListCtrl->SetRestriction(lpRes);

		if (MyData.GetCheck(2))
		{
			SetRestrictionType(mfcmapiFINDROW_RESTRICTION);
		}
		else
		{
			SetRestrictionType(mfcmapiNORMAL_RESTRICTION);
		}
	}
} // CContentsTableDlg::OnCreatePropertyStringRestriction

void CContentsTableDlg::OnCreateRangeRestriction()
{
	HRESULT			hRes = S_OK;
	LPSRestriction	lpRes = NULL;

	if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;

	CPropertyTagEditor MyPropertyTag(
		IDS_CREATERANGERES,
		NULL, // prompt
		PR_SUBJECT,
		m_bIsAB,
		m_lpContainer,
		this);

	WC_H(MyPropertyTag.DisplayDialog());
	if (S_OK == hRes)
	{
		CEditor MyData(
			this,
			IDS_SEARCHCRITERIA,
			IDS_RANGESEARCHCRITERIAPROMPT,
			2,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

		MyData.InitSingleLine(0,IDS_SUBSTRING,NULL,false);
		MyData.InitCheck(1,IDS_APPLYUSINGFINDROW,false,false);

		WC_H(MyData.DisplayDialog());
		if (S_OK != hRes) return;

		// Allocate and create our SRestriction
		EC_H(CreateRangeRestriction(
			CHANGE_PROP_TYPE(MyPropertyTag.GetPropertyTag(),PT_TSTRING),
			MyData.GetString(0),
			NULL,
			&lpRes));
		if (S_OK != hRes)
		{
			MAPIFreeBuffer(lpRes);
			lpRes = NULL;
			hRes = S_OK;
		}

		m_lpContentsTableListCtrl->SetRestriction(lpRes);

		if (MyData.GetCheck(1))
		{
			SetRestrictionType(mfcmapiFINDROW_RESTRICTION);
		}
		else
		{
			SetRestrictionType(mfcmapiNORMAL_RESTRICTION);
		}
	}
} // CContentsTableDlg::OnCreateRangeRestriction

void CContentsTableDlg::OnEditRestriction()
{
	HRESULT			hRes = S_OK;

	if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;

	CRestrictEditor MyRestrict(
		this,
		NULL, // No alloc parent - we must MAPIFreeBuffer the result
		m_lpContentsTableListCtrl->GetRestriction());

	WC_H(MyRestrict.DisplayDialog());
	if (S_OK != hRes) return;

	m_lpContentsTableListCtrl->SetRestriction(MyRestrict.DetachModifiedSRestriction());
} // CContentsTableDlg::OnEditRestriction

void CContentsTableDlg::OnOutputTable()
{
	if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;
	HRESULT	hRes = S_OK;
	WCHAR*	szFileName = NULL;
	INT_PTR	iDlgRet = 0;

	CStringW szFileSpec;
	EC_B(szFileSpec.LoadString(IDS_TEXTFILES));

	CFileDialogExW dlgFilePicker;

	EC_D_DIALOG(dlgFilePicker.DisplayDialog(
		FALSE, // Save As dialog
		L"txt", // STRING_OK
		L"table.txt", // STRING_OK
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		szFileSpec,
		this));
	if (IDOK == iDlgRet)
	{
		szFileName = dlgFilePicker.GetFileName();

		if (szFileName)
		{
			DebugPrintEx(DBGGeneric,CLASS,_T("OnOutputTable"),_T("saving to %ws\n"),szFileName);

			m_lpContentsTableListCtrl->OnOutputTable(szFileName);
		}
	}
} // CContentsTableDlg::OnOutputTable

void CContentsTableDlg::OnSetColumns()
{
	if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;
	m_lpContentsTableListCtrl->DoSetColumns(
		false,
		true,
		true,
		true);
} // CContentsTableDlg::OnSetColumns

void CContentsTableDlg::OnGetStatus()
{
	if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;
	m_lpContentsTableListCtrl->GetStatus();
} // CContentsTableDlg::OnGetStatus

void CContentsTableDlg::OnSortTable()
{
	HRESULT			hRes = S_OK;

	if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;

	CEditor MyData(
		this,
		IDS_SORTTABLE,
		IDS_SORTTABLEPROMPT1,
		6,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	MyData.InitSingleLine(0,IDS_CSORTS,NULL,false);
	MyData.InitSingleLine(1,IDS_CCATS,NULL,false);
	MyData.InitSingleLine(2,IDS_CEXPANDED,NULL,false);
	MyData.InitCheck(3,IDS_TBLASYNC,false,false);
	MyData.InitCheck(4,IDS_TBLBATCH,false,false);
	MyData.InitCheck(5,IDS_REFRESHAFTERSORT,true,false);

	WC_H(MyData.DisplayDialog());
	if (S_OK != hRes) return;

	ULONG cSorts = MyData.GetDecimal(0);
	ULONG cCategories = MyData.GetDecimal(1);
	ULONG cExpanded = MyData.GetDecimal(2);

	if (cSorts < cCategories || cCategories < cExpanded)
	{
		ErrDialog(__FILE__,__LINE__,IDS_EDBADSORTS,cSorts,cCategories,cExpanded);
		return;
	}

	LPSSortOrderSet lpMySortOrders = NULL;

	EC_H(MAPIAllocateBuffer(CbNewSSortOrderSet(cSorts),(LPVOID*)&lpMySortOrders));

	if (lpMySortOrders)
	{
		lpMySortOrders->cSorts = cSorts;
		lpMySortOrders->cCategories = cCategories;
		lpMySortOrders->cExpanded = cExpanded;
		BOOL bNoError = true;

		ULONG i = 0;
		for (i = 0 ; i < cSorts ; i++)
		{
			CPropertyTagEditor MyPropertyTag(
				NULL, // title
				NULL, // prompt
				NULL,
				m_bIsAB,
				m_lpContainer,
				this);

			WC_H(MyPropertyTag.DisplayDialog());
			if (S_OK == hRes)
			{
				lpMySortOrders->aSort[i].ulPropTag = MyPropertyTag.GetPropertyTag();
				CEditor MySortOrderDlg(
					this,
					IDS_SORTORDER,
					IDS_SORTORDERPROMPT,
					1,
					CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
				UINT uidDropDown[] = {
					IDS_DDTABLESORTASCEND,
					IDS_DDTABLESORTDESCEND,
					IDS_DDTABLESORTCOMBINE,
					IDS_DDTABLESORTCATEGMAX,
					IDS_DDTABLESORTCATEGMIN
				};
				MySortOrderDlg.InitDropDown(0,IDS_SORTORDER,_countof(uidDropDown),uidDropDown,true);

				WC_H(MySortOrderDlg.DisplayDialog());
				if (S_OK == hRes)
				{
					switch (MySortOrderDlg.GetDropDown(0))
					{
					default:
					case 0:
						lpMySortOrders->aSort[i].ulOrder = TABLE_SORT_ASCEND;
						break;
					case 1:
						lpMySortOrders->aSort[i].ulOrder = TABLE_SORT_DESCEND;
						break;
					case 2:
						lpMySortOrders->aSort[i].ulOrder = TABLE_SORT_COMBINE;
						break;
					case 3:
						lpMySortOrders->aSort[i].ulOrder = TABLE_SORT_CATEG_MAX;
						break;
					case 4:
						lpMySortOrders->aSort[i].ulOrder = TABLE_SORT_CATEG_MIN;
						break;
					}
				}
				else bNoError = false;
			}
			else bNoError = false;
		}

		if (bNoError)
		{
			EC_H(m_lpContentsTableListCtrl->SetSortTable(
				lpMySortOrders,
				(MyData.GetCheck(3)?TBL_ASYNC:0) | (MyData.GetCheck(4)?TBL_BATCH: 0) // flags
				));
		}
	}
	MAPIFreeBuffer(lpMySortOrders);

	if (MyData.GetCheck(5)) EC_H(m_lpContentsTableListCtrl->RefreshTable());
} // CContentsTableDlg::OnSortTable

// Since the strategy for opening the selected property may vary depending on the table we're displaying,
// this virtual function allows us to override the default method with the method used by the table we've written a special class for.
_Check_return_ HRESULT CContentsTableDlg::OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp)
{
	HRESULT hRes = S_OK;
	DebugPrintEx(DBGOpenItemProp,CLASS,_T("OpenItemProp"),_T("iSelectedItem = 0x%X\n"),iSelectedItem);

	if (!lppMAPIProp || !m_lpContentsTableListCtrl) return MAPI_E_INVALID_PARAMETER;

	if (-1 == iSelectedItem)
	{
		// Get the first selected item
		EC_H(m_lpContentsTableListCtrl->OpenNextSelectedItemProp(
			NULL,
			bModify,
			lppMAPIProp));
	}
	else
	{
		EC_H(m_lpContentsTableListCtrl->DefaultOpenItemProp(
			iSelectedItem,
			bModify,
			lppMAPIProp));
	}

	return hRes;
} // CContentsTableDlg::OpenItemProp

_Check_return_ HRESULT CContentsTableDlg::OpenAttachmentsFromMessage(_In_ LPMESSAGE lpMessage, BOOL fSaveMessageAtClose)
{
	HRESULT hRes = S_OK;
	LPMAPITABLE	lpTable = NULL;

	if (NULL == lpMessage) return MAPI_E_INVALID_PARAMETER;

	EC_H(lpMessage->OpenProperty(
		PR_MESSAGE_ATTACHMENTS,
		&IID_IMAPITable,
		0,
		0,
		(LPUNKNOWN *) &lpTable));

	if (lpTable)
	{
		new CAttachmentsDlg(
			m_lpParent,
			m_lpMapiObjects,
			lpTable,
			lpMessage,
			fSaveMessageAtClose);
		lpTable->Release();
	}

	return hRes;
} // CContentsTableDlg::OpenAttachmentsFromMessage

_Check_return_ HRESULT CContentsTableDlg::OpenRecipientsFromMessage(LPMESSAGE lpMessage)
{
	HRESULT hRes = S_OK;
	LPMAPITABLE	lpTable = NULL;

	EC_H(lpMessage->OpenProperty(
		PR_MESSAGE_RECIPIENTS,
		&IID_IMAPITable,
		0,
		0,
		(LPUNKNOWN *) &lpTable));

	if (lpTable)
	{
		new CRecipientsDlg(
			m_lpParent,
			m_lpMapiObjects,
			lpTable,
			lpMessage);
		lpTable->Release();
	}

	return hRes;
} // CContentsTableDlg::OpenRecipientsFromMessage

_Check_return_ BOOL CContentsTableDlg::HandleAddInMenu(WORD wMenuSelect)
{
	if (wMenuSelect < ID_ADDINMENU || ID_ADDINMENU+m_ulAddInMenuItems < wMenuSelect) return false;
	if (!m_lpContentsTableListCtrl) return false;
	LPMAPIPROP		lpMAPIProp = NULL;
	int				iItem = -1;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	LPMENUITEM lpAddInMenu = GetAddinMenuItem(m_hWnd,wMenuSelect);
	if (!lpAddInMenu) return false;

	ULONG ulFlags = lpAddInMenu->ulFlags;

	__mfcmapiModifyEnum	fRequestModify =
		(ulFlags & MENU_FLAGS_REQUESTMODIFY)?mfcmapiREQUEST_MODIFY:mfcmapiDO_NOT_REQUEST_MODIFY;

	// Get the stuff we need for any case
	_AddInMenuParams MyAddInMenuParams = {0};
	MyAddInMenuParams.lpAddInMenu = lpAddInMenu;
	MyAddInMenuParams.ulAddInContext = m_ulAddInContext;
	MyAddInMenuParams.hWndParent = m_hWnd;
	if (m_lpMapiObjects)
	{
		MyAddInMenuParams.lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		MyAddInMenuParams.lpMDB = m_lpMapiObjects->GetMDB(); // do not release
		MyAddInMenuParams.lpAdrBook = m_lpMapiObjects->GetAddrBook(false); // do not release
	}

	if (m_lpPropDisplay)
	{
		m_lpPropDisplay->GetSelectedPropTag(&MyAddInMenuParams.ulPropTag);
	}

	// MENU_FLAGS_SINGLESELECT and MENU_FLAGS_MULTISELECT can't both be set, so we can ignore this case
	if (!(ulFlags & (MENU_FLAGS_SINGLESELECT|MENU_FLAGS_MULTISELECT)))
	{
		HandleAddInMenuSingle(
			&MyAddInMenuParams,
			NULL,
			NULL);
	}
	else
	{
		// Add appropriate flag to context
		MyAddInMenuParams.ulCurrentFlags |= (ulFlags & (MENU_FLAGS_SINGLESELECT|MENU_FLAGS_MULTISELECT));
		while (true)
		{
			SortListData* lpData = (SortListData*) m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
			if (-1 == iItem) break;
			SRow MyRow = {0};

			// If we have a row to give, give it - it's free
			if (lpData)
			{
				MyRow.cValues = lpData->cSourceProps;
				MyRow.lpProps = lpData->lpSourceProps;
				MyAddInMenuParams.lpRow = &MyRow;
				MyAddInMenuParams.ulCurrentFlags |= MENU_FLAGS_ROW;
			}

			if (!(ulFlags & MENU_FLAGS_ROW))
			{
				(void) OpenItemProp(iItem, fRequestModify, &lpMAPIProp);
			}

			HandleAddInMenuSingle(
				&MyAddInMenuParams,
				lpMAPIProp,
				NULL);
			if (lpMAPIProp) lpMAPIProp->Release();

			// If we're not doing multiselect, then we're done after a single pass
			if (!(ulFlags & MENU_FLAGS_MULTISELECT)) break;
		}
	}
	return true;
} // CContentsTableDlg::HandleAddInMenu

void CContentsTableDlg::HandleAddInMenuSingle(
	_In_ LPADDINMENUPARAMS lpParams,
	_In_opt_ LPMAPIPROP lpMAPIProp,
	_In_opt_ LPMAPICONTAINER /*lpContainer*/)
{
	if (lpParams)
	{
		lpParams->lpTable = m_lpContentsTable;
		switch(lpParams->ulAddInContext)
		{
		case MENU_CONTEXT_RECIEVE_FOLDER_TABLE:
			lpParams->lpFolder = (LPMAPIFOLDER) lpMAPIProp; // OpenItemProp returns LPMAPIFOLDER
			break;
		case MENU_CONTEXT_HIER_TABLE:
			lpParams->lpFolder = (LPMAPIFOLDER) lpMAPIProp; // OpenItemProp returns LPMAPIFOLDER
			break;
		}
	}

	InvokeAddInMenu(lpParams);
} // CContentsTableDlg::HandleAddInMenuSingle