// SingleMAPIPropListCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "Error.h"
#include "registry.h"

#include "SortListCtrl.h"
#include "SingleMAPIPropListCtrl.h"
#include "StreamEditor.h"

#include "BaseDialog.h"
#include "MAPIFunctions.h"
#include "ColumnTags.h"
#include "MFCUtilityFunctions.h"
#include "MapiObjects.h"
#include "MySecInfo.h"
#include "Editor.h"
#include "PropertyEditor.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"
#include "File.h"
#include "TagArrayEditor.h"

#include "FileDialogEx.h"
#include "ImportProcs.h"
#include "RestrictEditor.h"
#include "PropertyTagEditor.h"
#include "MAPIProgress.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

static TCHAR* CLASS = _T("CSingleMAPIPropListCtrl");

// 26 columns should be enough for anybody
#define MAX_SORT_COLS 26

/////////////////////////////////////////////////////////////////////////////
// CSingleMAPIPropListCtrl

CSingleMAPIPropListCtrl::CSingleMAPIPropListCtrl(
												 CWnd* pCreateParent,
												 CBaseDialog *lpHostDlg,
												 BOOL bIsAB)
												 :CSortListCtrl()
{
	TRACE_CONSTRUCTOR(CLASS);
	HRESULT hRes = S_OK;

	EC_H(Create(pCreateParent,LVS_SINGLESEL,IDC_LIST_CTRL,true));

	m_lpMAPIProp = NULL;

	m_bHaveEverDisplayedSomething = FALSE;

	m_lpHostDlg = lpHostDlg;
	if (m_lpHostDlg) m_lpHostDlg->AddRef();

	for (ULONG i = 0;i<NUMPROPCOLUMNS;i++)
	{
		CString szHeaderName;
		szHeaderName.LoadString(PropColumns[i].uidName);
		InsertColumn(i,szHeaderName);
	}

	CHeaderCtrl* lpMyHeader = GetHeaderCtrl();

	// Column orders are stored as lowercase letters
	// bacdefghi would mean the first two columns are swapped
	if (lpMyHeader && RegKeys[regkeyPROP_COLUMN_ORDER].szCurSTRING[0] != '\0')
	{
		BOOL bSetCols = false;
		int nColumnCount = lpMyHeader->GetItemCount();
		size_t cchOrder = NULL;
		WC_H(StringCchLength(RegKeys[regkeyPROP_COLUMN_ORDER].szCurSTRING,MAX_SORT_COLS,&cchOrder));
		if (SUCCEEDED(hRes) && nColumnCount == (int) cchOrder)
		{
			LPINT pnOrder = new int[nColumnCount];

			if (pnOrder)
			{
				int i = 0;
				for (i = 0; i<nColumnCount; i++)
				{
					pnOrder[i] = RegKeys[regkeyPROP_COLUMN_ORDER].szCurSTRING[i] - 'a';
				}
				if (SetColumnOrderArray(nColumnCount,pnOrder))
				{
					bSetCols = true;
				}
			}
			delete[] pnOrder;
		}
		// If we didn't like the reg key, clear it so we don't see it again
		if (!bSetCols) RegKeys[regkeyPROP_COLUMN_ORDER].szCurSTRING[0] = '\0';
	}

	AutoSizeColumns();

	m_sptExtraProps = NULL;

	m_bIsAB = bIsAB;
	m_lpMAPIProp = NULL;
	m_lpSourceData = NULL;
	m_bRowModified = false;
}

CSingleMAPIPropListCtrl::~CSingleMAPIPropListCtrl()
{
	TRACE_DESTRUCTOR(CLASS);
	MAPIFreeBuffer(m_sptExtraProps);
	if (m_lpMAPIProp) m_lpMAPIProp->Release();
	if (m_lpHostDlg) m_lpHostDlg->Release();
}

BEGIN_MESSAGE_MAP(CSingleMAPIPropListCtrl, CSortListCtrl)
//{{AFX_MSG_MAP(CSingleMAPIPropListCtrl)
	ON_NOTIFY_REFLECT(NM_DBLCLK, OnDblclk)
	ON_WM_KEYDOWN()
	ON_WM_CONTEXTMENU()
	ON_MESSAGE(WM_MFCMAPI_SAVECOLUMNORDERLIST, msgOnSaveColumnOrder)
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CSingleMAPIPropListCtrl message handlers

//Some message handling's not worth breaking out into separate functions.
LRESULT CSingleMAPIPropListCtrl::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	/*
	HRESULT hRes = S_OK;
	LRESULT lResult = NULL;
	DebugPrint(DBGWindowProc,_T("CSingleMAPIPropListCtrl::WindowProc message = 0x%x, wParam = 0x%X, lParam = 0x%X\n"),message,wParam,lParam);

	switch (message)
	{
	case WM_LBUTTONDBLCLK:
		break;
	case WM_CONTEXTMENU:
		break;
	}//end switch
	*/
	return CSortListCtrl::WindowProc(message,wParam,lParam);
}

// WM_MFCMAPI_SAVECOLUMNORDERLIST
LRESULT	CSingleMAPIPropListCtrl::msgOnSaveColumnOrder(WPARAM /*wParam*/, LPARAM /*lParam*/)
{
	HRESULT hRes = S_OK;
	CHeaderCtrl* lpMyHeader = GetHeaderCtrl();

	if (lpMyHeader)
	{
		ULONG nColumnCount = lpMyHeader->GetItemCount();
		if (nColumnCount && nColumnCount <= MAX_SORT_COLS)
		{
			LPINT pnOrder = new int[nColumnCount];

			if (pnOrder)
			{
				EC_B(GetColumnOrderArray(pnOrder, nColumnCount));
				ULONG i = 0;
				for (i = 0; i<nColumnCount; i++)
					RegKeys[regkeyPROP_COLUMN_ORDER].szCurSTRING[i] = (char)('a' + pnOrder[i]);
				RegKeys[regkeyPROP_COLUMN_ORDER].szCurSTRING[nColumnCount] = '\0';
			}
			delete[] pnOrder;
		}
	}
	return S_OK;
}

void CSingleMAPIPropListCtrl::InitMenu(CMenu* pMenu)
{
	if (pMenu)
	{
		ULONG ulPropTag = NULL;
		BOOL bPropSelected = FALSE;

		GetSelectedPropTag(&ulPropTag);
		bPropSelected = (NULL != ulPropTag);

		if (m_lpHostDlg && m_lpHostDlg->m_lpMapiObjects)
		{
			ULONG ulStatus = m_lpHostDlg->m_lpMapiObjects->GetBufferStatus();
			LPENTRYLIST lpEIDsToCopy = m_lpHostDlg->m_lpMapiObjects->GetMessagesToCopy();
			pMenu->EnableMenuItem(ID_PASTE_PROPERTY,DIM(m_lpMAPIProp && (ulStatus & BUFFER_PROPTAG) && (ulStatus & BUFFER_SOURCEPROPOBJ)));
			pMenu->EnableMenuItem(ID_COPYTO,DIM(m_lpMAPIProp && (ulStatus & BUFFER_SOURCEPROPOBJ)));
			pMenu->EnableMenuItem(ID_PASTE_NAMEDPROPS, DIM(m_lpMAPIProp  && (ulStatus & BUFFER_MESSAGES) && lpEIDsToCopy && 1 == lpEIDsToCopy->cValues));
		}
		pMenu->EnableMenuItem(ID_COPY_PROPERTY,DIM(m_lpMAPIProp));

		pMenu->EnableMenuItem(ID_DELETEPROPERTY,DIM(m_lpMAPIProp && bPropSelected));
		pMenu->EnableMenuItem(ID_DISPLAYPROPERTYASSECURITYDESCRIPTORPROPSHEET,DIM(m_lpMAPIProp && bPropSelected && pfnEditSecurity));
		pMenu->EnableMenuItem(ID_DISPLAYPROPERTYASSECURITYDESCRIPTOR,DIM(m_lpMAPIProp && bPropSelected));
		pMenu->EnableMenuItem(ID_DISPLAYPROPERTYASTZREG,DIM(m_lpMAPIProp && bPropSelected));
		pMenu->EnableMenuItem(ID_DISPLAYPROPERTYASTZDEFINITION,DIM(m_lpMAPIProp && bPropSelected));
		pMenu->EnableMenuItem(ID_EDITPROPASBINARYSTREAM,DIM(m_lpMAPIProp && bPropSelected));
		pMenu->EnableMenuItem(ID_EDITPROPERTY,DIM(bPropSelected));
		pMenu->EnableMenuItem(ID_EDITPROPERTYASASCIISTREAM,DIM(m_lpMAPIProp && bPropSelected));
		pMenu->EnableMenuItem(ID_EDITPROPERTYASUNICODESTREAM,DIM(m_lpMAPIProp && bPropSelected));
		pMenu->EnableMenuItem(ID_EDITPROPERTYASPRRTFCOMPRESSEDSTREAM,DIM(m_lpMAPIProp && bPropSelected));
		pMenu->EnableMenuItem(ID_OPEN_PROPERTY,DIM(bPropSelected));

		pMenu->EnableMenuItem(ID_SAVEPROPERTIES,DIM(m_lpMAPIProp || GetPropVals()));
		pMenu->EnableMenuItem(ID_EDITGIVENPROPERTY,DIM(m_lpMAPIProp || GetPropVals()));
		pMenu->EnableMenuItem(ID_OPENPROPERTYASTABLE,DIM(m_lpMAPIProp));
		pMenu->EnableMenuItem(ID_FINDALLNAMEDPROPS,DIM(m_lpMAPIProp));

		if (m_lpHostDlg)
		{
			ULONG ulMenu = ID_ADDINPROPERTYMENU;
			for (ulMenu = ID_ADDINPROPERTYMENU ;; ulMenu++)
			{
				LPMENUITEM lpAddInMenu = GetAddinMenuItem(m_lpHostDlg->m_hWnd,ulMenu);
				if (!lpAddInMenu) break;

				UINT uiEnable = DIM(bPropSelected);

				pMenu->EnableMenuItem(ulMenu,uiEnable);
			}
		}
	}
}

BOOL CSingleMAPIPropListCtrl::HandleMenu(WORD wMenuSelect)
{
	DebugPrint(DBGMenu,_T("CSingleMAPIPropListCtrl::HandleMenu wMenuSelect = 0x%X = %d\n"),wMenuSelect,wMenuSelect);
//	HRESULT hRes = S_OK;
	switch (wMenuSelect)
	{
	case ID_COPY_PROPERTY: OnCopyProperty(); return true;
	case ID_COPYTO: OnCopyTo(); return true;
	case ID_DELETEPROPERTY: OnDeleteProperty(); return true;
	case ID_DISPLAYPROPERTYASSECURITYDESCRIPTOR: OnDisplayPropertyAsSecurityDescriptor(); return true;
	case ID_DISPLAYPROPERTYASSECURITYDESCRIPTORPROPSHEET: OnDisplayPropertyAsSecurityDescriptorPropSheet(); return true;
	case ID_DISPLAYPROPERTYASTZREG: OnDisplayPropAsTZREG(); return true;
	case ID_DISPLAYPROPERTYASTZDEFINITION: OnDisplayPropAsTZDEFINITION(); return true;
	case ID_EDITGIVENPROPERTY: OnEditGivenProperty(); return true;
	case ID_EDITPROPERTY: OnEditProp(); return true;
	case ID_EDITPROPASBINARYSTREAM: OnEditPropAsStream(PT_BINARY,false); return true;
	case ID_EDITPROPERTYASASCIISTREAM: OnEditPropAsStream(PT_STRING8,false); return true;
	case ID_EDITPROPERTYASUNICODESTREAM: OnEditPropAsStream(PT_UNICODE,false); return true;
	case ID_EDITPROPERTYASPRRTFCOMPRESSEDSTREAM: OnEditPropAsStream(PT_BINARY,true); return true;
	case ID_FINDALLNAMEDPROPS: FindAllNamedProps(); return true;
	case ID_MODIFYEXTRAPROPS: OnModifyExtraProps(); return true;
	case ID_OPEN_PROPERTY: OnOpenProperty(); return true;
	case ID_OPENPROPERTYASTABLE: OnOpenPropertyAsTable(); return true;
	case ID_PASTE_NAMEDPROPS: OnPasteNamedProps(); return true;
	case ID_PASTE_PROPERTY: OnPasteProperty(); return true;
	case ID_SAVEPROPERTIES: SavePropsToXML(); return true;
	}

	return HandleAddInMenu(wMenuSelect);
}

void CSingleMAPIPropListCtrl::GetSelectedPropTag(ULONG* lpPropTag)
{
	int	iItem = NULL;

	*lpPropTag = NULL;

	iItem = GetNextItem(-1,LVNI_FOCUSED | LVNI_SELECTED);

	if (-1 != iItem)
	{
		SortListData* lpListData = ((SortListData*)GetItemData(iItem));
		if (lpListData)
			*lpPropTag = lpListData->data.Prop.ulPropTag;
	}

	if (lpPropTag)
	{
		DebugPrintEx(DBGGeneric,CLASS,_T("GetSelectedPropTag"),_T("returning lpPropTag = 0x%X\n"),*lpPropTag);
	}
}//CSingleMAPIPropListCtrl::GetSelectedPropTag

//Call GetProps with NULL to get a list of (almost) all properties.
//Parse this list and render them in the control.
//Add any extra props we've asked for through the UI
HRESULT CSingleMAPIPropListCtrl::LoadMAPIPropList()
{
	HRESULT			hRes = S_OK;
	ULONG			ulCurListBoxRow = 0;
	ULONG			ulCurPropRow = 0;
	CWaitCursor		Wait;//Change the mouse to an hourglass while we work.

	LPSPropValue	lpPropVals = GetPropVals();//do not free!
	if (!m_lpMAPIProp && !lpPropVals) return MAPI_E_INVALID_PARAMETER;

	ulCurListBoxRow = 0;

	if (GetCountPropVals() &&
		lpPropVals &&
		(!m_lpMAPIProp || RegKeys[regkeyUSE_ROW_DATA_FOR_SINGLEPROPLIST].ulCurDWORD))
	{
		if (m_lpHostDlg)
			m_lpHostDlg->UpdateStatusBarText(STATUSRIGHTPANE,IDS_PROPSFROMROW);

		//get each property in turn and add it to the list
		for (ulCurPropRow = 0; ulCurPropRow < GetCountPropVals(); ulCurPropRow++)
		{
			AddPropToListBox(
				ulCurListBoxRow,
				lpPropVals[ulCurPropRow].ulPropTag,
				&lpPropVals[ulCurPropRow]);

			ulCurListBoxRow++;
		}
	}
	else if (m_lpMAPIProp)
	{
		if (m_lpHostDlg)
			m_lpHostDlg->UpdateStatusBarText(STATUSRIGHTPANE,IDS_PROPSFROMGETPROPS);
		if (RegKeys[regkeyDO_GETPROPS].ulCurDWORD)
		{
			ULONG			cProps = 0;
			LPSPropValue	pProps = NULL;
			//Can't use EC_H_GETPROPS because I want to suppress MAPI_E_CALL_FAILED as well
			hRes = GetPropsNULL(m_lpMAPIProp,
				fMapiUnicode,
				&cProps,
				&pProps);
			if (MAPI_E_CALL_FAILED == hRes)
			{
				//Some stores, like public folders, don't support properties on the root folder
				DebugPrint(DBGGeneric,_T("Failed to get call GetProps on this object!\n"));
			}
			else if (FAILED(hRes))//only report errors, not warnings
			{
				CHECKHRESMSG(hRes,IDS_GETPROPSNULLFAILED);
			}
			else if (pProps)
			{
				//get each property in turn and add it to the list
				for (ulCurPropRow = 0; ulCurPropRow < cProps; ulCurPropRow++)
				{
					AddPropToListBox(
						ulCurListBoxRow,
						pProps[ulCurPropRow].ulPropTag,
						&pProps[ulCurPropRow]);

					ulCurListBoxRow++;
				}
			}
			MAPIFreeBuffer(pProps);
		}

		if (m_sptExtraProps)
		{
			//Let's get each extra property one at a time
			ULONG			cExtraProps = 0;
			LPSPropValue	pExtraProps = NULL;
			SPropValue		ExtraPropForList;
			ULONG			iCurExtraProp = 0;
			SPropTagArray	pNewTag;
			pNewTag.cValues = 1;

			for (iCurExtraProp = 0 ; iCurExtraProp < m_sptExtraProps->cValues ; iCurExtraProp++)
			{
				hRes = S_OK;//clear the error flag before each run
				pNewTag.aulPropTag[0] = m_sptExtraProps->aulPropTag[iCurExtraProp];

				//Let's add some extra properties
				//Don't need to report since we're gonna put show the error in the UI
				WC_H(m_lpMAPIProp->GetProps(
					&pNewTag,
					fMapiUnicode,
					&cExtraProps,
					&pExtraProps));

				if (pExtraProps)
				{
					ExtraPropForList.dwAlignPad = pExtraProps[0].dwAlignPad;

					if (PROP_TYPE(pNewTag.aulPropTag[0]) == NULL)
					{
						//In this case, we started with a NULL tag, but we got a property back - let's 'fix' our tag for the UI
						pNewTag.aulPropTag[0]  =
							CHANGE_PROP_TYPE(pNewTag.aulPropTag[0],PROP_TYPE(pExtraProps[0].ulPropTag));
					}

					//We want to give our parser the tag that came back from GetProps
					ExtraPropForList.ulPropTag = pExtraProps[0].ulPropTag;

					ExtraPropForList.Value = pExtraProps[0].Value;
				}
				else
				{
					ExtraPropForList.dwAlignPad = NULL;
					ExtraPropForList.ulPropTag = CHANGE_PROP_TYPE(pNewTag.aulPropTag[0],PT_ERROR);
					ExtraPropForList.Value.err = hRes;
				}

				//Add the property to the list
				AddPropToListBox(
					ulCurListBoxRow,
					pNewTag.aulPropTag[0],//Tag to use in the UI
					&ExtraPropForList);//Tag + Value to parse - may differ in case of errors or NULL type.

				ulCurListBoxRow++;

				MAPIFreeBuffer(pExtraProps);
				pExtraProps = NULL;
			}
		}
	}

	DebugPrintEx(DBGGeneric,CLASS,_T("LoadMAPIPropList"),_T("added %d properties\n"),ulCurListBoxRow);

	SortColumn(m_iClickedColumn);

	//Don't report any errors from here - don't care at this point
	return S_OK;
}//CSingleMAPIPropListCtrl::LoadMAPIPropList

HRESULT	CSingleMAPIPropListCtrl::RefreshMAPIPropList()
{
	HRESULT hRes = S_OK;
	int iSelectedItem;
	DebugPrintEx(DBGGeneric,CLASS,_T("RefreshMAPIPropList"),_T("\n"));

	//Turn off redraw while we work on the window
	MySetRedraw(FALSE);
	POSITION MyPos = GetFirstSelectedItemPosition();

	iSelectedItem = GetNextSelectedItem(MyPos);

	EC_B(DeleteAllItems());

	if (m_lpMAPIProp || GetPropVals())
		EC_H(LoadMAPIPropList());

	SetItemState(iSelectedItem,LVIS_SELECTED | LVIS_FOCUSED,LVIS_SELECTED | LVIS_FOCUSED);

	EnsureVisible(iSelectedItem,FALSE);

	//Turn redraw back on to update our view
	MySetRedraw(TRUE);

	if (m_lpHostDlg)
		m_lpHostDlg->UpdateStatusBarText(STATUSMIDDLEPANE,IDS_STATUSTEXTNUMPROPS,GetItemCount());

	return hRes;
}//CSingleMAPIPropListCtrl::RefreshMAPIPropList

HRESULT CSingleMAPIPropListCtrl::AddPropToExtraProps(ULONG ulPropTag,BOOL bRefresh)
{
	HRESULT			hRes = S_OK;
	SPropTagArray	sptSingleProp;

	DebugPrintEx(DBGGeneric,CLASS,_T("AddPropToExtraProps"),_T("adding proptag 0x%X\n"),ulPropTag);

	//Cache this proptag so we continue to request it in this view
	//We've got code to refresh any props cached in m_sptExtraProps...let's add to that.

	sptSingleProp.cValues = 1;
	sptSingleProp.aulPropTag[0] = ulPropTag;

	EC_H(AddPropsToExtraProps(&sptSingleProp,bRefresh));

	return hRes;
}//CSingleMAPIPropListCtrl::AddPropToExtraProps

HRESULT CSingleMAPIPropListCtrl::AddPropsToExtraProps(LPSPropTagArray lpPropsToAdd,BOOL bRefresh)
{
	HRESULT			hRes = S_OK;
	LPSPropTagArray lpNewExtraProps = NULL;

	DebugPrintEx(DBGGeneric,CLASS,_T("AddPropsToExtraProps"),_T("adding prop array 0x%X\n"),lpPropsToAdd);

	EC_H(ConcatSPropTagArrays(
		m_sptExtraProps,
		lpPropsToAdd,
		&lpNewExtraProps));

	MAPIFreeBuffer(m_sptExtraProps);
	m_sptExtraProps = lpNewExtraProps;

	if (bRefresh)
	{
		WC_H(RefreshMAPIPropList());
	}

	return hRes;
}

#define NUMPROPTYPES 31
static ULONG _PropTypeIcons[NUMPROPTYPES][2] =
{
{PT_UNSPECIFIED,slIconUNSPECIFIED},
{PT_NULL,slIconNULL},
{PT_I2,slIconI2},
{PT_LONG,slIconLONG},
{PT_R4,slIconR4},
{PT_DOUBLE,slIconDOUBLE},
{PT_CURRENCY,slIconCURRENCY},
{PT_APPTIME,slIconAPPTIME},
{PT_ERROR,slIconERROR},
{PT_BOOLEAN,slIconBOOLEAN},
{PT_OBJECT,slIconOBJECT},
{PT_I8,slIconI8},
{PT_STRING8,slIconSTRING8},
{PT_UNICODE,slIconUNICODE},
{PT_SYSTIME,slIconSYSTIME},
{PT_CLSID,slIconCLSID},
{PT_BINARY,slIconBINARY},
{PT_MV_I2,slIconMV_I2},
{PT_MV_LONG,slIconMV_LONG},
{PT_MV_R4,slIconMV_R4},
{PT_MV_DOUBLE,slIconMV_DOUBLE},
{PT_MV_CURRENCY,slIconMV_CURRENCY},
{PT_MV_APPTIME,slIconMV_APPTIME},
{PT_MV_SYSTIME,slIconMV_SYSTIME},
{PT_MV_STRING8,slIconMV_STRING8},
{PT_MV_BINARY,slIconMV_BINARY},
{PT_MV_UNICODE,slIconMV_UNICODE},
{PT_MV_CLSID,slIconMV_CLSID},
{PT_MV_I8,slIconMV_I8},
{PT_SRESTRICTION,slIconSRESTRICTION},
{PT_ACTIONS,slIconACTIONS},
};

//Crack open the given SPropValue and render it to the given row in the list.
void CSingleMAPIPropListCtrl::AddPropToListBox(
											   int iRow,
											   ULONG ulPropTag,
											   LPSPropValue lpsPropToAdd)
{
	HRESULT hRes = S_OK;
	CString PropTag;
	CString PropString;
	CString AltPropString;
	SortListData* lpData = NULL;

	//DebugPrintEx(DBGGeneric,CLASS,_T("AddPropToListBox"),_T("adding property:"));
	//DebugPrint(DBGGeneric,_T("0x%X, "), ulPropTag);
	ULONG ulImage = slIconDefault;
	if (lpsPropToAdd)
	{
		int i = 0;
		for (i=0;i<NUMPROPTYPES;i++)
		{
			if (_PropTypeIcons[i][0] == PROP_TYPE(lpsPropToAdd->ulPropTag))
			{
				ulImage = _PropTypeIcons[i][1];
				break;
			}
		}
	}
	lpData = InsertRow(iRow,_T(""),0,ulImage);

	//Data used to refer to specific property tags. See GetSelectedPropTag.
	if (lpData)
	{
		lpData->data.Prop.ulPropTag = ulPropTag;
		lpData->bItemFullyLoaded = true;
		lpData->ulSortDataType = SORTLIST_PROP;
	}

	LPTSTR szExactMatches = NULL;
	LPTSTR szPartialMatches = NULL;
	EC_H(PropTagToPropName(ulPropTag,m_bIsAB,&szExactMatches,&szPartialMatches));
	PropTag.Format(_T("0x%08X"),ulPropTag);// STRING_OK
	SetItemText(iRow,pcPROPEXACTNAMES,szExactMatches?szExactMatches:(LPCTSTR) PropTag);
	SetItemText(iRow,pcPROPPARTIALNAMES,szPartialMatches?szPartialMatches:_T(""));
	delete[] szPartialMatches;
	delete[] szExactMatches;

	//DebugPrint(DBGGeneric,_T("%s, "), (LPCTSTR) PropTag);
	SetItemText(iRow,pcPROPTAG,(LPCTSTR) PropTag);

	SetItemText(iRow,pcPROPTYPE,(LPCTSTR) TypeToString(ulPropTag));

	if (lpsPropToAdd)
	{
		InterpretProp(lpsPropToAdd,&PropString,&AltPropString);
		//DebugPrint(DBGGeneric,_T("%s, %s, %s\n"), (LPCTSTR) TypeToString(ulPropTag), (LPCTSTR) PropString, (LPCTSTR) AltPropString);
		SetItemText(iRow,pcPROPVAL,(LPCTSTR) PropString);
		SetItemText(iRow,pcPROPVALALT,(LPCTSTR) AltPropString);

		LPTSTR szFlags = NULL;
		EC_H(InterpretFlags(lpsPropToAdd, &szFlags));
		if (szFlags)
		{
			SetItemText(iRow,pcPROPFLAGS,(LPCTSTR) szFlags);
		}
		MAPIFreeBuffer(szFlags);
		szFlags = NULL;
	}

	//Named Props
	LPTSTR szNamedPropName = 0;
	LPTSTR szNamedPropGUID = 0;
	GetPropName(m_lpMAPIProp, ulPropTag, &szNamedPropName, &szNamedPropGUID);
	if (szNamedPropName)
	{
		SetItemText(iRow,pcPROPNAMEDNAME,szNamedPropName);
	}
	if (szNamedPropGUID)
	{
		SetItemText(iRow,pcPROPNAMEDIID,szNamedPropGUID);
	}
	delete[] szNamedPropName;
	delete[] szNamedPropGUID;
}//CSingleMAPIPropListCtrl::AddPropToListBox

//to get the count of source properties
ULONG CSingleMAPIPropListCtrl::GetCountPropVals()
{
	if (m_lpSourceData)
	{
		return m_lpSourceData->cSourceProps;
	}
	return 0;
}

BOOL CSingleMAPIPropListCtrl::IsModifiedPropVals()
{
	if (m_lpSourceData && m_bRowModified) return true;
	return false;
}

LPSPropValue CSingleMAPIPropListCtrl::GetPropVals()
{
	if (m_lpSourceData && m_lpSourceData->lpSourceProps)
	{
			return m_lpSourceData->lpSourceProps;
	}
	return 0;
}

//Clear the current property list from the control.
//Load a new list from the IMAPIProp or lpSourceProps object passed in
//Most calls to this will come through CBaseDialog::OnUpdateSingleMAPIPropListCtrl, which will preserve the current bIsAB
//Exceptions will be where we need to set a specific bIsAB
HRESULT CSingleMAPIPropListCtrl::SetDataSource(LPMAPIPROP lpMAPIProp, SortListData* lpListData, BOOL bIsAB)
{
	HRESULT hRes = S_OK;
	//if nothing to do...do nothing
	if (lpMAPIProp == m_lpMAPIProp && m_lpSourceData == lpListData) return S_OK;

	//Check that we've really got an LMAPIPROP
/*	if (lpMAPIProp)
	{
		LPMAPIPROP lpTemp = NULL;
		EC_H(lpMAPIProp->QueryInterface(IID_IMAPIProp,(LPVOID*) &lpTemp));
		if (!lpTemp) return MAPI_E_INVALID_PARAMETER;
		lpTemp->Release();
	}*/

	DebugPrintEx(DBGGeneric,CLASS,_T("SetDataSource"),_T("clearing 0x%X and adding 0x%X\n"),m_lpMAPIProp, lpMAPIProp);

	//Turn off redraw while we work on the window
	MySetRedraw(FALSE);

	m_bIsAB = bIsAB;
	m_lpSourceData = lpListData;

	if (m_lpMAPIProp) m_lpMAPIProp->Release();
	m_lpMAPIProp = lpMAPIProp;

	if (m_lpMAPIProp) m_lpMAPIProp->AddRef();

	WC_H(RefreshMAPIPropList());

	//Reset our header widths if weren't showing anything before and are now
	if (S_OK == hRes && !m_bHaveEverDisplayedSomething && m_lpMAPIProp && GetItemCount())
	{
		int iCurCol;
		m_bHaveEverDisplayedSomething = TRUE;
//		DebugPrintEx(DBGGeneric,CLASS,_T("SetDataSource"),_T("Setting header Column widths\n"));

		CHeaderCtrl* lpMyHeader = GetHeaderCtrl();

		if (lpMyHeader)
		{
			//This fixes a ton of flashing problems
			lpMyHeader->SetRedraw(TRUE);
			for (iCurCol = 0;iCurCol<NUMPROPCOLUMNS;iCurCol++)
			{
				SetColumnWidth(iCurCol,LVSCW_AUTOSIZE_USEHEADER);
				if (GetColumnWidth(iCurCol) > 200) SetColumnWidth(iCurCol,200);
			}
			lpMyHeader->SetRedraw(FALSE);
		}
	}

	//Turn redraw back on to update our view
	MySetRedraw(TRUE);
	return hRes;
}

void CSingleMAPIPropListCtrl::SavePropsToXML()
{
	TCHAR *szFileName = NULL;
	HRESULT hRes = S_OK;
	INT_PTR iDlgRet = 0;

	CString szFileSpec;
	szFileSpec.LoadString(IDS_XMLFILES);

	CFileDialogEx dlgFilePicker(
		FALSE,//Save As dialog
		_T("xml"),// STRING_OK
		_T("props.xml"),// STRING_OK
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT | OFN_EXPLORER,
		szFileSpec,
		this);

	EC_D_DIALOG(dlgFilePicker.DoModal());
	if (IDOK == iDlgRet)
	{
		szFileName = dlgFilePicker.m_ofn.lpstrFile;

		if (szFileName)
		{
			FILE* fProps = NULL;
			fProps = OpenFile(szFileName,true);
			if (fProps)
			{
				DebugPrintEx(DBGGeneric,CLASS,_T("SavePropsToXML"),_T("saving to %s\n"),szFileName);

				//force a sort on the tag column to make output consistent
				m_bSortUp = false;
				m_iClickedColumn = pcPROPTAG;
				SortColumn(m_iClickedColumn);

				int iRow;
				int iItemCount = GetItemCount();

				OutputToFile(fProps,_T("<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>\n"));
				OutputToFile(fProps,_T("<propertypane>\n"));
				for (iRow = 0; iRow<iItemCount; iRow++)
				{
					CString szTemp1;
					CString szTemp2;
					szTemp1 = GetItemText(iRow,pcPROPTAG);
					szTemp2 = GetItemText(iRow,pcPROPTYPE);
					OutputToFilef(fProps,_T("\t<property tag = \"%s\" type = \"%s\">\n"),szTemp1,szTemp2);

					szTemp1 = GetItemText(iRow,pcPROPEXACTNAMES);
					OutputXMLValueToFile(fProps,PropXMLNames[pcPROPEXACTNAMES].uidName,(LPCTSTR) szTemp1,2);

					szTemp1 = GetItemText(iRow,pcPROPPARTIALNAMES);
					OutputXMLValueToFile(fProps,PropXMLNames[pcPROPPARTIALNAMES].uidName,(LPCTSTR) szTemp1,2);

					szTemp1 = GetItemText(iRow,pcPROPNAMEDIID);
					OutputXMLValueToFile(fProps,PropXMLNames[pcPROPNAMEDIID].uidName,(LPCTSTR) szTemp1,2);

					szTemp1 = GetItemText(iRow,pcPROPNAMEDNAME);
					OutputXMLValueToFile(fProps,PropXMLNames[pcPROPNAMEDNAME].uidName,(LPCTSTR) szTemp1,2);

					SortListData* lpListData = ((SortListData*)GetItemData(iRow));
					ULONG ulPropType = PT_NULL;
					if (lpListData)
						ulPropType = PROP_TYPE(lpListData->data.Prop.ulPropTag);

					szTemp1 = GetItemText(iRow,pcPROPVAL);
					szTemp2 = GetItemText(iRow,pcPROPVALALT);
					switch (ulPropType)
					{
					case PT_STRING8:
					case PT_UNICODE:
						{
							OutputXMLCDataValueToFile(fProps,PropXMLNames[pcPROPVAL].uidName,(LPCTSTR) szTemp1,2);
							OutputXMLValueToFile(fProps,PropXMLNames[pcPROPVALALT].uidName,(LPCTSTR) szTemp2,2);
							break;
						}
					case PT_BINARY:
						{
							OutputXMLValueToFile(fProps,PropXMLNames[pcPROPVAL].uidName,(LPCTSTR) szTemp1,2);
							OutputXMLCDataValueToFile(fProps,PropXMLNames[pcPROPVALALT].uidName,(LPCTSTR) szTemp2,2);
							break;
						}
					default:
						{
							OutputXMLValueToFile(fProps,PropXMLNames[pcPROPVAL].uidName,(LPCTSTR) szTemp1,2);
							OutputXMLValueToFile(fProps,PropXMLNames[pcPROPVALALT].uidName,(LPCTSTR) szTemp2,2);
							break;
						}
					}

					szTemp1 = GetItemText(iRow,pcPROPFLAGS);
					OutputXMLValueToFile(fProps,PropXMLNames[pcPROPFLAGS].uidName,(LPCTSTR) szTemp1,2);

					OutputToFile(fProps,_T("\t</property>\n"));
				}
				OutputToFile(fProps,_T("</propertypane>"));
				CloseFile(fProps);
			}
		}
	}
}//CSingleMAPIPropListCtrl::SavePropsToXML

void CSingleMAPIPropListCtrl::OnDblclk(NMHDR* /*pNMHDR*/, LRESULT* pResult)
{
	DebugPrintEx(DBGGeneric,CLASS,_T("OnDblclk"),_T("calling OnEditProp\n"));
	OnEditProp();
	*pResult = 0;
}

void CSingleMAPIPropListCtrl::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	DebugPrintEx(DBGMenu,CLASS,_T("OnKeyDown"),_T("0x%X\n"),nChar);

	HRESULT hRes = S_OK;
	ULONG bCtrlPressed = GetKeyState(VK_CONTROL) <0;
	ULONG bShiftPressed = GetKeyState(VK_SHIFT) <0;
	ULONG bMenuPressed = GetKeyState(VK_MENU) <0;

	if (!bMenuPressed)
	{
		if ('X' == nChar && bCtrlPressed && !bShiftPressed)
		{
			OnDeleteProperty();
		}
		else if (VK_DELETE == nChar)
		{
			OnDeleteProperty();
		}
		else if ('S' == nChar && bCtrlPressed)
		{
			SavePropsToXML();
		}
		else if ('E' == nChar && bCtrlPressed)
		{
			OnEditProp();
		}
		else if ('C' == nChar && bCtrlPressed && !bShiftPressed)
		{
			OnCopyProperty();
		}
		else if ('V' == nChar && bCtrlPressed && !bShiftPressed)
		{
			OnPasteProperty();
		}
		else if (VK_F5 == nChar)
		{
			WC_H(RefreshMAPIPropList());
		}
		else if (VK_RETURN == nChar)
		{
			if (!bCtrlPressed)
			{
				DebugPrintEx(DBGMenu,CLASS,_T("OnKeyDown"),_T("calling OnEditProp\n"));
				OnEditProp();
			}
			else
			{
				DebugPrintEx(DBGMenu,CLASS,_T("OnKeyDown"),_T("calling OnOpenProperty\n"));
				OnOpenProperty();
			}
		}
		else if (!m_lpHostDlg || !m_lpHostDlg->HandleKeyDown(nChar,bShiftPressed,bCtrlPressed,bMenuPressed))
		{
			CSortListCtrl::OnKeyDown(nChar,nRepCnt,nFlags);
		}
	}
}

void CSingleMAPIPropListCtrl::OnContextMenu(CWnd* /*pWnd*/, CPoint pos)
{
	HRESULT hRes = S_OK;

	EC_B(DisplayContextMenu(IDR_MENU_PROPERTY_POPUP,NULL,this->m_lpHostDlg,pos.x, pos.y));
}

HRESULT CSingleMAPIPropListCtrl::FindAllNamedProps()
{
	HRESULT hRes = S_OK;
	LPMAPINAMEID FAR * lppPropNames = 0;
	LPSPropTagArray lptag = NULL;
	ULONG ulPropNames = 0;

	if (!m_lpMAPIProp) return hRes;

	//Exchange can return MAPI_E_NOT_ENOUGH_MEMORY when I call this - give it a try - PSTs support it
	DebugPrintEx(DBGNamedProp,CLASS,_T("FindAllNamedProps"),_T("Calling GetNamesFromIDs with a NULL\n"));
	WC_H( m_lpMAPIProp->GetNamesFromIDs(
		&lptag,
		NULL,
		NULL,
		&ulPropNames,
		&lppPropNames));
	if (S_OK == hRes)
	{
		//now we have a named prop NAME - need to get the ID.
		EC_H(m_lpMAPIProp->GetIDsFromNames(
			ulPropNames,
			lppPropNames,
			NULL,
			&lptag));
		if (lptag)
		{
			//Now we have an array of tags - add them in:
			EC_H(AddPropsToExtraProps(lptag,FALSE));
			MAPIFreeBuffer(lptag);
			lptag = NULL;
		}
	}
	else
	{
		hRes = S_OK;

		DebugPrintEx(DBGNamedProp,CLASS,_T("FindAllNamedProps"),_T("Exchange didn't support GetNamesFromIDs(NULL).\n"));

#define __LOWERBOUND 0x8000
#define __UPPERBOUNDDEFAULT 0x8FFF
#define __UPPERBOUND 0xFFFF

		CEditor MyData(
			this,
			IDS_FINDNAMEPROPSLIMIT,
			IDS_FINDNAMEPROPSLIMITPROMPT,
			2,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
		MyData.InitSingleLine(0,IDS_LOWERBOUND,NULL,false);
		MyData.SetHex(0,__LOWERBOUND);
		MyData.InitSingleLine(1,IDS_UPPERBOUND,NULL,false);
		MyData.SetHex(1,__UPPERBOUNDDEFAULT);

		WC_H(MyData.DisplayDialog());

		if (S_OK == hRes)
		{
			ULONG ulLowerBound = MyData.GetHex(0);
			ULONG ulUpperBound = MyData.GetHex(1);
			SPropTagArray tag = {0};

			DebugPrintEx(DBGNamedProp,CLASS,_T("FindAllNamedProps"),_T("Walking through all IDs from 0x%X to 0x%X, looking for mappings to names\n"),ulLowerBound,ulUpperBound);
			tag.cValues = 1;
			if (ulLowerBound < __LOWERBOUND)
			{
				ErrDialog(__FILE__,__LINE__,IDS_EDLOWERBOUNDTOOLOW,ulLowerBound,__LOWERBOUND);
				hRes = MAPI_E_INVALID_PARAMETER;
			}
			else if (ulUpperBound > __UPPERBOUND)
			{
				ErrDialog(__FILE__,__LINE__,IDS_EDUPPERBOUNDTOOHIGH,ulUpperBound,__UPPERBOUND);
				hRes = MAPI_E_INVALID_PARAMETER;
			}
			else if (ulLowerBound > ulUpperBound)
			{
				ErrDialog(__FILE__,__LINE__,IDS_EDLOWEROVERUPPER,ulLowerBound,ulUpperBound);
				hRes = MAPI_E_INVALID_PARAMETER;
			}
			else
			{
				lptag = &tag;
				ULONG iTag = 0;
				for (iTag = ulLowerBound ; iTag <= ulUpperBound ; iTag++)
				{
					hRes = S_OK;
					tag.aulPropTag[0] = PROP_TAG(NULL,iTag);

					WC_H(m_lpMAPIProp->GetNamesFromIDs(
						&lptag,
						NULL,
						NULL,
						&ulPropNames,
						&lppPropNames));
					if (S_OK == hRes && ulPropNames == 1 && lppPropNames && *lppPropNames)
					{
						DebugPrintEx(DBGNamedProp,CLASS,_T("FindAllNamedProps"),_T("Found an ID with a name (0x%X). Adding to extra prop list.\n"),iTag);
						EC_H(AddPropToExtraProps(PROP_TAG(NULL,iTag),FALSE));
					}
					MAPIFreeBuffer(lppPropNames);
					lppPropNames = NULL;
				}
			}
		}
	}

	//Refresh the display
	WC_H(RefreshMAPIPropList());

	MAPIFreeBuffer(lppPropNames);
	return hRes;
}

//Delete the selected property
void CSingleMAPIPropListCtrl::OnDeleteProperty()
{
	HRESULT		hRes = S_OK;
	ULONG		ulPropTag = NULL;

	if (!m_lpMAPIProp) return;

	GetSelectedPropTag(&ulPropTag);
	if (!ulPropTag) return;

	CEditor Query(
		this,
		IDS_DELETEPROPERTY,
		IDS_DELETEPROPERTYPROMPT,
		(ULONG) 0,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
	WC_H(Query.DisplayDialog());
	if (S_OK == hRes)
	{
		DebugPrintEx(DBGGeneric,CLASS,_T("OnDeleteProperty"),_T("deleting property 0x%X from 0x%X\n"),ulPropTag,m_lpMAPIProp);

		EC_H(DeleteProperty(m_lpMAPIProp,ulPropTag));

		//Refresh the display
		WC_H(RefreshMAPIPropList());
	}

	return;
}//CSingleMAPIPropListCtrl::OnDeleteProperty

//Display the selected property as a security dscriptor
void CSingleMAPIPropListCtrl::OnDisplayPropertyAsSecurityDescriptor()
{
	HRESULT					hRes = S_OK;
	ULONG					ulPropTag = NULL;
	if (!m_lpMAPIProp) return;

	GetSelectedPropTag(&ulPropTag);
	if (!ulPropTag) return;

	DebugPrintEx(DBGGeneric,CLASS,_T("OnDisplayPropertyAsSecurityDescriptor"),_T("interpreting 0x%X on 0x%X as Security Descriptor\n"), ulPropTag, m_lpMAPIProp);

	EC_H(DisplayPropAsSD(m_lpMAPIProp,ulPropTag));
}//CSingleMAPIPropListCtrl::OnDisplayPropertyAsSecurityDescriptor

//Display the selected property as a security dscriptor using a property sheet
void CSingleMAPIPropListCtrl::OnDisplayPropertyAsSecurityDescriptorPropSheet()
{
	HRESULT					hRes = S_OK;
	ULONG					ulPropTag = NULL;
	CMySecInfo*				MySecInfo = NULL;

	if (!m_lpMAPIProp || !pfnEditSecurity) return;

	GetSelectedPropTag(&ulPropTag);
	if (!ulPropTag) return;

	DebugPrintEx(DBGGeneric,CLASS,_T("OnDisplayPropertyAsSecurityDescriptorPropSheet"),_T("interpreting 0x%X on 0x%X as Security Descriptor\n"), ulPropTag, m_lpMAPIProp);

	MySecInfo = new CMySecInfo(m_lpMAPIProp,ulPropTag);

	if (MySecInfo)
	{
		EC_B(pfnEditSecurity(m_hWnd,MySecInfo));

		MySecInfo->Release();
	}
}//CSingleMAPIPropListCtrl::OnDisplayPropertyAsSecurityDescriptorPropSheet

void CSingleMAPIPropListCtrl::OnEditProp()
{
	ULONG			ulPropTag = NULL;

	if (!m_lpMAPIProp && !GetPropVals()) return;

	GetSelectedPropTag(&ulPropTag);
	if (!ulPropTag) return;

	OnEditGivenProp(ulPropTag);

	return;
}//CSingleMAPIPropListCtrl::OnEditProp

void CSingleMAPIPropListCtrl::OnEditPropAsRestriction(ULONG ulPropTag)
{
	HRESULT hRes = S_OK;

	if (!ulPropTag || PT_SRESTRICTION != PROP_TYPE(ulPropTag)) return;

	LPSPropValue lpEditProp = NULL;
	lpEditProp = PpropFindProp(
		GetPropVals(),
		GetCountPropVals(),
		ulPropTag);

	LPSRestriction lpResIn = NULL;
	if (lpEditProp)
	{
		lpResIn = (LPSRestriction) lpEditProp->Value.lpszA;
	}

	DebugPrint(DBGGeneric,_T("Source restriction before editing:\n"));
	DebugPrintRestriction(DBGGeneric,lpResIn,m_lpMAPIProp);
	CRestrictEditor MyResEditor(
		this,
		NULL,
		lpResIn);
	WC_H(MyResEditor.DisplayDialog());

	if (S_OK == hRes)
	{
		LPSRestriction lpModRes = MyResEditor.DetachModifiedSRestriction();
		//what to do with this now that I've got it?

		if (lpModRes)
		{
			DebugPrint(DBGGeneric,_T("Modified restriction:\n"));
			DebugPrintRestriction(DBGGeneric,lpModRes,m_lpMAPIProp);

			lpEditProp->Value.lpszA = (LPSTR) lpModRes;
			//don't free lpModRes since it'll get freed when the source array gets freed.

			m_bRowModified = true;

			//refresh
			RefreshMAPIPropList();
		}
	}
}//CSingleMAPIPropListCtrl::OnEditPropAsRestriction

// Concatenate two property arrays without duplicates
// Entries in the first array trump entries in the second
// Will also eliminate any duplicates already existing within the arrays
HRESULT	ConcatLPSPropValue(
							 ULONG ulVal1,
							 LPSPropValue lpVal1,
							 ULONG ulVal2,
							 LPSPropValue lpVal2,
							 ULONG* lpulRetVal,
							 LPSPropValue* lppRetVal)
{
	if (!lpulRetVal || !lppRetVal) return MAPI_E_INVALID_PARAMETER;
	if (ulVal1 && !lpVal1) return MAPI_E_INVALID_PARAMETER;
	if (ulVal2 && !lpVal2) return MAPI_E_INVALID_PARAMETER;
	*lpulRetVal = NULL;
	*lppRetVal = NULL;
	HRESULT hRes = S_OK;

	ULONG ulSourceArray = 0;
	ULONG ulTargetArray = 0;
	ULONG ulNewArraySize = NULL;
	LPSPropValue lpNewArray = NULL;

	//Add the sizes of the passed in arrays
	if (ulVal2 && ulVal1)
	{
		ulNewArraySize = ulVal1;
		// Only count props in the second array if they're not in the first
		for (ulSourceArray = 0; ulSourceArray < ulVal2; ulSourceArray++)
		{
			if (!PpropFindProp(lpVal1, ulVal1, CHANGE_PROP_TYPE(lpVal2[ulSourceArray].ulPropTag,PT_UNSPECIFIED)))
			{
				ulNewArraySize++;
			}
		}
	}
	else
	{
		ulNewArraySize = ulVal1 + ulVal2;
	}

	if (ulNewArraySize)
	{
		// Allocate the base array - PropCopyMore will allocmore as needed for string/bin/etc
		EC_H(MAPIAllocateBuffer(ulNewArraySize*sizeof(SPropValue),(LPVOID*) &lpNewArray));

		if (SUCCEEDED(hRes) && lpNewArray)
		{
			if (ulVal1)
			{
				for (ulSourceArray = 0;ulSourceArray<ulVal1;ulSourceArray++)
				{
					if (!ulTargetArray || // if it's NULL, we haven't added anything yet
						!PpropFindProp(
						lpNewArray,
						ulTargetArray,
						CHANGE_PROP_TYPE(lpVal1[ulSourceArray].ulPropTag,PT_UNSPECIFIED)))
					{
						EC_H(PropCopyMore(
							&lpNewArray[ulTargetArray],
							&lpVal1[ulSourceArray],
							MAPIAllocateMore,
							lpNewArray));
						if (SUCCEEDED(hRes))
						{
							ulTargetArray++;
						}
						else break;
					}
				}
			}

			if (SUCCEEDED(hRes) && ulVal2)
			{
				for (ulSourceArray = 0;ulSourceArray<ulVal2;ulSourceArray++)
				{
					if (!ulTargetArray || // if it's NULL, we haven't added anything yet
						!PpropFindProp(
						lpNewArray,
						ulTargetArray,
						CHANGE_PROP_TYPE(lpVal2[ulSourceArray].ulPropTag,PT_UNSPECIFIED)))
					{
						// make sure we don't overrun.
						if (ulTargetArray >= ulNewArraySize)
						{
							hRes = MAPI_E_CALL_FAILED;
							break;
						}
						EC_H(PropCopyMore(
							&lpNewArray[ulTargetArray],
							&lpVal2[ulSourceArray],
							MAPIAllocateMore,
							lpNewArray));
						if (SUCCEEDED(hRes))
						{
							ulTargetArray++;
						}
						else break;
					}
				}
			}

			if (FAILED(hRes))
			{
				MAPIFreeBuffer(lpNewArray);
			}
			else
			{
				*lpulRetVal = ulTargetArray;
				*lppRetVal = lpNewArray;
			}
		}
	}

	return hRes;
}

void CSingleMAPIPropListCtrl::OnEditGivenProp(ULONG ulPropTag)
{
	HRESULT			hRes = S_OK;
	if (!m_lpMAPIProp && !GetPropVals()) return;

	DebugPrintEx(DBGGeneric,CLASS,_T("OnEditGivenProp"),_T("editing property 0x%X (==%s) from 0x%X\n"),ulPropTag,(LPCTSTR) TagToString(ulPropTag,m_lpMAPIProp,m_bIsAB,true),m_lpMAPIProp);

	if (PT_SRESTRICTION == PROP_TYPE(ulPropTag))
	{
		OnEditPropAsRestriction(ulPropTag);
		return;
	}
	LPSPropValue lpSourceArray = GetPropVals();
	if (m_lpMAPIProp && !RegKeys[regkeyUSE_ROW_DATA_FOR_SINGLEPROPLIST].ulCurDWORD)
	{
		CPropertyEditor MyEditor(
			this,
			IDS_PROPEDITOR,
			IDS_PROPEDITORPROMPT,
			m_bIsAB,
			NULL);
		MyEditor.InitPropValue(m_lpMAPIProp,ulPropTag);
		WC_H(MyEditor.DisplayDialog());

		//Refresh the display
		WC_H(RefreshMAPIPropList());
	}
	else if (lpSourceArray)
	{
		ULONG ulSourceArray = GetCountPropVals();

		CPropertyEditor MyEditor(
			this,
			IDS_PROPEDITOR,
			IDS_PROPEDITORPROMPT,
			m_bIsAB,
			lpSourceArray);

		LPSPropValue lpEditProp = NULL;
		lpEditProp = PpropFindProp(
			lpSourceArray,
			ulSourceArray,
			ulPropTag);

		if (lpEditProp)
		{
			MyEditor.InitPropValue(lpEditProp);
		}
		else
		{
			MyEditor.InitPropValue(NULL,ulPropTag);
		}
		WC_H(MyEditor.DisplayDialog());

		if (S_OK == hRes)
		{
			LPSPropValue lpModProp = MyEditor.DetachModifiedSPropValue();// I 'own' the memory now.

			// Special case - if we're in 'row' view of a real object, then we didn't
			// call SetProps or SaveChanges just now, because no object was passed to
			// the editor. Since we expected SaveChanges to happen, do it here:
			if (m_lpMAPIProp && RegKeys[regkeyUSE_ROW_DATA_FOR_SINGLEPROPLIST].ulCurDWORD)
			{
				LPSPropProblemArray lpProblemArray = NULL;

				EC_H(m_lpMAPIProp->SetProps(
					1,
					lpModProp,
					&lpProblemArray));

				EC_PROBLEMARRAY(lpProblemArray);
				MAPIFreeBuffer(lpProblemArray);

				EC_H(m_lpMAPIProp->SaveChanges(KEEP_OPEN_READWRITE));
			}

			if (lpModProp)
			{
				// Merge the new prop with the old array
				// Pass new prop first so it's value will replace old value if it existed
				ULONG ulNewArray = NULL;
				LPSPropValue lpNewArray = NULL;
				EC_H(ConcatLPSPropValue(
					1,
					lpModProp,
					m_lpSourceData->cSourceProps,
					m_lpSourceData->lpSourceProps,
					&ulNewArray,
					&lpNewArray));

				if (SUCCEEDED(hRes))
				{
					MAPIFreeBuffer(m_lpSourceData->lpSourceProps);
					m_lpSourceData->cSourceProps = ulNewArray;
					m_lpSourceData->lpSourceProps = lpNewArray;
					m_bRowModified = true;

					//refresh
					RefreshMAPIPropList();
				}
				MAPIFreeBuffer(lpModProp);
			}
		}
	}
}//OnEditGivenProp

//Display the selected property as a stream using CStreamEditor
void CSingleMAPIPropListCtrl::OnEditPropAsStream(ULONG ulType, BOOL bEditAsRTF)
{
	HRESULT			hRes = S_OK;
	ULONG			ulPropTag = NULL;

	if (!m_lpMAPIProp) return;

	GetSelectedPropTag(&ulPropTag);
	if (!ulPropTag) return;

	DebugPrintEx(DBGGeneric,CLASS,_T("OnEditPropAsStream"),_T("editing property 0x%X (== %s) from 0x%X as stream, ulType = 0x%08X, bEditAsRTF = 0x%X\n"),
		ulPropTag,
		(LPCTSTR) TagToString(ulPropTag,m_lpMAPIProp,m_bIsAB,true),
		m_lpMAPIProp,
		ulType,
		bEditAsRTF);

	ulPropTag = CHANGE_PROP_TYPE(ulPropTag,ulType);

	BOOL bUseWrapEx = false;
	ULONG ulRTFFlags = NULL;
	ULONG ulInCodePage = NULL;
	ULONG ulOutCodePage = CP_ACP;//Default to ANSI - check if this is valid for UNICODE builds

	if (bEditAsRTF)
	{
		if (pfnWrapEx)
		{
			CEditor MyPrompt(
				this,
				IDS_USEWRAPEX,
				IDS_USEWRAPEXPROMPT,
				1,
				CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
			MyPrompt.InitCheck(0,IDS_USEWRAPEX,true,false);

			WC_H(MyPrompt.DisplayDialog());
			if (S_OK == hRes)
			{
				if (MyPrompt.GetCheck(0))
				{
					bUseWrapEx = true;
					SPropTagArray pTag = {0};
					pTag.cValues = 1;
					pTag.aulPropTag[0] = PR_INTERNET_CPID;
					ULONG ulPropVal = NULL;
					LPSPropValue lpProp = NULL;

					WC_H(m_lpMAPIProp->GetProps(
						&pTag,
						fMapiUnicode,
						&ulPropVal,
						&lpProp));
					if (lpProp && 1 == ulPropVal && PT_LONG == PROP_TYPE(lpProp[0].ulPropTag))
					{
						ulInCodePage = lpProp[0].Value.l;
					}

					MAPIFreeBuffer(lpProp);

					CEditor MyPrompt2(
						this,
						IDS_WRAPEXFLAGS,
						IDS_WRAPEXFLAGSPROMPT,
						3,
						CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
					MyPrompt2.InitSingleLine(0,IDS_WRAPEXFLAGS,NULL,false);
					MyPrompt2.SetHex(0,MAPI_NATIVE_BODY);
					MyPrompt2.InitSingleLine(1,IDS_ULINCODEPAGE,NULL,false);
					MyPrompt2.SetDecimal(1,ulInCodePage);
					MyPrompt2.InitSingleLine(2,IDS_ULCODEPAGE,NULL,false);
					MyPrompt2.SetDecimal(2,0);

					WC_H(MyPrompt2.DisplayDialog());

					if (S_OK == hRes)
					{
						ulRTFFlags = MyPrompt2.GetHex(0);
						ulInCodePage = MyPrompt2.GetDecimal(1);
						ulOutCodePage = MyPrompt2.GetDecimal(2);
					}
				}
			}
		}
	}

	if (S_OK == hRes)
	{
		CStreamEditor MyEditor(
			this,
			IDS_PROPEDITOR,
			IDS_STREAMEDITORPROMPT,
			m_lpMAPIProp,
			ulPropTag,
			m_bIsAB,
			bEditAsRTF,
			bUseWrapEx,
			ulRTFFlags,
			ulInCodePage,
			ulOutCodePage);

		WC_H(MyEditor.DisplayDialog());
	}
}//CSingleMAPIPropListCtrl::OnEditPropAsStream

void CSingleMAPIPropListCtrl::OnCopyProperty()
{
	//for now, we only copy from objects - copying from rows would be difficult to generalize
	if (!m_lpHostDlg || !m_lpHostDlg->m_lpMapiObjects || !m_lpMAPIProp) return;

	ULONG ulPropTag = NULL;
	GetSelectedPropTag(&ulPropTag);

	m_lpHostDlg->m_lpMapiObjects->SetPropertyToCopy(ulPropTag,m_lpMAPIProp);
}

// Allocates return value with new.
// clean up with delete.
TZREG* BinToTZREG(ULONG cbReg, LPBYTE lpbReg)
{
	if (!lpbReg) return NULL;

	// Update this if parsing code is changed!
	if (cbReg < 3*sizeof(long) + 2*sizeof(WORD) + 2*sizeof(SYSTEMTIME)) return NULL;

	TZREG tzReg = {0};
	LPBYTE lpPtr = lpbReg;

	tzReg.lBias = *((long*)lpPtr);
	lpPtr += sizeof(long);
	tzReg.lStandardBias = *((long*)lpPtr);
	lpPtr += sizeof(long);
	tzReg.lDaylightBias = *((long*)lpPtr);
	lpPtr += sizeof(long);
	lpPtr += sizeof(WORD);// reserved

	tzReg.stStandardDate = *((SYSTEMTIME*)lpPtr);
	lpPtr += sizeof(SYSTEMTIME);
	lpPtr += sizeof(WORD);// reserved
	tzReg.stDaylightDate = *((SYSTEMTIME*)lpPtr);
	lpPtr += sizeof(SYSTEMTIME);

	TZREG* ptzReg = NULL;
	ptzReg = new TZREG;
	if (ptzReg)
	{
		*ptzReg = tzReg;
	}

	return ptzReg;
}

// result allocated with new
// clean up with delete[]
LPTSTR TZREGToString(TZREG tzReg)
{
	HRESULT hRes = S_OK;
	CString szReg;
	LPTSTR szOut = NULL;

	szReg.FormatMessage(IDS_TZREGTOSTRING,
		tzReg.lBias,
		tzReg.lStandardBias,
		tzReg.lDaylightBias,
		tzReg.stStandardDate.wYear,
		tzReg.stStandardDate.wMonth,
		tzReg.stStandardDate.wDayOfWeek,
		tzReg.stStandardDate.wDay,
		tzReg.stStandardDate.wHour,
		tzReg.stStandardDate.wMinute,
		tzReg.stStandardDate.wSecond,
		tzReg.stStandardDate.wMilliseconds,
		tzReg.stDaylightDate.wYear,
		tzReg.stDaylightDate.wMonth,
		tzReg.stDaylightDate.wDayOfWeek,
		tzReg.stDaylightDate.wDay,
		tzReg.stDaylightDate.wHour,
		tzReg.stDaylightDate.wMinute,
		tzReg.stDaylightDate.wSecond,
		tzReg.stDaylightDate.wMilliseconds);

	size_t cchReg = szReg.GetLength()+1;
	szOut = new TCHAR[cchReg];
	if (szOut)
	{
		EC_H(StringCchCopy(szOut,cchReg,(LPCTSTR)szReg));
	}
	return szOut;
}

void CSingleMAPIPropListCtrl::OnDisplayPropAsTZREG()
{
	HRESULT			hRes = S_OK;
	ULONG			ulPropTag = NULL;
	LPTSTR			szReg = NULL;
	LPSPropValue	lpsPropFromProp = NULL;
	LPSPropValue	lpsPropToDisplay = NULL;

	if (!m_lpMAPIProp && !GetPropVals()) return;

	GetSelectedPropTag(&ulPropTag);
	if (!ulPropTag || PROP_TYPE(ulPropTag) != PT_BINARY) return;
	DebugPrintEx(DBGGeneric,CLASS,_T("OnDisplayPropAsTZREG"),_T("interpreting 0x%X on 0x%X as TZREG\n"), ulPropTag, m_lpMAPIProp);

	if (m_lpMAPIProp && !RegKeys[regkeyUSE_ROW_DATA_FOR_SINGLEPROPLIST].ulCurDWORD)
	{
		WC_H(HrGetOneProp(m_lpMAPIProp,ulPropTag,&lpsPropFromProp));
		lpsPropToDisplay = lpsPropFromProp;
	}
	else if (GetPropVals())
	{
		lpsPropToDisplay = PpropFindProp(GetPropVals(),GetCountPropVals(),ulPropTag);
	}

	if (lpsPropToDisplay && PT_BINARY == PROP_TYPE(lpsPropToDisplay->ulPropTag))
	{
		TZREG* ptzReg = BinToTZREG(lpsPropToDisplay->Value.bin.cb,lpsPropToDisplay->Value.bin.lpb);
		if (ptzReg)
		{
			szReg = TZREGToString(*ptzReg);
			CEditor MyTZREG(
				this,
				IDS_TZREG,
				IDS_TZREGPROMPT,
				2,
				CEDITOR_BUTTON_OK);
			MyTZREG.InitSingleLine(0,IDS_PROPTAG,NULL,true);
			MyTZREG.SetHex(0,ulPropTag);
			MyTZREG.InitMultiLine(1,IDS_TZREGSTRUCT,NULL,true);
			MyTZREG.SetString(1,szReg);

			WC_H(MyTZREG.DisplayDialog());

			if (szReg) delete[] szReg;
			delete ptzReg;
		}
		else
		{
			CEditor MyTZREGErr(
				this,
				IDS_TZREG,
				IDS_TZERR,
				0,
				CEDITOR_BUTTON_OK);

			WC_H(MyTZREGErr.DisplayDialog());
		}
	}
	MAPIFreeBuffer(lpsPropFromProp);
}

// Allocates return value with new.
// clean up with delete[].
TZDEFINITION* BinToTZDEFINITION(ULONG cbDef, LPBYTE lpbDef)
{
	if (!lpbDef) return NULL;

	// Update this if parsing code is changed!
	// this checks the size up to the flags member
	if (cbDef < 2*sizeof(BYTE) + 2*sizeof(WORD)) return NULL;

	TZDEFINITION tzDef = {0};
	TZRULE* lpRules = NULL;
	LPBYTE lpPtr = lpbDef;
	WORD cchKeyName = NULL;
	WCHAR* szKeyName = NULL;
	WORD i = 0;

	BYTE bMajorVersion = *((BYTE*)lpPtr);
	lpPtr += sizeof(BYTE);
	BYTE bMinorVersion = *((BYTE*)lpPtr);
	lpPtr += sizeof(BYTE);

	// We only understand TZ_BIN_VERSION_MAJOR
	if (TZ_BIN_VERSION_MAJOR != bMajorVersion) return NULL;

	// We only understand if >= TZ_BIN_VERSION_MINOR
	if (TZ_BIN_VERSION_MINOR > bMinorVersion) return NULL;

	lpPtr += sizeof(WORD);

	tzDef.wFlags = *((WORD*)lpPtr);
	lpPtr += sizeof(WORD);

	if (TZDEFINITION_FLAG_VALID_GUID & tzDef.wFlags)
	{
		if (lpbDef + cbDef - lpPtr < sizeof(GUID)) return NULL;
		tzDef.guidTZID = *((GUID*)lpPtr);
		lpPtr += sizeof(GUID);
	}

	if (TZDEFINITION_FLAG_VALID_KEYNAME & tzDef.wFlags)
	{
		if (lpbDef + cbDef - lpPtr < sizeof(WORD)) return NULL;
		cchKeyName = *((WORD*)lpPtr);
		lpPtr += sizeof(WORD);
		if (cchKeyName)
		{
			if (lpbDef + cbDef - lpPtr < (BYTE)sizeof(WORD)*cchKeyName) return NULL;
			szKeyName = (WCHAR*)lpPtr;
			lpPtr += cchKeyName*sizeof(WORD);
		}
	}

	if (lpbDef+ cbDef - lpPtr < sizeof(WORD)) return NULL;
	tzDef.cRules = *((WORD*)lpPtr);
	lpPtr += sizeof(WORD);

	if (tzDef.cRules)
	{
		lpRules = new TZRULE[tzDef.cRules];
		if (!lpRules) return NULL;

		LPBYTE lpNextRule = lpPtr;
		BOOL bRuleOK = false;

		for (i = 0;i<tzDef.cRules;i++)
		{
			bRuleOK = false;
			lpPtr = lpNextRule;

			if (lpbDef + cbDef - lpPtr < 2*sizeof(BYTE) + 2*sizeof(WORD) + 3*sizeof(long) + 2*sizeof(SYSTEMTIME)) return NULL;
			bRuleOK = true;
			BYTE bRuleMajorVersion = *((BYTE*)lpPtr);
			lpPtr += sizeof(BYTE);
			BYTE bRuleMinorVersion = *((BYTE*)lpPtr);
			lpPtr += sizeof(BYTE);

			// We only understand TZ_BIN_VERSION_MAJOR
			if (TZ_BIN_VERSION_MAJOR != bRuleMajorVersion) return NULL;

			// We only understand if >= TZ_BIN_VERSION_MINOR
			if (TZ_BIN_VERSION_MINOR > bRuleMinorVersion) return NULL;

			WORD cbRule = *((WORD*)lpPtr);
			lpPtr += sizeof(WORD);

			lpNextRule = lpPtr + cbRule;

			lpRules[i].wFlags = *((WORD*)lpPtr);
			lpPtr += sizeof(WORD);

			lpRules[i].stStart = *((SYSTEMTIME*)lpPtr);
			lpPtr += sizeof(SYSTEMTIME);

			lpRules[i].TZReg.lBias = *((long*)lpPtr);
			lpPtr += sizeof(long);
			lpRules[i].TZReg.lStandardBias = *((long*)lpPtr);
			lpPtr += sizeof(long);
			lpRules[i].TZReg.lDaylightBias = *((long*)lpPtr);
			lpPtr += sizeof(long);

			lpRules[i].TZReg.stStandardDate = *((SYSTEMTIME*)lpPtr);
			lpPtr += sizeof(SYSTEMTIME);
			lpRules[i].TZReg.stDaylightDate = *((SYSTEMTIME*)lpPtr);
			lpPtr += sizeof(SYSTEMTIME);
		}
		if (!bRuleOK)
		{
			delete[] lpRules;
			return NULL;
		}
	}

	// Now we've read everything - allocate a structure and copy it in

	size_t cbTZDef = sizeof(TZDEFINITION) +
		sizeof(WCHAR)*(cchKeyName+1) +
		sizeof(TZRULE)*tzDef.cRules;

	TZDEFINITION* ptzDef = (TZDEFINITION*) new BYTE[cbTZDef];

	if (ptzDef)
	{
		// Copy main struct over
		*ptzDef = tzDef;
		lpPtr = (LPBYTE) ptzDef;
		lpPtr += sizeof(TZDEFINITION);

		if (szKeyName)
		{
			ptzDef->pwszKeyName = (WCHAR*)lpPtr;
			memcpy(lpPtr,szKeyName,cchKeyName*sizeof(WCHAR));
			ptzDef->pwszKeyName[cchKeyName] = 0;

			lpPtr += (cchKeyName+1)*sizeof(WCHAR);
		}

		if (ptzDef->cRules && lpRules)
		{
			ptzDef->rgRules = (TZRULE*)lpPtr;
			for (i = 0;i<ptzDef->cRules;i++)
			{
				ptzDef->rgRules[i] = lpRules[i];
			}
		}
	}

	delete[] lpRules;

	return ptzDef;
}

// result allocated with new
// clean up with delete[]
LPTSTR TZDEFINITIONToString(TZDEFINITION tzDef)
{
	CString szDef;
	LPTSTR szOut = NULL;
	LPTSTR szGUID = GUIDToString(&tzDef.guidTZID);

	szDef.FormatMessage(IDS_TZDEFTOSTRING,
		tzDef.wFlags,
		szGUID?szGUID:_T("null"),// STRING_OK
		tzDef.pwszKeyName,
		tzDef.cRules);

	if (tzDef.cRules && tzDef.rgRules)
	{
		CString szRule;

		if (szRule)
		{
			WORD i = 0;

			for (i = 0;i<tzDef.cRules;i++)
			{
				szRule.FormatMessage(IDS_TZRULETOSTRING,
					i,
					tzDef.rgRules[i].wFlags,
					tzDef.rgRules[i].stStart.wYear,
					tzDef.rgRules[i].stStart.wMonth,
					tzDef.rgRules[i].stStart.wDayOfWeek,
					tzDef.rgRules[i].stStart.wDay,
					tzDef.rgRules[i].stStart.wHour,
					tzDef.rgRules[i].stStart.wMinute,
					tzDef.rgRules[i].stStart.wSecond,
					tzDef.rgRules[i].stStart.wMilliseconds,
					tzDef.rgRules[i].TZReg.lBias,
					tzDef.rgRules[i].TZReg.lStandardBias,
					tzDef.rgRules[i].TZReg.lDaylightBias,
					tzDef.rgRules[i].TZReg.stStandardDate.wYear,
					tzDef.rgRules[i].TZReg.stStandardDate.wMonth,
					tzDef.rgRules[i].TZReg.stStandardDate.wDayOfWeek,
					tzDef.rgRules[i].TZReg.stStandardDate.wDay,
					tzDef.rgRules[i].TZReg.stStandardDate.wHour,
					tzDef.rgRules[i].TZReg.stStandardDate.wMinute,
					tzDef.rgRules[i].TZReg.stStandardDate.wSecond,
					tzDef.rgRules[i].TZReg.stStandardDate.wMilliseconds,
					tzDef.rgRules[i].TZReg.stDaylightDate.wYear,
					tzDef.rgRules[i].TZReg.stDaylightDate.wMonth,
					tzDef.rgRules[i].TZReg.stDaylightDate.wDayOfWeek,
					tzDef.rgRules[i].TZReg.stDaylightDate.wDay,
					tzDef.rgRules[i].TZReg.stDaylightDate.wHour,
					tzDef.rgRules[i].TZReg.stDaylightDate.wMinute,
					tzDef.rgRules[i].TZReg.stDaylightDate.wSecond,
					tzDef.rgRules[i].TZReg.stDaylightDate.wMilliseconds);
				szDef += szRule;
			}
		}
	}

	size_t cchDef = szDef.GetLength()+1;
	szOut = new TCHAR[cchDef];
	if (szOut)
	{
		HRESULT hRes = S_OK;
		EC_H(StringCchCopy(szOut,cchDef,(LPCTSTR)szDef));
	}

	delete[] szGUID;
	return szOut;
}

void CSingleMAPIPropListCtrl::OnDisplayPropAsTZDEFINITION()
{
	HRESULT			hRes = S_OK;
	ULONG			ulPropTag = NULL;
	LPTSTR			szDefinition = NULL;
	LPSPropValue	lpsPropFromProp = NULL;
	LPSPropValue	lpsPropToDisplay = NULL;

	if (!m_lpMAPIProp && !GetPropVals()) return;

	GetSelectedPropTag(&ulPropTag);
	if (!ulPropTag || PROP_TYPE(ulPropTag) != PT_BINARY) return;
	DebugPrintEx(DBGGeneric,CLASS,_T("OnDisplayPropAsTZDEFINITION"),_T("interpreting 0x%X on 0x%X as TZREG\n"), ulPropTag, m_lpMAPIProp);

	if (m_lpMAPIProp && !RegKeys[regkeyUSE_ROW_DATA_FOR_SINGLEPROPLIST].ulCurDWORD)
	{
		WC_H(HrGetOneProp(m_lpMAPIProp,ulPropTag,&lpsPropFromProp));
		lpsPropToDisplay = lpsPropFromProp;
	}
	else if (GetPropVals())
	{
		lpsPropToDisplay = PpropFindProp(GetPropVals(),GetCountPropVals(),ulPropTag);
	}

	if (lpsPropToDisplay && PT_BINARY == PROP_TYPE(lpsPropToDisplay->ulPropTag))
	{
		TZDEFINITION* ptzDefinition = BinToTZDEFINITION(lpsPropToDisplay->Value.bin.cb,lpsPropToDisplay->Value.bin.lpb);
		if (ptzDefinition)
		{
			szDefinition = TZDEFINITIONToString(*ptzDefinition);
			CEditor MyTZDEFINITION(
				this,
				IDS_TZDEFINITION,
				IDS_TZDEFINITIONPROMPT,
				2,
				CEDITOR_BUTTON_OK);
			MyTZDEFINITION.InitSingleLine(0,IDS_PROPTAG,NULL,true);
			MyTZDEFINITION.SetHex(0,ulPropTag);
			MyTZDEFINITION.InitMultiLine(1,IDS_TZDEFINITIONSTRUCT,NULL,true);
			MyTZDEFINITION.SetString(1,szDefinition);

			WC_H(MyTZDEFINITION.DisplayDialog());

			if (szDefinition) delete[] szDefinition;
			delete[] ptzDefinition;
		}
		else
		{
			CEditor MyTZREGErr(
				this,
				IDS_TZREG,
				IDS_TZERR,
				0,
				CEDITOR_BUTTON_OK);

			WC_H(MyTZREGErr.DisplayDialog());
		}
	}
	MAPIFreeBuffer(lpsPropFromProp);
}

void CSingleMAPIPropListCtrl::OnPasteProperty()
{
	//for now, we only paste to objects - copying to rows would be difficult to generalize
	if (!m_lpHostDlg || !m_lpHostDlg->m_lpMapiObjects || !m_lpMAPIProp) return;

	ULONG ulSourcePropTag = m_lpHostDlg->m_lpMapiObjects->GetPropertyToCopy();
	LPMAPIPROP lpSourcePropObj = m_lpHostDlg->m_lpMapiObjects->GetSourcePropObject();
	if (!lpSourcePropObj) return;

	HRESULT hRes = S_OK;
	LPSPropProblemArray lpProblems = NULL;
	SPropTagArray TagArray = {0};
	TagArray.cValues = 1;
	TagArray.aulPropTag[0] = ulSourcePropTag;

	CEditor MyData(
		this,
		IDS_PASTEPROP,
		IDS_PASTEPROPPROMPT,
		3,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	UINT uidDropDown[3] = {
		IDS_DDCOPYPROPS,
			IDS_DDGETSETPROPS,
			IDS_DDCOPYSTREAM
	};
	MyData.InitDropDown(0,IDS_COPYSTYLE,3,uidDropDown,true);
	MyData.InitSingleLine(1,IDS_SOURCEPROP,NULL,false);
	MyData.SetHex(1,ulSourcePropTag);
	MyData.InitSingleLine(2,IDS_TARGETPROP,NULL,false);
	MyData.SetHex(2,ulSourcePropTag);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		ULONG ulSourceTag = MyData.GetHex(1);
		ULONG ulTargetTag = MyData.GetHex(2);
		TagArray.aulPropTag[0] = ulSourceTag;

		if (PROP_TYPE(ulTargetTag) != PROP_TYPE(ulSourceTag))
			ulTargetTag = CHANGE_PROP_TYPE(ulTargetTag,PROP_TYPE(ulSourceTag));

		switch (MyData.GetDropDown(0))
		{
		case 0:
			{
				CEditor MyCopyData(
					this,
					IDS_PASTEPROP,
					IDS_COPYPASTEPROMPT,
					2,
					CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

				LPTSTR szGuid = GUIDToStringAndName(&IID_IMAPIProp);
				MyCopyData.InitSingleLineSz(0,IDS_INTERFACE,szGuid,false);
				delete[] szGuid;
				szGuid = NULL;
				MyCopyData.InitSingleLine(1,IDS_FLAGS,NULL,false);
				MyCopyData.SetHex(1,MAPI_DIALOG);

				WC_H(MyCopyData.DisplayDialog());
				if (S_OK == hRes)
				{
					GUID	MyGUID = {0};
					CString szTemp = MyCopyData.GetString(0);
					EC_H(StringToGUID((LPCTSTR) szTemp,&MyGUID));

					LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IMAPIProp::CopyProps"), m_lpHostDlg->m_hWnd);// STRING_OK

					ULONG ulCopyFlags = MyCopyData.GetHex(1);

					if(lpProgress)
						ulCopyFlags |= MAPI_DIALOG;

					EC_H(lpSourcePropObj->CopyProps(
						&TagArray,
						lpProgress ? (ULONG_PTR)m_lpHostDlg->m_hWnd : NULL,//ui param
						lpProgress,//progress
						&MyGUID,
						m_lpMAPIProp,
						ulCopyFlags,
						&lpProblems));

					if(lpProgress)
						lpProgress->Release();

					lpProgress = NULL;
				}
			}
			break;
		case 1:
			{
				ULONG			ulValues = NULL;
				LPSPropValue	lpSourceProp = NULL;
				EC_H(lpSourcePropObj->GetProps(
					&TagArray,
					fMapiUnicode,
					&ulValues,
					&lpSourceProp));
				if (!FAILED(hRes) && ulValues && lpSourceProp && PT_ERROR != lpSourceProp->ulPropTag)
				{
					lpSourceProp->ulPropTag = ulTargetTag;
					EC_H(m_lpMAPIProp->SetProps(
						ulValues,
						lpSourceProp,
						&lpProblems));
				}
			}
			break;
		case 2:
			EC_H(CopyPropertyAsStream(lpSourcePropObj,m_lpMAPIProp,ulSourceTag,ulTargetTag));
			break;
		}
	}

	EC_PROBLEMARRAY(lpProblems);
	MAPIFreeBuffer(lpProblems);

	if (!FAILED(hRes))
	{
		EC_H(m_lpMAPIProp->SaveChanges(KEEP_OPEN_READWRITE));

		//refresh
		RefreshMAPIPropList();
	}

	lpSourcePropObj->Release();
}

void CSingleMAPIPropListCtrl::OnCopyTo()
{
	//for now, we only copy from objects - copying from rows would be difficult to generalize
	if (!m_lpHostDlg || !m_lpHostDlg->m_lpMapiObjects || !m_lpMAPIProp) return;

	LPMAPIPROP lpSourcePropObj = m_lpHostDlg->m_lpMapiObjects->GetSourcePropObject();
	if (!lpSourcePropObj) return;

	HRESULT hRes = S_OK;
	LPSPropProblemArray lpProblems = NULL;

	CEditor MyData(
		this,
		IDS_COPYTO,
		IDS_COPYPASTEPROMPT,
		2,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	LPTSTR szGuid = GUIDToStringAndName(&IID_IMAPIProp);
	MyData.InitSingleLineSz(0,IDS_INTERFACE,szGuid,false);
	delete[] szGuid;
	szGuid = NULL;
	MyData.InitSingleLine(1,IDS_FLAGS,NULL,false);
	MyData.SetHex(1,MAPI_DIALOG);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		GUID	MyGUID = {0};
		CString szTemp = MyData.GetString(0);
		EC_H(StringToGUID((LPCTSTR) szTemp,&MyGUID));

		CTagArrayEditor TagEditor(
			this,
			IDS_TAGSTOEXCLUDE,
			IDS_TAGSTOEXCLUDEPROMPT,
			NULL,
			false,
			m_lpMAPIProp);
		WC_H(TagEditor.DisplayDialog());

		if (S_OK == hRes)
		{
			LPSPropTagArray lpTagArray = TagEditor.DetachModifiedTagArray();

			LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IMAPIProp::CopyTo"), m_lpHostDlg->m_hWnd);// STRING_OK

			ULONG ulCopyFlags = MyData.GetHex(1);

			if(lpProgress)
				ulCopyFlags |= MAPI_DIALOG;

			EC_H(lpSourcePropObj->CopyTo(
				0,
				NULL, //TODO: exclude interfaces?
				lpTagArray,
				lpProgress ? (ULONG_PTR)m_lpHostDlg->m_hWnd : NULL, //UI param
				lpProgress, //progress
				&MyGUID,
				m_lpMAPIProp,
				ulCopyFlags,//flags
				&lpProblems));
			MAPIFreeBuffer(lpTagArray);

			if(lpProgress)
				lpProgress->Release();

			lpProgress = NULL;
		}
	}
	if (lpProblems)
	{
		EC_PROBLEMARRAY(lpProblems);
		MAPIFreeBuffer(lpProblems);
	}

	if (!FAILED(hRes))
	{
		EC_H(m_lpMAPIProp->SaveChanges(KEEP_OPEN_READWRITE));

		//refresh
		RefreshMAPIPropList();
	}

	lpSourcePropObj->Release();
}

void CSingleMAPIPropListCtrl::OnOpenProperty()
{
	HRESULT			hRes = S_OK;
	ULONG			ulPropTag = NULL;

	if (!m_lpHostDlg) return;

	GetSelectedPropTag(&ulPropTag);
	if (!ulPropTag) return;

	DebugPrintEx(DBGGeneric,CLASS,_T("OnOpenProperty"),_T("asked to open 0x%X on 0x%X\n"), ulPropTag, m_lpMAPIProp);
	LPSPropValue lpProp = NULL;
	if (m_lpMAPIProp)
	{
		EC_H(HrGetOneProp(m_lpMAPIProp,ulPropTag,&lpProp));
	}
	else if (GetPropVals())
	{
		lpProp = PpropFindProp(
			GetPropVals(),
			GetCountPropVals(),
			ulPropTag);
	}

	if (!FAILED(hRes) && lpProp)
	{
		if (m_lpMAPIProp && PT_OBJECT == PROP_TYPE(lpProp->ulPropTag))
		{
			EC_H(DisplayTable(
				m_lpMAPIProp,
				lpProp->ulPropTag,
				otDefault,
				m_lpHostDlg));
		}
		else if (PT_BINARY == PROP_TYPE(lpProp->ulPropTag) || PT_MV_BINARY == PROP_TYPE(lpProp->ulPropTag) )
		{
			switch (PROP_TYPE(lpProp->ulPropTag))
			{
			case (PT_BINARY):
				{
					DebugPrintEx(DBGGeneric,CLASS,_T("OnOpenProperty"),_T("property is PT_BINARY\n"));
					m_lpHostDlg->OnOpenEntryID(&lpProp->Value.bin);
					break;
				}
			case (PT_MV_BINARY):
				{
					ULONG i = 0;

					DebugPrintEx(DBGGeneric,CLASS,_T("OnOpenProperty"),_T("property is PT_MV_BINARY\n"));
					if (S_OK == hRes && lpProp && PT_MV_BINARY == PROP_TYPE(lpProp->ulPropTag))
					{
						DebugPrintEx(DBGGeneric,CLASS,_T("OnOpenProperty"),_T("opened MV structure. There are 0x%X binaries in it.\n"),lpProp->Value.MVbin.cValues);
						for (i = 0; i<lpProp->Value.MVbin.cValues; i++)
						{
							m_lpHostDlg->OnOpenEntryID(&lpProp->Value.MVbin.lpbin[i]);
						}
					}
					break;
				}
			}
		}
	}
	if (m_lpMAPIProp)
	{
		MAPIFreeBuffer(lpProp);
	}
	return;
}//OnOpenProperty

void CSingleMAPIPropListCtrl::OnModifyExtraProps()
{
	if (!m_lpHostDlg || !m_lpHostDlg->m_lpMapiObjects) return;
	m_lpHostDlg->m_lpMapiObjects->MAPIInitialize(NULL);

	HRESULT hRes = S_OK;

	CTagArrayEditor MyTagArrayEditor(
		this,
		IDS_EXTRAPROPS,
		IDS_EXTRAPROPSPROMPT,
		m_sptExtraProps,
		m_bIsAB,
		m_lpMAPIProp);

	WC_H(MyTagArrayEditor.DisplayDialog());
	if (S_OK != hRes) return;

	LPSPropTagArray lpNewTagArray = MyTagArrayEditor.DetachModifiedTagArray();
	if (lpNewTagArray)
	{
		MAPIFreeBuffer(m_sptExtraProps);
		m_sptExtraProps = lpNewTagArray;
	}

	WC_H(RefreshMAPIPropList());
}//CSingleMAPIPropListCtrl::OnModifyExtraProps

void CSingleMAPIPropListCtrl::OnEditGivenProperty()
{
	if (!m_lpMAPIProp && !GetPropVals()) return;

	//Display a dialog to get a property number.
	HRESULT		hRes = S_OK;

	CPropertyTagEditor MyPropertyTag(
		IDS_EDITGIVENPROP,
		NULL,//prompt
		NULL,
		m_bIsAB,
		m_lpMAPIProp,
		this);

	WC_H(MyPropertyTag.DisplayDialog());
	if (S_OK == hRes)
	{
		OnEditGivenProp(MyPropertyTag.GetPropertyTag());
	}
}//CSingleMAPIPropListCtrl::OnEditGivenProperty

void CSingleMAPIPropListCtrl::OnOpenPropertyAsTable()
{
	if (!m_lpMAPIProp) return;

	//Display a dialog to get a property number.
	HRESULT		hRes = S_OK;
	CPropertyTagEditor MyPropertyTag(
		IDS_OPENPROPASTABLE,
		NULL,//prompt
		NULL,
		m_bIsAB,
		m_lpMAPIProp,
		this);

	WC_H(MyPropertyTag.DisplayDialog());
	if (S_OK == hRes)
	{
		CEditor MyData(
			this,
			IDS_OPENPROPASTABLE,
			IDS_OPENPROPASTABLEPROMPT,
			1,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

		MyData.InitCheck(0,IDS_OPENASEXTABLE,false,false);

		WC_H(MyData.DisplayDialog());
		if (S_OK == hRes)
		{
			if (MyData.GetCheck(0))
			{
				EC_H(DisplayExchangeTable(
					m_lpMAPIProp,
					CHANGE_PROP_TYPE(MyPropertyTag.GetPropertyTag(),PT_OBJECT),
					otDefault,
					m_lpHostDlg));
			}
			else
			{
				EC_H(DisplayTable(
					m_lpMAPIProp,
					CHANGE_PROP_TYPE(MyPropertyTag.GetPropertyTag(),PT_OBJECT),
					otDefault,
					m_lpHostDlg));
			}
		}
	}
}//OnOpenPropertyAsTable

void CSingleMAPIPropListCtrl::OnPasteNamedProps()
{
	if (!m_lpMAPIProp || !m_lpHostDlg || !m_lpHostDlg->m_lpMapiObjects) return;

	HRESULT		hRes = S_OK;
	LPENTRYLIST	lpSourceMsgEID = m_lpHostDlg->m_lpMapiObjects->GetMessagesToCopy();

	if (m_lpHostDlg->m_lpMapiObjects->GetBufferStatus() & BUFFER_MESSAGES
		&& lpSourceMsgEID
		&& 1 == lpSourceMsgEID->cValues)
	{
		CEditor MyData(
			this,
			IDS_PASTENAMEDPROPS,
			IDS_PASTENAMEDPROPSPROMPT,
			3,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

		LPTSTR szGuid = GUIDToStringAndName(&PS_PUBLIC_STRINGS);
		MyData.InitSingleLineSz(0,IDS_GUID,szGuid,false);
		delete[] szGuid;
		szGuid = NULL;
		MyData.InitCheck(1,IDS_MAPIMOVE,false,false);
		MyData.InitCheck(2,IDS_MAPINOREPLACE,false,false);

		WC_H(MyData.DisplayDialog());

		if (S_OK == hRes)
		{
			ULONG		ulObjType = 0;
			LPMAPIPROP	lpSource = NULL;
			GUID		propSetGUID = {0};

			EC_H(StringToGUID(MyData.GetString(0), &propSetGUID));

			if (S_OK == hRes)
			{
				EC_H(CallOpenEntry(
					NULL,
					NULL,
					m_lpHostDlg->m_lpMapiObjects->GetSourceParentFolder(),
					NULL,
					lpSourceMsgEID->lpbin,
					NULL,
					MAPI_BEST_ACCESS,
					&ulObjType,
					(LPUNKNOWN*)&lpSource));

				if(S_OK == hRes && MAPI_MESSAGE == ulObjType && lpSource)
				{
					EC_H(CopyNamedProps(
						lpSource,
						&propSetGUID,
						MyData.GetCheck(1),
						MyData.GetCheck(2),
						m_lpMAPIProp,
						m_lpHostDlg->m_hWnd));

					EC_H(m_lpMAPIProp->SaveChanges(KEEP_OPEN_READWRITE));

					RefreshMAPIPropList();

					lpSource->Release();
				}
			}
		}
	}
}

BOOL CSingleMAPIPropListCtrl::HandleAddInMenu(WORD wMenuSelect)
{
	if (wMenuSelect < ID_ADDINPROPERTYMENU) return false;
	CWaitCursor	Wait;//Change the mouse to an hourglass while we work.

	LPMENUITEM lpAddInMenu = GetAddinMenuItem(m_lpHostDlg->m_hWnd,wMenuSelect);
	if (!lpAddInMenu) return false;

	_AddInMenuParams MyAddInMenuParams = {0};
	MyAddInMenuParams.lpAddInMenu = lpAddInMenu;
	MyAddInMenuParams.ulAddInContext = MENU_CONTEXT_PROPERTY;
	MyAddInMenuParams.hWndParent = m_hWnd;
	MyAddInMenuParams.lpMAPIProp = m_lpMAPIProp;
	if (m_lpHostDlg && m_lpHostDlg->m_lpMapiObjects)
	{
		MyAddInMenuParams.lpMAPISession = m_lpHostDlg->m_lpMapiObjects->GetSession();//do not release
		MyAddInMenuParams.lpMDB = m_lpHostDlg->m_lpMapiObjects->GetMDB();//do not release
		MyAddInMenuParams.lpAdrBook = m_lpHostDlg->m_lpMapiObjects->GetAddrBook(false);//do not release
	}

	SRow MyRow = {0};
	MyRow.cValues = GetCountPropVals();
	MyRow.lpProps = GetPropVals();
	MyAddInMenuParams.lpRow = &MyRow;
	MyAddInMenuParams.ulCurrentFlags |= MENU_FLAGS_ROW;

	GetSelectedPropTag(&MyAddInMenuParams.ulPropTag);

	InvokeAddInMenu(&MyAddInMenuParams);
	return true;
} // CSingleMAPIPropListCtrl::HandleAddInMenu