// ContentsTableListCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "SortListCtrl.h"
#include "ContentsTableListCtrl.h"
#include "MapiObjects.h"
#include "ContentsTableDlg.h"
#include "MAPIFunctions.h"
#include "MFCUtilityFunctions.h"
#include "AdviseSink.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"
#include "Editor.h"
#include "SingleMAPIPropListCtrl.h"
#include "TagArrayEditor.h"
#include "Guids.h"
#include "PropTagArray.h"
#include <process.h>

static TCHAR* CLASS = _T("CContentsTableListCtrl");

#define NODISPLAYNAME 0xffffffff
/////////////////////////////////////////////////////////////////////////////
// CContentsTableListCtrl

CContentsTableListCtrl::CContentsTableListCtrl(
									 CWnd* pCreateParent,
									 CMapiObjects* lpMapiObjects,
									 LPSPropTagArray	sptExtraColumnTags,
									 ULONG ulNumExtraDisplayColumns,
									 TagNames *lpExtraDisplayColumns,
									 UINT nIDContextMenu,
									 BOOL bIsAB,
									 CContentsTableDlg *lpHostDlg)
									 :CSortListCtrl()
{
	TRACE_CONSTRUCTOR(CLASS);

	HRESULT hRes = S_OK;

	EC_H(Create(pCreateParent,NULL,IDC_LIST_CTRL,true));

	m_bAbortLoad = false; // no need to synchronize this - the thread hasn't started yet
	m_bInLoadOp = FALSE;
	m_LoadThreadHandle = NULL;

	// We borrow our parent's Mapi objects
	m_lpMapiObjects = lpMapiObjects;
	if (m_lpMapiObjects) m_lpMapiObjects->AddRef();

	m_lpHostDlg = lpHostDlg;
	if (m_lpHostDlg) m_lpHostDlg->AddRef();

	m_sptExtraColumnTags = sptExtraColumnTags;
	m_ulNumExtraDisplayColumns = ulNumExtraDisplayColumns;
	m_lpExtraDisplayColumns = lpExtraDisplayColumns;
	m_ulDisplayFlags = dfNormal;
	m_ulDisplayNameColumn = NODISPLAYNAME;

	m_ulHeaderColumns = 0;
	m_RestrictionType = mfcmapiNO_RESTRICTION;
	m_lpRes = NULL;
	m_lpContentsTable = NULL;
	m_ulContainerType = NULL;
	m_ulAdviseConnection = 0;
	m_lpAdviseSink = NULL;
	m_nIDContextMenu = nIDContextMenu;
	m_bIsAB = bIsAB;
}

CContentsTableListCtrl::~CContentsTableListCtrl()
{
	TRACE_DESTRUCTOR(CLASS);

	if (m_LoadThreadHandle) CloseHandle(m_LoadThreadHandle);

	NotificationOff();

	MAPIFreeBuffer(m_lpRes);
	if (m_lpContentsTable) m_lpContentsTable->Release();
	if (m_lpMapiObjects) m_lpMapiObjects->Release();
	if (m_lpHostDlg) m_lpHostDlg->Release();
}

BEGIN_MESSAGE_MAP(CContentsTableListCtrl, CSortListCtrl)
	ON_NOTIFY_REFLECT(LVN_ITEMCHANGED, OnItemChanged)
	ON_WM_KEYDOWN()
	ON_WM_CONTEXTMENU()
    ON_MESSAGE(WM_MFCMAPI_ADDITEM, msgOnAddItem)
	ON_MESSAGE(WM_MFCMAPI_THREADADDITEM, msgOnThreadAddItem)
	ON_MESSAGE(WM_MFCMAPI_DELETEITEM, msgOnDeleteItem)
	ON_MESSAGE(WM_MFCMAPI_MODIFYITEM, msgOnModifyItem)
	ON_MESSAGE(WM_MFCMAPI_REFRESHTABLE, msgOnRefreshTable)
END_MESSAGE_MAP()

// Some message handling's not worth breaking out into separate functions.
LRESULT CContentsTableListCtrl::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	HRESULT hRes = S_OK;

	switch (message)
	{
	case WM_LBUTTONDBLCLK:
		WC_H(DoExpandCollapse());
		if (S_FALSE == hRes)
		{
			// Post the message to display the item
			if (m_lpHostDlg)
				m_lpHostDlg->PostMessage(WM_COMMAND,ID_DISPLAYSELECTEDITEM,NULL);
		}
		else
		{
			CHECKHRESMSG(hRes,IDS_EXPANDCOLLAPSEFAILED);
		}
		return NULL;
		break;
	} // end switch
	return CSortListCtrl::WindowProc(message,wParam,lParam);
}
/////////////////////////////////////////////////////////////////////////////
// CContentsTableListCtrl message handlers

void CContentsTableListCtrl::OnContextMenu(CWnd* /*pWnd*/, CPoint pos)
{
	if (m_nIDContextMenu)
	{
		DisplayContextMenu(m_nIDContextMenu,IDR_MENU_TABLE,m_lpHostDlg, pos.x, pos.y);
	}
	else
	{
		DisplayContextMenu(IDR_MENU_DEFAULT_POPUP,IDR_MENU_TABLE,m_lpHostDlg, pos.x, pos.y);
	}
} // CContentsTableListCtrl::OnContextMenu

ULONG CContentsTableListCtrl::GetContainerType()
{
	return m_ulContainerType;
} // CContentsTableListCtrl::GetContainerType

BOOL CContentsTableListCtrl::IsContentsTableSet()
{
	return m_lpContentsTable?true:false;
} // CContentsTableListCtrl::IsContentsTableSet

HRESULT CContentsTableListCtrl::SetContentsTable(
						 LPMAPITABLE lpContentsTable,
						 ULONG ulDisplayFlags,
						 ULONG ulContainerType)
{
	HRESULT	hRes = S_OK;

	// If nothing to do, exit early
	if (lpContentsTable == m_lpContentsTable) return S_OK;
	if (m_bInLoadOp) return MAPI_E_INVALID_PARAMETER;

	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	m_ulDisplayFlags = ulDisplayFlags;
	m_ulContainerType = ulContainerType;

	DebugPrintEx(DBGGeneric,CLASS,_T("SetContentsTable"),_T("replacing 0x%X with 0x%X\n"),m_lpContentsTable,lpContentsTable);
	DebugPrintEx(DBGGeneric,CLASS,_T("SetContentsTable"),_T("New container type: 0x%X\n"),m_ulContainerType);
	// Clean up the old contents table and grab the new one
	if (m_lpContentsTable)
	{
		// If we don't Unadvise before releasing our reference, we'll leak an advise sink
		NotificationOff();

		m_lpContentsTable->Release();
		m_lpContentsTable = NULL;

		// Clear out the selected item view since we have a new contents table
		if (m_lpHostDlg)
			m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(NULL, NULL);
	}
	m_lpContentsTable = lpContentsTable;
	if (m_lpContentsTable)
	{
		m_lpContentsTable->AddRef();

		// Set up the columns on the new contents table
		DoSetColumns(
			true,
			RegKeys[regkeyEDIT_COLUMNS_ON_LOAD].ulCurDWORD,
			false,
			false);

		// Turn on notification...
		WC_H(NotificationOn());
		hRes = S_OK;
	}
	else
	{
		// Delete all of the old column headers
		// This way, when we clear a view, we clean up all the header data
		DeleteAllColumns();
	}
	// ...and pull the list into our nicely formatted table!
	EC_H(RefreshTable());

	return hRes;
} // CContentsTableListCtrl::SetContentsTable

void CContentsTableListCtrl::GetStatus()
{
	HRESULT hRes = S_OK;

	if (!IsContentsTableSet()) return;

	ULONG ulTableStatus = NULL;
	ULONG ulTableType = NULL;

	EC_H(m_lpContentsTable->GetStatus(
		&ulTableStatus,
		&ulTableType));

	if (!FAILED(hRes))
	{
		CEditor MyData(
			this,
			IDS_GETSTATUS,
			IDS_GETSTATUSPROMPT,
			4,
			CEDITOR_BUTTON_OK);
		MyData.InitSingleLine(0,IDS_ULTABLESTATUS,NULL,true);
		MyData.SetHex(0,ulTableStatus);
		LPTSTR szFlags = NULL;
		EC_H(InterpretFlags(flagTableStatus, ulTableStatus, &szFlags));
		MyData.InitMultiLine(1,IDS_ULTABLESTATUS,szFlags,true);
		delete[] szFlags;
		szFlags = NULL;

		MyData.InitSingleLine(2,IDS_ULTABLETYPE,NULL,true);
		MyData.SetHex(2,ulTableType);
		EC_H(InterpretFlags(flagTableType,ulTableType, &szFlags));
		MyData.InitMultiLine(3,IDS_ULTABLETYPE,szFlags,true);
		delete[] szFlags;
		szFlags = NULL;

		WC_H(MyData.DisplayDialog());
	}
} // CContentsTableListCtrl::GetStatus

// Takes a tag array and builds the UI out of it - does NOT touch the table
HRESULT CContentsTableListCtrl::SetUIColumns(LPSPropTagArray lpTags)
{
	HRESULT			hRes = S_OK;
	if (!lpTags) return MAPI_E_INVALID_PARAMETER;

	// find a PR_DISPLAY_NAME column for later use
	m_ulDisplayNameColumn = NODISPLAYNAME;
	for (ULONG i = 0 ; i < lpTags->cValues ; i++)
	{
		if (PROP_ID(lpTags->aulPropTag[i]) == PROP_ID(PR_DISPLAY_NAME))
		{
			m_ulDisplayNameColumn = i;
			break;
		}
	}

	// Didn't find display name - fall back on some other columns
	if (NODISPLAYNAME == m_ulDisplayNameColumn)
	{
		for (ULONG i = 0 ; i < lpTags->cValues ; i++)
		{
			if (PROP_ID(lpTags->aulPropTag[i]) == PROP_ID(PR_SUBJECT) ||
				PROP_ID(lpTags->aulPropTag[i]) == PROP_ID(PR_RULE_NAME) ||
				PROP_ID(lpTags->aulPropTag[i]) == PROP_ID(PR_MEMBER_NAME))

			{
				m_ulDisplayNameColumn = i;
				break;
			}
		}
	}

	DebugPrintEx(DBGGeneric,CLASS,_T("SetColumns"),_T("calculating and inserting column headers\n"));
	MySetRedraw(FALSE);

	// Delete all of the old column headers
	DeleteAllColumns();

	hRes = S_OK;
	EC_H(AddColumns(lpTags));

	AutoSizeColumns();

	DebugPrintEx(DBGGeneric,CLASS,_T("SetColumns"),_T("Done inserting column headers\n"));

	MySetRedraw(TRUE);
	return hRes;
}


void CContentsTableListCtrl::DoSetColumns(BOOL bAddExtras, BOOL bDisplayEditor, BOOL bQueryFlags, BOOL bDoRefresh)
{
	HRESULT			hRes = S_OK;

	if (!IsContentsTableSet()) return;

	// these arrays get allocated during the func and need to be freed
	LPSPropTagArray lpConcatTagArray = NULL;
	LPSPropTagArray lpModifiedTags = NULL;
	LPSPropTagArray lpOriginalColSet = NULL;

	// this is just a pointer - do not free
	LPSPropTagArray lpFinalTagArray = lpOriginalColSet;

	ULONG ulQueryColumnFlags = NULL;
	ULONG ulSetColumnFlags = TBL_BATCH;

	CWaitCursor		Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric,CLASS,_T("DoSetColumns"),
		_T("bAddExtras = %d, bDisplayEditor = %d, bQueryFlags = %d, bDoRefresh = %d\n"), // STRING_OK
		bAddExtras,
		bDisplayEditor,
		bQueryFlags,
		bDoRefresh);

	if (bQueryFlags)
	{
		CEditor MyData(
			this,
			IDS_SETCOLUMNS,
			IDS_SETCOLUMNSPROMPT,
			2,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitSingleLine(0,IDS_QUERYCOLUMNFLAGS,NULL,false);
		MyData.SetHex(0,ulQueryColumnFlags );
		MyData.InitSingleLine(1,IDS_SETCOLUMNFLAGS,NULL,false);
		MyData.SetHex(1,ulSetColumnFlags);

		WC_H(MyData.DisplayDialog());
		if (S_OK == hRes)
		{
			ulQueryColumnFlags = MyData.GetHex(0);
			ulSetColumnFlags = MyData.GetHex(1);
		}
	}

	if (S_OK == hRes)
	{
		EC_H(m_lpContentsTable->QueryColumns(
			ulQueryColumnFlags,
			&lpOriginalColSet));
		hRes = S_OK;

		lpFinalTagArray = lpOriginalColSet;
	}

	if (bAddExtras)
	{
		// build an array with the source set and m_sptExtraColumnTags combined
		EC_H(ConcatSPropTagArrays(
			m_sptExtraColumnTags,
			lpFinalTagArray, // build on the final array we've computed thus far
			&lpConcatTagArray));
		lpFinalTagArray = lpConcatTagArray;
	}

	if (bDisplayEditor)
	{
		LPMDB lpMDB = NULL;
		if (m_lpMapiObjects)
		{
			lpMDB = m_lpMapiObjects->GetMDB(); // do not release
		}

		CTagArrayEditor MyEditor(
			this,
			IDS_COLUMNSET,
			IDS_COLUMNSETPROMPT,
			lpFinalTagArray, // build on the final array we've computed thus far
			m_bIsAB,
			lpMDB);

		WC_H(MyEditor.DisplayDialog());

		if (S_OK == hRes)
		{
			lpModifiedTags = MyEditor.DetachModifiedTagArray();
			if (lpModifiedTags)
			{
				lpFinalTagArray = lpModifiedTags;
			}
		}
		hRes = S_OK;
	}

	if (lpFinalTagArray)
	{
		// Apply lpFinalTagArray through SetColumns
		EC_H(m_lpContentsTable->SetColumns(
			lpFinalTagArray,
			ulSetColumnFlags)); // Flags
		EC_H(SetUIColumns(lpFinalTagArray));
		if (bDoRefresh) EC_H(RefreshTable());
	}

	MAPIFreeBuffer(lpModifiedTags);
	MAPIFreeBuffer(lpConcatTagArray);
	MAPIFreeBuffer(lpOriginalColSet);
} // CContentsTableListCtrl::DoSetColumns

HRESULT CContentsTableListCtrl::AddColumn(UINT uidHeaderName, ULONG ulCurHeaderCol, ULONG ulCurTagArrayRow, ULONG ulPropTag)
{
	HRESULT			hRes = S_OK;
	CHeaderCtrl*	lpMyHeader = NULL;
	int				iRetVal = NULL;
	HDITEM			hdItem = {0};
	lpMyHeader = GetHeaderCtrl();

	CString szHeaderString;

	if (uidHeaderName)
	{
		szHeaderString.LoadString(uidHeaderName);
	}
	else
	{
		LPTSTR szExactMatches = NULL;
		EC_H(PropTagToPropName(ulPropTag,m_bIsAB,&szExactMatches,NULL));
		if (!szExactMatches)
			szHeaderString.Format(_T("0x%08X"),ulPropTag); // STRING_OK
		else
			szHeaderString = szExactMatches;
		delete[] szExactMatches;
	}

	iRetVal = InsertColumn(ulCurHeaderCol,szHeaderString);

	if (-1 == iRetVal)
	{
		// We failed to insert a column header
		ErrDialog(__FILE__,__LINE__,IDS_EDCOLUMNHEADERFAILED);
	}

	if (lpMyHeader)
	{
		hdItem.mask = HDI_LPARAM;
		LPHEADERDATA lpHeaderData = NULL;

		lpHeaderData = new HeaderData; // Will be deleted in CSortListCtrl::DeleteAllColumns
		if (lpHeaderData)
		{
			LPMDB lpMDB = NULL;
			if (m_lpMapiObjects) lpMDB = m_lpMapiObjects->GetMDB(); // do not release

			lpHeaderData->ulTagArrayRow = ulCurTagArrayRow;
			lpHeaderData->ulPropTag = ulPropTag;
			lpHeaderData->bIsAB = m_bIsAB;
			EC_H(StringCchCopy(lpHeaderData->szTipString,_countof(lpHeaderData->szTipString),
				(LPCTSTR) TagToString(ulPropTag,lpMDB,m_bIsAB,false)));

			hdItem.lParam = (LPARAM) lpHeaderData;
			EC_B(lpMyHeader->SetItem(ulCurHeaderCol,&hdItem));
		}
	}

	return hRes;
} // CContentsTableListCtrl::AddColumn

// Sets up column headers based on passed in named columns
// Put all named columns first, followed by a column for each property in the contents table
HRESULT CContentsTableListCtrl::AddColumns(LPSPropTagArray lpCurColTagArray)
{
	HRESULT			hRes = S_OK;

	if (!lpCurColTagArray || !m_lpHostDlg) return MAPI_E_INVALID_PARAMETER;

	m_ulHeaderColumns = lpCurColTagArray->cValues;

	ULONG	ulCurHeaderCol = 0;
	if (RegKeys[regkeyDO_COLUMN_NAMES].ulCurDWORD)
	{
		DebugPrintEx(DBGGeneric,CLASS,_T("AddColumns"),_T("Adding named columns\n"));
		// If we have named columns, put them up front
		if (m_lpExtraDisplayColumns)
		{
			ULONG	ulCurExtraCol = 0;
			// Walk through the list of named/extra columns and add them to our header list
			for (ulCurExtraCol = 0; ulCurExtraCol<m_ulNumExtraDisplayColumns; ulCurExtraCol++)
			{
				ULONG ulExtraColRowNum = m_lpExtraDisplayColumns[ulCurExtraCol].ulMatchingTableColumn;
				ULONG ulExtraColTag = m_sptExtraColumnTags->aulPropTag[ulExtraColRowNum];

				ULONG ulCurTagArrayRow = 0;
				if (FindPropInPropTagArray(lpCurColTagArray,ulExtraColTag,&ulCurTagArrayRow))
				{
					hRes = S_OK;
					EC_H(AddColumn(
						m_lpExtraDisplayColumns[ulCurExtraCol].uidName,
						ulCurHeaderCol,
						ulCurTagArrayRow,
						lpCurColTagArray->aulPropTag[ulCurTagArrayRow]));
					// Strike out the value in the tag array so we can ignore it later!
					lpCurColTagArray->aulPropTag[ulCurTagArrayRow] = NULL;

					ulCurHeaderCol++;
				}
			}
		}
	}

	ULONG ulCurTableCol = 0;
	DebugPrintEx(DBGGeneric,CLASS,_T("AddColumns"),_T("Adding unnamed columns\n"));
	// Now, walk through the current tag table and add each unstruck column to our list
	for (ulCurTableCol = 0; ulCurTableCol < lpCurColTagArray->cValues; ulCurTableCol++)
	{
		if (lpCurColTagArray->aulPropTag[ulCurTableCol] != NULL)
		{
			hRes = S_OK;
			EC_H(AddColumn(
				NULL,
				ulCurHeaderCol,
				ulCurTableCol,
				lpCurColTagArray->aulPropTag[ulCurTableCol]));
			ulCurHeaderCol++;
		}
	}

	// this would be bad
	if (ulCurHeaderCol<m_ulHeaderColumns)
	{
		ErrDialog(__FILE__,__LINE__,IDS_EDTOOMANYCOLUMNS);
	}

	DebugPrintEx(DBGGeneric,CLASS,_T("AddColumns"),_T("Done adding columns\n"));
	return hRes;
} // CContentsTableListCtrl::AddColumns

void CContentsTableListCtrl::SetRestriction(LPSRestriction lpRes)
{
	MAPIFreeBuffer(m_lpRes);
	m_lpRes = lpRes;
} // CContentsTableListCtrl::SetRestriction

LPSRestriction CContentsTableListCtrl::GetRestriction()
{
	return m_lpRes;
} // CContentsTableListCtrl::GetRestriction

__mfcmapiRestrictionTypeEnum CContentsTableListCtrl::GetRestrictionType()
{
	return m_RestrictionType;
} // CContentsTableListCtrl::GetRestrictionType

void CContentsTableListCtrl::SetRestrictionType(__mfcmapiRestrictionTypeEnum RestrictionType)
{
	m_RestrictionType = RestrictionType;
} // CContentsTableListCtrl::SetRestrictionType

HRESULT CContentsTableListCtrl::ApplyRestriction()
{
	if (!m_lpContentsTable) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;
	DebugPrintEx(DBGGeneric,CLASS,_T("ApplyRestriction"),_T("m_RestrictionType = 0x%X\n"),m_RestrictionType);
	// Apply our restrictions
	if (mfcmapiNORMAL_RESTRICTION == m_RestrictionType)
	{
		DebugPrintEx(DBGGeneric,CLASS,_T("ApplyRestriction"),_T("applying restriction:\n"));

		if (m_lpMapiObjects)
		{
			LPMDB lpMDB = m_lpMapiObjects->GetMDB(); // do not release
			DebugPrintRestriction(DBGGeneric,m_lpRes,lpMDB);
		}

		EC_H(m_lpContentsTable->Restrict(
			m_lpRes,
			TBL_BATCH));
	}
	else
	{
		WC_H_MSG(m_lpContentsTable->Restrict(
			NULL,
			TBL_BATCH),
			IDS_TABLENOSUPPORTRES);
	}

	return hRes;
}

enum LoadTableStateEnum { ltsWorking, ltsDone};

struct ThreadLoadTableInfo
{
	HWND							hWndHost;
	CContentsTableListCtrl*			lpListCtrl;
	LPMAPITABLE						lpContentsTable;
	LONG volatile*					lpbAbort;
};

#define bABORTSET (*lpThreadInfo->lpbAbort) // This is safe
#define BREAKONABORT if (bABORTSET) break;
#define CHECKABORT(__fn) if (!bABORTSET) {__fn;}
#define NUMROWSPERLOOP 255

// Idea here is to do our MAPI work here on this thread, then send messages (SendMessage) back to the control to add the data to the view
// This way, control functions only happen on the main thread
// ::SendMessage will be handled on main thread, but block until the call returns.
// This is the ideal behavior for this worker thread.
unsigned STDAPICALLTYPE ThreadFuncLoadTable(void* lpParam)
{
	HRESULT					hRes = S_OK;
	ULONG					ulTotal = 0;
	ULONG					ulThrottleLevel = 0;
	LPSRowSet				pRows = NULL;
	ULONG					iCurPropRow = 0;
	ULONG					iCurListBoxRow = 0;
	ThreadLoadTableInfo*	lpThreadInfo = (ThreadLoadTableInfo*) lpParam;
	if (!lpThreadInfo || !lpThreadInfo->lpbAbort)	return 0;

	CContentsTableListCtrl* lpListCtrl = lpThreadInfo->lpListCtrl;
	LPMAPITABLE				lpContentsTable = lpThreadInfo->lpContentsTable;
	if (!lpListCtrl || !lpContentsTable) return 0;

	HWND					hWndHost = lpThreadInfo->hWndHost;
	CString					szStatusText;

	// required on da new thread before we do any MAPI work
	EC_H(MAPIInitialize(NULL));

	(void)::SendMessage(hWndHost,WM_MFCMAPI_CLEARSINGLEMAPIPROPLIST,NULL,NULL);
	szStatusText.Format(IDS_STATUSTEXTNUMITEMS,lpListCtrl->GetItemCount());
	(void)::SendMessage(hWndHost,WM_MFCMAPI_UPDATESTATUSBAR,STATUSLEFTPANE,(LPARAM)(LPCTSTR) szStatusText);

	// potentially lengthy op - check abort before and after
	CHECKABORT(WC_H(lpListCtrl->ApplyRestriction()));
	hRes = S_OK; // Don't care if the restrict failed - let's try to go on

	if (!bABORTSET) // only check abort once for this group of ops
	{
		// go to the first row
		EC_H(lpContentsTable->SeekRow(
			BOOKMARK_BEGINNING,
			0,
			NULL));
		hRes = S_OK; // don't let failure here fail the whole load

		EC_H(lpContentsTable->GetRowCount(
			NULL,
			&ulTotal));
		hRes = S_OK; // don't let failure here fail the whole load

		DebugPrintEx(DBGGeneric,CLASS,_T("ThreadFuncLoadTable"),_T("ulTotal = 0x%X\n"),ulTotal);

		ulThrottleLevel = RegKeys[regkeyTHROTTLE_LEVEL].ulCurDWORD;

		szStatusText.FormatMessage(IDS_LOADINGITEMS,0,ulTotal);
		(void)::SendMessage(hWndHost,WM_MFCMAPI_UPDATESTATUSBAR,STATUSMIDDLEPANE,(LPARAM)(LPCTSTR) szStatusText);
	}

	LPSRestriction lpRes = lpListCtrl->GetRestriction();
	// get rows and add them to the list
	if (!FAILED(hRes)) for (;;)
	{
		BREAKONABORT;
		szStatusText.LoadString(IDS_ESCSTOPLOADING);
		(void)::SendMessage(hWndHost,WM_MFCMAPI_UPDATESTATUSBAR,STATUSRIGHTPANE,(LPARAM)(LPCTSTR) szStatusText);
		hRes = S_OK;
		if (pRows) FreeProws(pRows);
		pRows = NULL;
		if (mfcmapiFINDROW_RESTRICTION == lpListCtrl->GetRestrictionType() && lpRes)
		{
			DebugPrintEx(DBGGeneric,CLASS,_T("DoFindRows"),_T("running FindRow with restriction:\n"));
			DebugPrintRestriction(DBGGeneric,lpRes,NULL);

			CHECKABORT(WC_H(lpContentsTable->FindRow(
				lpRes,
				BOOKMARK_CURRENT,
				NULL)));

			if (MAPI_E_NOT_FOUND != hRes) // MAPI_E_NOT_FOUND signals we didn't find any more rows.
			{
				CHECKABORT(EC_H(lpContentsTable->QueryRows(
					1,
					NULL,
					&pRows)));
			}
			else
			{
				hRes = S_OK;
				break;
			}
		}
		else
		{
			DebugPrintEx(DBGGeneric,CLASS,_T("ThreadFuncLoadTable"),_T("Calling QueryRows. Asking for 0x%X rows.\n"),(ulThrottleLevel)?ulThrottleLevel:NUMROWSPERLOOP);
			// Pull back a sizable block of rows to add to the list box
			CHECKABORT(EC_H(lpContentsTable->QueryRows(
				(ulThrottleLevel)?ulThrottleLevel:NUMROWSPERLOOP,
				NULL,
				&pRows)));
			if (FAILED(hRes)) break;
		}
		if (FAILED(hRes) || !pRows || !pRows->cRows) break;

		DebugPrintEx(DBGGeneric,CLASS,_T("ThreadFuncLoadTable"),_T("Got this many rows: 0x%X\n"),pRows->cRows);
		for (iCurPropRow = 0;iCurPropRow<pRows->cRows;iCurPropRow++)
		{
			hRes = S_OK;
			BREAKONABORT; // This check is cheap enough not to be a perf concern anymore
			szStatusText.FormatMessage(IDS_LOADINGITEMS,iCurListBoxRow+1,ulTotal);
			(void)::SendMessage(hWndHost,WM_MFCMAPI_UPDATESTATUSBAR,STATUSMIDDLEPANE,(LPARAM)(LPCTSTR) szStatusText);

			DebugPrintEx(DBGGeneric,CLASS,_T("ThreadFuncLoadTable"),_T("Asking to add 0x%08X to %d\n"),&pRows->aRow[iCurPropRow],iCurListBoxRow);
			(void)::SendMessage(lpListCtrl->m_hWnd,WM_MFCMAPI_THREADADDITEM,iCurListBoxRow,(LPARAM)&pRows->aRow[iCurPropRow]);
			if (FAILED(hRes)) continue;
			iCurListBoxRow++;
		}

		// Note - we're saving the rows off, so we don't FreeProws this...we just MAPIFreeBuffer the array
		MAPIFreeBuffer(pRows);
		pRows = NULL;

		if (ulThrottleLevel && iCurListBoxRow >= ulThrottleLevel) break; // Only render ulThrottleLevel rows if throttle is on
	}

	if (bABORTSET)
	{
		szStatusText.LoadString(IDS_TABLELOADCANCELLED);
		(void)::SendMessage(hWndHost,WM_MFCMAPI_UPDATESTATUSBAR,STATUSRIGHTPANE,(LPARAM)(LPCTSTR) szStatusText);
	}
	else
	{
		szStatusText.LoadString(IDS_TABLELOADED);
		(void)::SendMessage(hWndHost,WM_MFCMAPI_UPDATESTATUSBAR,STATUSRIGHTPANE,(LPARAM)(LPCTSTR) szStatusText);
	}
	(void)::SendMessage(hWndHost,WM_MFCMAPI_UPDATESTATUSBAR,STATUSMIDDLEPANE,(LPARAM)_T(""));
	DebugPrintEx(DBGGeneric,CLASS,_T("ThreadFuncLoadTable"),_T("added %d items\n"),iCurListBoxRow);

	DebugPrintEx(DBGGeneric,CLASS,_T("ThreadFuncLoadTable"),_T("Releasing pointers.\n"));

	lpListCtrl->ClearLoading();

	// Bunch of cleanup
	if (pRows) FreeProws(pRows);
	if (lpContentsTable) lpContentsTable->Release();
	if (lpListCtrl) lpListCtrl->Release();
	DebugPrintEx(DBGGeneric,CLASS,_T("ThreadFuncLoadTable"),_T("Pointers released.\n"));

	MAPIUninitialize();

	delete lpThreadInfo;

	return 0;
} // ThreadFuncLoadTable

BOOL CContentsTableListCtrl::IsLoading()
{
	return m_bInLoadOp;
} // CContentsTableListCtrl::IsLoading

void CContentsTableListCtrl::ClearLoading()
{
	m_bInLoadOp = false;
} // CContentsTableListCtrl::ClearLoading

HRESULT CContentsTableListCtrl::LoadContentsTableIntoView()
{
	HRESULT			hRes = S_OK;
	CWaitCursor		Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric,CLASS,_T("LoadContentsTableIntoView"),_T("\n"));

	if (m_bInLoadOp) return MAPI_E_INVALID_PARAMETER;
	if (!this || !m_lpHostDlg) return MAPI_E_INVALID_PARAMETER;

	EC_B(DeleteAllItems());

	// whack the old thread handle if we still have it
	if (m_LoadThreadHandle) CloseHandle(m_LoadThreadHandle);
	m_LoadThreadHandle = NULL;

	if (!m_lpContentsTable) return S_OK;
	m_bInLoadOp = TRUE;
	// Do not call return after this point!

	ThreadLoadTableInfo* lpThreadInfo = 0;

	lpThreadInfo = new ThreadLoadTableInfo;

	if (lpThreadInfo)
	{
		lpThreadInfo->hWndHost = m_lpHostDlg->m_hWnd;
		lpThreadInfo->lpbAbort = &m_bAbortLoad;
		m_bAbortLoad = false; // no need to synchronize this - the thread hasn't started yet

		lpThreadInfo->lpListCtrl = this;
		if (this) this->AddRef();

		lpThreadInfo->lpContentsTable = m_lpContentsTable;
		m_lpContentsTable->AddRef();

		DebugPrintEx(DBGGeneric,CLASS,_T("LoadContentsTableIntoView"),_T("Creating load thread.\n"));

		HANDLE hThread = 0;
		EC_D(hThread,(HANDLE) _beginthreadex(NULL, 0, ThreadFuncLoadTable, lpThreadInfo, 0, 0));

		if (!hThread)
		{
			DebugPrintEx(DBGGeneric,CLASS,_T("LoadContentsTableIntoView"),_T("Load thread creation failed.\n"));
			if (lpThreadInfo->lpContentsTable) lpThreadInfo->lpContentsTable->Release();
			if (lpThreadInfo->lpListCtrl) lpThreadInfo->lpListCtrl->Release();
			delete lpThreadInfo;
		}
		else
		{
			DebugPrintEx(DBGGeneric,CLASS,_T("LoadContentsTableIntoView"),_T("Load thread created.\n"));
			m_LoadThreadHandle = hThread;
		}
	}

	return hRes;
} // CContentsTableListCtrl::LoadContentsTableIntoView

void CContentsTableListCtrl::OnCancelTableLoad()
{
	DebugPrintEx(DBGGeneric,CLASS,_T("OnCancelTableLoad"),_T("Setting abort flag and waiting for thread to discover it\n"));
	// Wait here until the thread we spun off has shut down
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.
	DWORD dwRet = 0;
	BOOL bVKF5Hit = false;

	// See if the thread is still active
	while (m_LoadThreadHandle) // this won't change, but if it's NULL, we just skip the loop
	{
		MSG msg;

		InterlockedExchange(&m_bAbortLoad,true);

		// Wait for the thread to shutdown/signal, or messages posted to our queue
		dwRet = MsgWaitForMultipleObjects(
			1,
			&m_LoadThreadHandle,
			FALSE,
			INFINITE,
			QS_ALLINPUT);
		if (dwRet == (WAIT_OBJECT_0 + 0)) break;

		// Read all of the messages in this next loop, removing each message as we read it.
		// If we don't do this, the thread never stops
		while (PeekMessage(&msg, m_hWnd, 0, 0, PM_REMOVE))
		{
			if (msg.message == WM_KEYDOWN && msg.wParam == VK_F5)
			{
				DebugPrintEx(DBGGeneric,CLASS,_T("OnCancelTableLoad"),_T("Ditching refresh (F5)\n"));
				bVKF5Hit = true;
			}
			else
			{
				DispatchMessage(&msg);
			}
		}
	}

	DebugPrintEx(DBGGeneric,CLASS,_T("OnCancelTableLoad"),_T("Load thread has shut down.\n"));

	if (m_LoadThreadHandle) CloseHandle(m_LoadThreadHandle);
	m_LoadThreadHandle = NULL;
	m_bAbortLoad = false;

	if (bVKF5Hit) // If we ditched a refresh message, repost it now
	{
		DebugPrintEx(DBGGeneric,CLASS,_T("OnCancelTableLoad"),_T("Posting skipped refresh message\n"));
		PostMessage(WM_KEYDOWN,VK_F5,0);
	}
} // CContentsTableListCtrl::OnCancelTableLoad

// Sets data from the LPSRow into the SortListData structure
// Assumes the structure is either an existing structure or a new one which has been memset to 0
// If it's an existing structure - we need to free up some memory
void CContentsTableListCtrl::BuildDataItem(LPSRow lpsRowData,SortListData* lpData)
{
	if (!lpData || !lpsRowData) return;

	HRESULT hRes = S_OK;
	LPSPropValue	lpProp = NULL; // do not free this

	lpData->bItemFullyLoaded = false;
	MAPIFreeBuffer(lpData->szSortText);
	lpData->szSortText = NULL;

	// this guy gets stolen from lpsRowData and is freed separately in FreeSortListData
	// So I do need to free it here before losing the pointer
	MAPIFreeBuffer(lpData->lpSourceProps);
	lpData->lpSourceProps = NULL;

	lpData->ulSortValue.QuadPart = NULL;
	lpData->cSourceProps = 0;
	lpData->ulSortDataType = SORTLIST_CONTENTS;
	memset(&lpData->data,0,sizeof(_ContentsData));

	// Save off the source props
	lpData->lpSourceProps = lpsRowData->lpProps;
	lpData->cSourceProps = lpsRowData->cValues;

	// Save the instance key into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_INSTANCE_KEY);
	if (lpProp && PR_INSTANCE_KEY == lpProp->ulPropTag)
	{
		EC_H(MAPIAllocateMore(
			(ULONG)sizeof(SBinary),
			lpData,
			(LPVOID*) &lpData->data.Contents.lpInstanceKey));
		EC_H(CopySBinary(lpData->data.Contents.lpInstanceKey,&lpProp->Value.bin, lpData));
	}

	// Save the attachment number into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ATTACH_NUM);
	if (lpProp && PR_ATTACH_NUM == lpProp->ulPropTag)
	{
		DebugPrint(DBGGeneric,_T("\tPR_ATTACH_NUM = %d\n"),lpProp->Value.l);
		lpData->data.Contents.ulAttachNum = lpProp->Value.l;
	}

	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ATTACH_METHOD);
	if (lpProp && PR_ATTACH_METHOD == lpProp->ulPropTag)
	{
		DebugPrint(DBGGeneric,_T("\tPR_ATTACH_METHOD = %d\n"),lpProp->Value.l);
		lpData->data.Contents.ulAttachMethod = lpProp->Value.l;
	}

	// Save the row ID (recipients) into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ROWID);
	if (lpProp && PR_ROWID == lpProp->ulPropTag)
	{
		DebugPrint(DBGGeneric,_T("\tPR_ROWID = %d\n"),lpProp->Value.l);
		lpData->data.Contents.ulRowID = lpProp->Value.l;
	}

	// Save the row type (header/leaf) into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ROW_TYPE);
	if (lpProp && PR_ROW_TYPE == lpProp->ulPropTag)
	{
		DebugPrint(DBGGeneric,_T("\tPR_ROW_TYPE = %d\n"),lpProp->Value.l);
		lpData->data.Contents.ulRowType = lpProp->Value.l;
	}

	// Save the Entry ID into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_ENTRYID);
	if (lpProp && PR_ENTRYID == lpProp->ulPropTag)
	{
		EC_H(MAPIAllocateMore(
			(ULONG)sizeof(SBinary),
			lpData,
			(LPVOID*) &lpData->data.Contents.lpEntryID));
		EC_H(CopySBinary(lpData->data.Contents.lpEntryID,&lpProp->Value.bin, lpData));
	}

	// Save the Longterm Entry ID into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_LONGTERM_ENTRYID_FROM_TABLE);
	if (lpProp && PR_LONGTERM_ENTRYID_FROM_TABLE == lpProp->ulPropTag)
	{
		EC_H(MAPIAllocateMore(
			(ULONG)sizeof(SBinary),
			lpData,
			(LPVOID*) &lpData->data.Contents.lpLongtermID));
		EC_H(CopySBinary(lpData->data.Contents.lpLongtermID,&lpProp->Value.bin, lpData));
	}

	// Save the Service ID into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_SERVICE_UID);
	if (lpProp && PR_SERVICE_UID == lpProp->ulPropTag)
	{
		// Allocate some space
		EC_H(MAPIAllocateMore(
			(ULONG)sizeof(SBinary),
			lpData,
			(LPVOID*) &lpData->data.Contents.lpServiceUID));
		EC_H(CopySBinary(lpData->data.Contents.lpServiceUID,&lpProp->Value.bin, lpData));
	}

	// Save the Provider ID into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_PROVIDER_UID);
	if (lpProp && PR_PROVIDER_UID == lpProp->ulPropTag)
	{
		// Allocate some space
		EC_H(MAPIAllocateMore(
			(ULONG)sizeof(SBinary),
			lpData,
			(LPVOID*) &lpData->data.Contents.lpProviderUID));
		EC_H(CopySBinary(lpData->data.Contents.lpProviderUID,&lpProp->Value.bin, lpData));
	}

	// Save the DisplayName into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_DISPLAY_NAME_A); // We pull this properties for profiles, which do not support Unicode
	if (CheckStringProp(lpProp,PT_STRING8))
	{
		DebugPrint(DBGGeneric,_T("\tPR_DISPLAY_NAME_A = %hs\n"),lpProp->Value.lpszA);

		EC_H(CopyStringA(
			&lpData->data.Contents.szProfileDisplayName,
			lpProp->Value.lpszA,
			lpData));
	}

	// Save the e-mail address (if it exists on the object) into lpData
	lpProp = PpropFindProp(
		lpsRowData->lpProps,
		lpsRowData->cValues,
		PR_EMAIL_ADDRESS);
	if (CheckStringProp(lpProp,PT_TSTRING))
	{
		DebugPrint(DBGGeneric,_T("\tPR_EMAIL_ADDRESS = %s\n"),lpProp->Value.LPSZ);
		EC_H(CopyString(
			&lpData->data.Contents.szDN,
			lpProp->Value.LPSZ,
			lpData));
	}
} // CContentsTableListCtrl::BuildDataItem

void CContentsTableListCtrl::SetRowStrings(int iRow,LPSRow lpsRowData)
{
	if (!lpsRowData) return;

	HRESULT		hRes = S_OK;
	ULONG		iColumn = 0;
	CHeaderCtrl* lpMyHeader = GetHeaderCtrl();

	if (!lpMyHeader) return;

	for (iColumn = 0;iColumn<m_ulHeaderColumns;iColumn++)
	{
		HDITEM hdItem = {0};
		hdItem.mask = HDI_LPARAM;
		EC_B(lpMyHeader->GetItem(iColumn,&hdItem));

		if (hdItem.lParam)
		{
			ULONG ulCol = ((LPHEADERDATA)hdItem.lParam)->ulTagArrayRow;
			hRes = S_OK;

			if (ulCol < lpsRowData->cValues)
			{
				CString PropString;
				LPTSTR szFlags = NULL;
				LPSPropValue pProp = &lpsRowData->lpProps[ulCol];

				InterpretProp(pProp,&PropString,NULL);

				EC_H(InterpretFlags(pProp, &szFlags));
				if (szFlags)
				{
					PropString += _T(" ("); // STRING_OK
					PropString += szFlags;
					PropString += _T(")"); // STRING_OK
				}
				delete[] szFlags;
				szFlags = NULL;

				WC_B(SetItemText(iRow,iColumn,(LPCTSTR) PropString));
			}
			else
			{
				// This is an odd case which just shouldn't happen.
				// If SetColumns failed in DoSetColumns, we might have columns
				// mapped past the end of the table. Just log the error and give up.
				WARNHRESMSG(MAPI_E_NOT_FOUND,IDS_COLOUTOFRANGE);
				break;
			}
		}
	}
} // CContentsTableListCtrl::SetRowStrings

#define NUMOBJTYPES 12
static LONG _ObjTypeIcons[NUMOBJTYPES][2] =
{
{MAPI_STORE,slIconMAPI_STORE},
{MAPI_ADDRBOOK,slIconMAPI_ADDRBOOK},
{MAPI_FOLDER,slIconMAPI_FOLDER},
{MAPI_ABCONT,slIconMAPI_ABCONT},
{MAPI_MESSAGE,slIconMAPI_MESSAGE},
{MAPI_MAILUSER,slIconMAPI_MAILUSER},
{MAPI_ATTACH,slIconMAPI_ATTACH},
{MAPI_DISTLIST,slIconMAPI_DISTLIST},
{MAPI_PROFSECT,slIconMAPI_PROFSECT},
{MAPI_STATUS,slIconMAPI_STATUS},
{MAPI_SESSION,slIconMAPI_SESSION},
{MAPI_FORMINFO,slIconMAPI_FORMINFO},
};

void GetDepthAndImage(LPSRow lpsRowData,ULONG* lpulDepth, ULONG* lpulImage)
{
	if (lpulDepth) *lpulDepth = 0;
	if (lpulImage) *lpulImage = slIconDefault;
	if (!lpsRowData) return;

	ULONG ulDepth = NULL;
	ULONG ulImage = slIconDefault;
	LPSPropValue lpDepth = NULL;
	lpDepth = PpropFindProp(lpsRowData->lpProps,lpsRowData->cValues,PR_DEPTH);
	if (lpDepth && PR_DEPTH == lpDepth->ulPropTag) ulDepth = lpDepth->Value.l;
	if (ulDepth > 5) ulDepth = 5; // Just in case

	LPSPropValue lpRowType = NULL;
	lpRowType = PpropFindProp(lpsRowData->lpProps,lpsRowData->cValues,PR_ROW_TYPE);
	if (lpRowType && PR_ROW_TYPE == lpRowType->ulPropTag)
	{
		switch(lpRowType->Value.l)
		{
		case TBL_LEAF_ROW:
			break;
		case TBL_EMPTY_CATEGORY:
		case TBL_COLLAPSED_CATEGORY:
			ulImage = slIconNodeCollapsed;
			break;
		case TBL_EXPANDED_CATEGORY:
			ulImage = slIconNodeExpanded;
			break;
		}
	}

	if (slIconDefault == ulImage)
	{
		LPSPropValue lpObjType = NULL;
		lpObjType = PpropFindProp(lpsRowData->lpProps,lpsRowData->cValues,PR_OBJECT_TYPE);
		if (lpObjType && PR_OBJECT_TYPE == lpObjType->ulPropTag)
		{
			int i = 0;
			for (i=0;i<NUMOBJTYPES;i++)
			{
				if (_ObjTypeIcons[i][0] == lpObjType->Value.l)
				{
					ulImage = _ObjTypeIcons[i][1];
					break;
				}
			}

		}
	}

	if (lpulDepth) *lpulDepth = ulDepth;
	if (lpulImage) *lpulImage = ulImage;
}
HRESULT CContentsTableListCtrl::RefreshItem(int iRow,LPSRow lpsRowData, BOOL bItemExists)
{
	HRESULT			hRes = S_OK;
	SortListData*	lpData = 0;

	DebugPrintEx(DBGGeneric,CLASS,_T("RefreshItem"),_T("item %d\n"),iRow);

	if (bItemExists)
	{
		lpData = (SortListData*) GetItemData(iRow);
	}
	else
	{
		ULONG ulDepth = NULL;
		ULONG ulImage = slIconDefault;
		GetDepthAndImage(lpsRowData,&ulDepth,&ulImage);

		lpData = InsertRow(iRow, _T("TempRefreshItem"),ulDepth,ulImage); // STRING_OK
	}

	if (lpData)
	{
		BuildDataItem(lpsRowData,lpData);

		SetRowStrings(iRow,lpsRowData);
		// Do this last so that our row can't get sorted before we're done!
		lpData->bItemFullyLoaded = true;
	}

	return hRes;
} // CContentsTableListCtrl::RefreshItem

// Crack open the given SPropValue and render it to the given row in the list.
HRESULT CContentsTableListCtrl::AddItemToListBox(int iRow, LPSRow lpsRowToAdd)
{
	HRESULT			hRes = S_OK;

	DebugPrintEx(DBGGeneric,CLASS,_T("AddItemToListBox"),_T("item %d\n"),iRow);

	EC_H(RefreshItem(iRow,lpsRowToAdd,false));

	if (m_lpHostDlg)
		m_lpHostDlg->UpdateStatusBarText(STATUSLEFTPANE,IDS_STATUSTEXTNUMITEMS,GetItemCount());

	return hRes;
} // CContentsTableListCtrl::AddItemToListBox

void CContentsTableListCtrl::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	DebugPrintEx(DBGMenu,CLASS,_T("OnKeyDown"),_T("0x%X\n"),nChar);

	if (!m_lpHostDlg) return;
	ULONG bCtrlPressed = GetKeyState(VK_CONTROL) < 0;
	ULONG bShiftPressed = GetKeyState(VK_SHIFT) < 0;
	ULONG bMenuPressed = GetKeyState(VK_MENU) < 0;

	if (!bMenuPressed)
	{
		if ('A' == nChar && bCtrlPressed)
		{
			SelectAll();
		}
		else if (VK_RETURN == nChar && S_FALSE != DoExpandCollapse())
		{
			// nothing to do - DoExpandCollapse did the work
			// we only want to go to the next case if it returned S_FALSE
		}
		else if (!m_lpHostDlg || !m_lpHostDlg->HandleKeyDown(nChar,bShiftPressed,bCtrlPressed,bMenuPressed))
		{
			CSortListCtrl::OnKeyDown(nChar,nRepCnt,nFlags);
		}
	}
}

HRESULT CContentsTableListCtrl::GetSelectedItemEIDs(LPENTRYLIST* lppEntryIDs)
{
	*lppEntryIDs = NULL;
	HRESULT hRes = S_OK;
	int iNumItems = GetSelectedCount();

	if (!iNumItems) return S_OK;
	if (iNumItems > ULONG_MAX/sizeof(SBinary)) return MAPI_E_INVALID_PARAMETER;

	LPENTRYLIST lpTempList = NULL;

	EC_H(MAPIAllocateBuffer(sizeof(ENTRYLIST),(LPVOID*) &lpTempList));

	if (lpTempList)
	{
		lpTempList->cValues = iNumItems;
		lpTempList->lpbin = NULL;

		EC_H(MAPIAllocateMore(
			(ULONG)sizeof(SBinary) * iNumItems,
			lpTempList,
			(LPVOID*) &lpTempList->lpbin));
		if (lpTempList->lpbin)
		{
			int iArrayPos = 0;
			int iSelectedItem = -1;

			for(iArrayPos = 0 ; iArrayPos < iNumItems ; iArrayPos++)
			{
				lpTempList->lpbin[iArrayPos].cb = 0;
				lpTempList->lpbin[iArrayPos].lpb = NULL;
				iSelectedItem = GetNextItem(
					iSelectedItem,
					LVNI_SELECTED);
				if (-1 != iSelectedItem)
				{
					SortListData* lpData = (SortListData*) GetItemData(iSelectedItem);
					if (lpData && lpData->data.Contents.lpEntryID)
					{
						lpTempList->lpbin[iArrayPos].cb = lpData->data.Contents.lpEntryID->cb;
						EC_H(MAPIAllocateMore(
								lpData->data.Contents.lpEntryID->cb,
								lpTempList,
								(LPVOID *) &lpTempList->lpbin[iArrayPos].lpb));
						if (lpTempList->lpbin[iArrayPos].lpb)
						{
							CopyMemory(
								lpTempList->lpbin[iArrayPos].lpb,
								lpData->data.Contents.lpEntryID->lpb,
								lpData->data.Contents.lpEntryID->cb);
						}
					}
				}
			}
		}
	}

	*lppEntryIDs = lpTempList;
	return hRes;
} // CContentsTableListCtrl::GetSelectedItemEIDs

// Pass iCurItem as -1 to get the primary selected item.
// Call again with the previous iCurItem to get the next one.
// Stop calling when iCurItem = -1 and/or lppProp is NULL
// If iCurItem is NULL, just returns the focused item
int	CContentsTableListCtrl::GetNextSelectedItemNum(
												   int *iCurItem)
{
	int	iItem = NULL;

	if (iCurItem) // intentionally not dereffing - checking to see if NULL was actually passed
	{
		iItem = *iCurItem;
	}
	else
	{
		iItem = -1;
	}
	DebugPrintEx(DBGGeneric,CLASS,_T("GetNextSelectedItemNum"),_T("iItem before = 0x%X\n"),iItem);

	iItem = GetNextItem(
		iItem,
		LVNI_SELECTED);

	DebugPrintEx(DBGGeneric,CLASS,_T("GetNextSelectedItemNum"),_T("iItem after = 0x%X\n"),iItem);

	if (iCurItem) *iCurItem = iItem;

	return iItem;
}

SortListData* CContentsTableListCtrl::GetNextSelectedItemData(int *iCurItem)
{
	int iItem;

	iItem = GetNextSelectedItemNum(iCurItem);
	if (-1 == iItem) return NULL;
	return (SortListData*) GetItemData(iItem);
}

// Pass iCurItem as -1 to get the primary selected item.
// Call again with the previous iCurItem to get the next one.
// Stop calling when iCurItem = -1 and/or lppProp is NULL
// If iCurItem is NULL, just returns the focused item
HRESULT	CContentsTableListCtrl::OpenNextSelectedItemProp(
													int *iCurItem,
													__mfcmapiModifyEnum bModify,
													LPMAPIPROP* lppProp)
{
	HRESULT	hRes = S_OK;
	int iItem;

	*lppProp = NULL;

	iItem = GetNextSelectedItemNum(iCurItem);
	if (-1 != iItem)
		WC_H(m_lpHostDlg->OpenItemProp(iItem, bModify, lppProp));

	return hRes;
} // CContentsTableListCtrl::OpenNextSelectedItemProp

HRESULT CContentsTableListCtrl::DefaultOpenItemProp(
							   int iItem,
							   __mfcmapiModifyEnum bModify,
							   LPMAPIPROP* lppProp)
{
	HRESULT			hRes = S_OK;
	LPSBinary		lpEID = NULL;

	*lppProp = NULL;

	if (!m_lpMapiObjects || -1 == iItem) return S_OK;

	DebugPrintEx(DBGGeneric,CLASS,_T("DefaultOpenItemProp"),_T("iItem = %d, bModify = %d, m_ulContainerType = 0x%X\n"),iItem, bModify, m_ulContainerType);

	SortListData* lpListData = NULL;
	lpListData = (SortListData*) GetItemData(iItem);

	if (!lpListData) return S_OK;
	lpEID = lpListData->data.Contents.lpEntryID;
	if (!lpEID || (lpEID->cb == 0)) return S_OK;

	DebugPrint(DBGGeneric,_T("Item being opened:\n"));
	DebugPrintBinary(DBGGeneric,lpEID);

	// Find the highlighted item EID
	switch(m_ulContainerType)
	{
	case (MAPI_ABCONT):
		{
			LPADRBOOK lpAB = m_lpMapiObjects->GetAddrBook(false); // do not release
			WC_H(CallOpenEntry(
				NULL,
				lpAB, // use AB
				NULL,
				NULL,
				lpEID,
				NULL,
				(bModify == mfcmapiREQUEST_MODIFY)?MAPI_MODIFY:MAPI_BEST_ACCESS,
				NULL,
				(LPUNKNOWN*)lppProp));
		}
		break;
	case(MAPI_FOLDER):
		{
			LPMDB lpMDB = m_lpMapiObjects->GetMDB(); // do not release
			LPCIID lpInterface = NULL;

			if (RegKeys[regkeyUSE_MESSAGERAW].ulCurDWORD)
			{
				lpInterface = &IID_IMessageRaw;
			}

			WC_H(CallOpenEntry(
				lpMDB, // use MDB
				NULL,
				NULL,
				NULL,
				lpEID,
				lpInterface,
				(bModify == mfcmapiREQUEST_MODIFY)?MAPI_MODIFY:MAPI_BEST_ACCESS,
				NULL,
				(LPUNKNOWN*)lppProp));
			if (MAPI_E_INTERFACE_NOT_SUPPORTED == hRes && RegKeys[regkeyUSE_MESSAGERAW].ulCurDWORD)
			{
				ErrDialog(__FILE__,__LINE__,IDS_EDMESSAGERAWNOTSUPPORTED);
			}
		}
		break;
	default:
		{
			LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
			WC_H(CallOpenEntry(
				NULL,
				NULL,
				NULL,
				lpMAPISession, // use session
				lpEID,
				NULL,
				(bModify == mfcmapiREQUEST_MODIFY)?MAPI_MODIFY:MAPI_BEST_ACCESS,
				NULL,
				(LPUNKNOWN*)lppProp));
		}
		break;
	}
	if (!*lppProp && FAILED(hRes) && mfcmapiREQUEST_MODIFY == bModify && MAPI_E_NOT_FOUND != hRes)
	{
		DebugPrint(DBGGeneric,_T("\tOpenEntry failed: 0x%X. Will try again without MAPI_MODIFY\n"),hRes);
		// We got access denied when we passed MAPI_MODIFY
		// Let's try again without it.
		hRes = S_OK;
		EC_H(DefaultOpenItemProp(
			iItem,
			mfcmapiDO_NOT_REQUEST_MODIFY,
			lppProp));
	}

	if (MAPI_E_NOT_FOUND == hRes)
	{
		DebugPrint(DBGGeneric,_T("\tDefaultOpenItemProp encountered an entry ID for an item that doesn't exist\n\tThis happens often when we're deleting items.\n"));
		hRes = S_OK;
	}

	DebugPrintEx(DBGGeneric,CLASS,_T("DefaultOpenItemProp"),_T("returning *lppProp = 0x%X and hRes = 0x%X\n"),*lppProp, hRes);
	return hRes;
} // CContentsTableListCtrl::DefaultOpenItemProp

void CContentsTableListCtrl::SelectAll()
{
	HRESULT hRes = S_OK;
	int iIndex = 0;
	DebugPrintEx(DBGGeneric,CLASS,_T("SelectAll"),_T("\n"));
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.
	MySetRedraw(FALSE);
	for (iIndex = 0;iIndex < GetItemCount();iIndex++)
	{
		EC_B(SetItemState(iIndex,LVIS_SELECTED,LVIS_SELECTED | LVIS_FOCUSED));
		hRes = S_OK;
	}
	MySetRedraw(TRUE);
	if (m_lpHostDlg)
		m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(NULL, NULL);
}

// This is a tough function. I wonder if I'm handling this event correctly
void CContentsTableListCtrl::OnItemChanged(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMLISTVIEW pNMListView = (LPNMLISTVIEW) pNMHDR;

	*pResult = 0;

	if (!pNMListView || !(pNMListView->uChanged & LVIF_STATE)) return;
	// We get spurious ItemChanged events while scrolling with the keyboard. Ignore them.
	if (GetKeyState(VK_RIGHT) < 0 || GetKeyState(VK_LEFT) < 0) return;

	// Keep all our logic in here
	if ((pNMListView->uNewState & LVIS_FOCUSED) && m_lpHostDlg)
	{
		LPMAPIPROP		lpMAPIProp = NULL;
		SortListData*	lpData = 0;
		CString			szTitle;
		if (1 == GetSelectedCount())
		{
			HRESULT			hRes = S_OK;
			LPSPropValue	lpProp = NULL;

			// go get the original row for display in the prop list control
			lpData = (SortListData*) GetItemData(pNMListView->iItem);
			ULONG			cValues = 0;
			LPSPropValue	lpProps = NULL;
			if (lpData)
			{
				cValues = lpData->cSourceProps;
				lpProps = lpData->lpSourceProps;
			}

			EC_H(m_lpHostDlg->OpenItemProp(pNMListView->iItem, mfcmapiREQUEST_MODIFY, &lpMAPIProp));

			szTitle.LoadString(IDS_DISPLAYNAMENOTFOUND);

			// try to use our rowset first
			if (NODISPLAYNAME != m_ulDisplayNameColumn
				&& lpProps
				&& m_ulDisplayNameColumn < cValues)
			{
				if (CheckStringProp(&lpProps[m_ulDisplayNameColumn],PT_STRING8))
				{
					szTitle.Format(_T("%hs"),lpProps[m_ulDisplayNameColumn].Value.lpszA); // STRING_OK
				}
				else if (CheckStringProp(&lpProps[m_ulDisplayNameColumn],PT_UNICODE))
				{
					szTitle.Format(_T("%ws"),lpProps[m_ulDisplayNameColumn].Value.lpszW); // STRING_OK
				}
			}
			else if (lpMAPIProp)
			{
				// Get a property from the item for the title bar
				hRes = HrGetOneProp(
					lpMAPIProp,
					PR_DISPLAY_NAME,
					&lpProp);
				if (MAPI_E_NOT_FOUND == hRes)
				{
					hRes = S_OK;
					// Let's try a different property
					WC_H(HrGetOneProp(
						lpMAPIProp,
						PR_SUBJECT,
						&lpProp));
				}
				else WARNHRESMSG(hRes,IDS_TITLEPROPNOTFOUND);
				if (lpProp)
				{
					if (CheckStringProp(lpProp,PT_TSTRING))
					{
						szTitle = lpProp->Value.LPSZ;
					}
					MAPIFreeBuffer(lpProp);
				}
			}
		}

		// Update the main window with our changes
		m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(lpMAPIProp, lpData);
		m_lpHostDlg->UpdateTitleBarText(szTitle);

		if (lpMAPIProp) lpMAPIProp->Release();
	}
}

BOOL CContentsTableListCtrl::IsAdviseSet()
{
	return m_lpAdviseSink?true:false;
} // CContentsTableListCtrl::IsAdviseSet

HRESULT CContentsTableListCtrl::NotificationOn()
{
	HRESULT			hRes = S_OK;

	if (m_lpAdviseSink) return S_OK;
	if (!m_lpContentsTable) return MAPI_E_INVALID_PARAMETER;

	DebugPrintEx(DBGGeneric,CLASS,_T("NotificationOn"),_T("registering table notification on 0x%X\n"),m_lpContentsTable);

	m_lpAdviseSink =  new CAdviseSink(m_hWnd,NULL);

	if (m_lpAdviseSink)
	{
		WC_H(m_lpContentsTable->Advise(
			fnevTableModified,
			(IMAPIAdviseSink *)m_lpAdviseSink,
			&m_ulAdviseConnection));
		if (MAPI_E_NO_SUPPORT == hRes) // Some tables don't support this!
		{
			if (m_lpAdviseSink) m_lpAdviseSink->Release();
			m_lpAdviseSink = NULL;
			DebugPrint(DBGGeneric, _T("This table doesn't support notifications\n"));
			hRes = S_OK; // mask the error
		}
		else if (S_OK == hRes)
		{
			LPSPropValue	lpProp = NULL;

			LPMDB lpMDB = m_lpMapiObjects->GetMDB(); // do not release
			if (lpMDB)
			{
				// Try to trigger some RPC to get the notifications going
				WC_H(HrGetOneProp(
					lpMDB,
					PR_TEST_LINE_SPEED,
					&lpProp));
				if (MAPI_E_NOT_FOUND == hRes)
				{
					// We're not on an Exchange server. We don't need to generate RPC after all.
					hRes = S_OK;
				}
				MAPIFreeBuffer(lpProp);
			}
		}
	}

	DebugPrintEx(DBGGeneric,CLASS,_T("NotificationOn"),_T("Table notification results (Sink:0x%X, ulConnection:0x%X) on 0x%X\n"),m_lpAdviseSink, m_ulAdviseConnection, m_lpContentsTable);
	return hRes;
}

// This function gets called a lot, make sure it's ok to call it too often...:)
// If there exists a current advise sink, unadvise it. Otherwise, don't complain.
void CContentsTableListCtrl::NotificationOff()
{
	if (!m_lpAdviseSink) return;
	DebugPrintEx(DBGGeneric,CLASS,_T("NotificationOff"),_T("clearing table notification (Sink:0x%X, ulConnection:0x%X) on 0x%X\n"),m_lpAdviseSink, m_ulAdviseConnection, m_lpContentsTable);

	if (m_ulAdviseConnection && m_lpContentsTable)
		m_lpContentsTable->Unadvise(m_ulAdviseConnection);

	m_ulAdviseConnection = NULL;
	if (m_lpAdviseSink) m_lpAdviseSink->Release();
	m_lpAdviseSink = NULL;
}


HRESULT CContentsTableListCtrl::RefreshTable()
{
	HRESULT hRes = S_OK;
	if (!m_lpHostDlg) return MAPI_E_INVALID_PARAMETER;
	if (m_bInLoadOp)
	{
		DebugPrintEx(DBGGeneric,CLASS,_T("RefreshTable"),_T("called during table load - ditching call\n"));
		return S_OK;
	}

	DebugPrintEx(DBGGeneric,CLASS,_T("RefreshTable"),_T("\n"));

	EC_H(LoadContentsTableIntoView());

	// Reset the title while we're at it
	m_lpHostDlg->UpdateTitleBarText(NULL);

	return hRes;
}

// Call ExpandRow or CollapseRow as needed
// Returns
// S_OK - made a call to ExpandRow/CollapseRow and it succeeded
// S_FALSE - no call was needed
// other errors as appropriate
HRESULT CContentsTableListCtrl::DoExpandCollapse()
{
	HRESULT			hRes = S_FALSE;
	SortListData*	lpData = 0;

	int iItem = GetNextItem(
		-1,
		LVNI_SELECTED);

	// nothing selected, no work done
	if (-1 == iItem) return S_FALSE;

	lpData = (SortListData*) GetItemData(iItem);

	// No lpData or wrong type of row - no work done
	if (!lpData ||
		lpData->ulSortDataType != SORTLIST_CONTENTS ||
		lpData->data.Contents.ulRowType == TBL_LEAF_ROW ||
		lpData->data.Contents.ulRowType == TBL_EMPTY_CATEGORY)
		return S_FALSE;

	BOOL bDidWork = false;
	LVITEM lvItem = {0};
	lvItem.iItem = iItem;
	lvItem.iSubItem = 0;
	lvItem.mask = LVIF_IMAGE;
	switch(lpData->data.Contents.ulRowType)
	{
	default:
		break;
	case TBL_COLLAPSED_CATEGORY:
		{
			if (lpData->data.Contents.lpInstanceKey)
			{
				LPSRowSet lpRowSet = NULL;
				ULONG ulRowsAdded = 0;

				EC_H(m_lpContentsTable->ExpandRow(
					lpData->data.Contents.lpInstanceKey->cb,
					lpData->data.Contents.lpInstanceKey->lpb,
					256,
					NULL,
					&lpRowSet,
					&ulRowsAdded));
				if (S_OK == hRes && lpRowSet)
				{
					ULONG i = 0;
					for (i = 0 ; i < lpRowSet->cRows ; i++)
					{
						// add the item to the NEXT slot
						EC_H(AddItemToListBox(iItem + i + 1,&lpRowSet->aRow[i]));
					}
				}
				MAPIFreeBuffer(lpRowSet);
				lpData->data.Contents.ulRowType = TBL_EXPANDED_CATEGORY;
				lvItem.iImage = slIconNodeExpanded;
				bDidWork = true;
			}
		}
		break;
	case TBL_EXPANDED_CATEGORY:
		if (lpData->data.Contents.lpInstanceKey)
		{
			ULONG ulRowsRemoved = 0;

			EC_H(m_lpContentsTable->CollapseRow(
				lpData->data.Contents.lpInstanceKey->cb,
				lpData->data.Contents.lpInstanceKey->lpb,
				NULL,
				&ulRowsRemoved));
			if (S_OK == hRes && ulRowsRemoved)
			{
				int i = 0;
				for (i = iItem + ulRowsRemoved ; i > iItem ; i--)
				{
					EC_B(DeleteItem(i));
				}
			}
			lpData->data.Contents.ulRowType = TBL_COLLAPSED_CATEGORY;
			lvItem.iImage = slIconNodeCollapsed;
			bDidWork = true;
		}
		break;
	}
	if (bDidWork)
	{
		EC_B(SetItem(&lvItem)); // Set new image for the row
		LPSPropValue lpProp = NULL;

		// Save the row type (header/leaf) into lpData
		lpProp = PpropFindProp(
			lpData->lpSourceProps,
			lpData->cSourceProps,
			PR_ROW_TYPE);
		if (lpProp && PR_ROW_TYPE == lpProp->ulPropTag)
		{
			lpProp->Value.l = lpData->data.Contents.ulRowType;
		}
		SRow sRowData = {0};
		sRowData.cValues = lpData->cSourceProps;
		sRowData.lpProps = lpData->lpSourceProps;
		SetRowStrings(iItem,&sRowData);
	}
	return hRes;
}

void CContentsTableListCtrl::OnOutputTable(LPCWSTR szFileName)
{
	if (m_bInLoadOp) return;
	FILE* fTable = NULL;
	fTable = OpenFile(szFileName,true);
	if (fTable)
	{
		OutputTableToFile(fTable,m_lpContentsTable);
		CloseFile(fTable);
	}
} // CContentsTableListCtrl::OnOutputTable

HRESULT CContentsTableListCtrl::SetSortTable(LPSSortOrderSet lpSortOrderSet, ULONG ulFlags)
{
	HRESULT hRes = S_OK;
	if (!m_lpContentsTable) return MAPI_E_INVALID_PARAMETER;

	EC_H(m_lpContentsTable->SortTable(
		lpSortOrderSet,
		ulFlags));

	return hRes;
} // CContentsTableListCtrl::SetSortOrder

// WM_MFCMAPI_THREADADDITEM
LRESULT	CContentsTableListCtrl::msgOnThreadAddItem(WPARAM wParam, LPARAM lParam)
{
	HRESULT hRes = S_OK;
	int		iNewRow = (int) wParam;
	LPSRow	lpsRow = (LPSRow) lParam;

	if (!lpsRow) return MAPI_E_INVALID_PARAMETER;

	DebugPrintEx(DBGGeneric,CLASS,_T("msgOnThreadAddItem"),_T("Received message to add 0x%08X to row %d\n"),lpsRow,iNewRow);
	EC_H(AddItemToListBox(iNewRow,lpsRow));

	return hRes;
}

// WM_MFCMAPI_ADDITEM
LRESULT	CContentsTableListCtrl::msgOnAddItem(WPARAM wParam, LPARAM /*lParam*/)
{
	HRESULT hRes = S_OK;
	TABLE_NOTIFICATION*	tab = (TABLE_NOTIFICATION*) wParam;

	if (!tab) return MAPI_E_INVALID_PARAMETER;

	// If a row is added, propPrior will contain information about the row preceding the
	// added row. If propPrior.ulPropTag is NULL, then the new item goes first on the list
	int iNewRow = 0;
	if (PR_NULL == tab->propPrior.ulPropTag)
	{
		iNewRow = 0;
	}
	else
	{
		iNewRow = FindRow(&tab->propPrior.Value.bin) + 1;
	}

	// We make this copy here and pass it in to AddItemToListBox, where it is grabbed by BuildDataItem to be part of the item data
	// The mem will be freed when the item data is cleaned up - do not free here
	SRow NewRow = {0};
	NewRow.cValues = tab->row.cValues;
	NewRow.ulAdrEntryPad = tab->row.ulAdrEntryPad;
	EC_H(ScDupPropset(
		tab->row.cValues,
		tab->row.lpProps,
		MAPIAllocateBuffer,
		&NewRow.lpProps));

	DebugPrintEx(DBGGeneric,CLASS,_T("msgOnAddItem"),_T("Received message to add 0x%08X to row %d\n"),tab->row,iNewRow);
	EC_H(AddItemToListBox(iNewRow,&NewRow));

	return hRes;
}

// WM_MFCMAPI_DELETEITEM
LRESULT	CContentsTableListCtrl::msgOnDeleteItem(WPARAM wParam, LPARAM /*lParam*/)
{
	HRESULT hRes = S_OK;
	TABLE_NOTIFICATION*	tab = (TABLE_NOTIFICATION*) wParam;

	if (!tab) return MAPI_E_INVALID_PARAMETER;

	int	iItem = FindRow(&tab->propIndex.Value.bin);

	DebugPrintEx(DBGGeneric,CLASS,_T("msgOnDeleteItem"),_T("Received message to delete item 0x%d\n"),iItem);

	if (iItem == -1) return S_OK;

	EC_B(DeleteItem(iItem));

	if (S_OK != hRes || !m_lpHostDlg) return hRes;

	int iCount = GetItemCount();
	if (iCount == 0)
	{
		m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(NULL, NULL);
	}

	m_lpHostDlg->UpdateStatusBarText(STATUSLEFTPANE,IDS_STATUSTEXTNUMITEMS,iCount);
	return hRes;
}

// WM_MFCMAPI_MODIFYITEM
LRESULT	CContentsTableListCtrl::msgOnModifyItem(WPARAM wParam, LPARAM /*lParam*/)
{
	HRESULT hRes = S_OK;
	TABLE_NOTIFICATION*	tab = (TABLE_NOTIFICATION*) wParam;

	if (!tab) return MAPI_E_INVALID_PARAMETER;

	int	iItem = FindRow(&tab->propIndex.Value.bin);

	if (-1 != iItem)
	{
		DebugPrintEx(DBGGeneric,CLASS,_T("msgOnModifyItem"),_T("Received message to modify row %d with 0x%08X\n"),iItem,&tab->row);

		// We make this copy here and pass it in to RefreshItem, where it is grabbed by BuildDataItem to be part of the item data
		// The mem will be freed when the item data is cleaned up - do not free here
		SRow NewRow = {0};
		NewRow.cValues = tab->row.cValues;
		NewRow.ulAdrEntryPad = tab->row.ulAdrEntryPad;
		EC_H(ScDupPropset(
			tab->row.cValues,
			tab->row.lpProps,
			MAPIAllocateBuffer,
			&NewRow.lpProps));

		EC_H(RefreshItem(iItem,&NewRow,TRUE));
	}

	return hRes;
}

// WM_MFCMAPI_REFRESHTABLE
LRESULT	CContentsTableListCtrl::msgOnRefreshTable(WPARAM /*wParam*/, LPARAM /*lParam*/)
{
	HRESULT hRes = S_OK;
	DebugPrintEx(DBGGeneric,CLASS,_T("msgOnRefreshTable"),_T("Received message refresh table\n"));
	EC_H(RefreshTable());

	return hRes;
}

// This function steps through the list control to find the entry with this instance key
// return -1 if item not found
int	CContentsTableListCtrl::FindRow(LPSBinary lpInstance)
{
	int iItem = 0;

	LPSBinary lpCurInstance = NULL;
	SortListData* lpListData = NULL;

	DebugPrintEx(DBGGeneric,CLASS,_T("msgOnGetIndex"),_T("Getting index for 0x%08X\n"),lpInstance);

	if (!lpInstance) return -1;

	for (iItem = 0;iItem<GetItemCount();iItem++)
	{
		lpListData = (SortListData*) GetItemData(iItem);
		if (lpListData)
		{
			lpCurInstance = lpListData->data.Contents.lpInstanceKey;
			if (lpCurInstance)
			{
				if (!memcmp(lpCurInstance->lpb, lpInstance->lpb, lpInstance->cb))
				{
					DebugPrintEx(DBGGeneric,CLASS,_T("msgOnGetIndex"),_T("Matched at 0x%08X\n"),iItem);
					return iItem;
				}
			}
		}
	}

	DebugPrintEx(DBGGeneric,CLASS,_T("msgOnGetIndex"),_T("No match found: 0x%08X\n"),iItem);
	return -1;
}