// MsgStoreDlg.cpp : implementation file
// Displays the hierarchy tree of folders in a message store

#include "stdafx.h"
#include "Error.h"

#include "MsgStoreDlg.h"

#include "HierarchyTableTreeCtrl.h"
#include "MapiObjects.h"
#include "MAPIFunctions.h"
#include "MAPIStoreFunctions.h"
#include "SingleMAPIPropListCtrl.h"
#include "MFCUtilityFunctions.h"
#include "Editor.h"
#include "FolderDlg.h"
#include "MailboxTableDlg.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"

#include "Dumpstore.h"
#include "File.h"
#include "FileDialogEx.h"
#include "ExtraPropTags.h"
#include "MAPIProgress.h"
#include "FormContainerDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

static TCHAR* CLASS = _T("CMsgStoreDlg");

/////////////////////////////////////////////////////////////////////////////
// CMsgStoreDlg dialog


CMsgStoreDlg::CMsgStoreDlg(
					   CParentWnd* pParentWnd,
					   CMapiObjects *lpMapiObjects,
					   LPMAPIFOLDER lpRootFolder,
					   __mfcmapiDeletedItemsEnum bShowingDeletedFolders
					   ):
CHierarchyTableDlg(
						   pParentWnd,
						   lpMapiObjects,
						   IDS_FOLDERTREE,
						   lpRootFolder,
						   IDR_MENU_MESSAGESTORE_POPUP,
						   MENU_CONTEXT_FOLDER_TREE)
{
	TRACE_CONSTRUCTOR(CLASS);
	HRESULT			hRes = S_OK;
	LPSPropValue	lpProp = NULL;

	m_bShowingDeletedFolders = bShowingDeletedFolders;

	if (m_lpMapiObjects)
	{
		LPMDB lpMDB = m_lpMapiObjects->GetMDB();//do not release
		if (lpMDB)
		{
			WC_H(HrGetOneProp(
				lpMDB,
				PR_DISPLAY_NAME,
				&lpProp));
			if (lpProp)
			{
				//Check for a NULL value and get a different prop if needed
				if (_T('\0') == lpProp->Value.LPSZ[0])
				{
					MAPIFreeBuffer(lpProp);

					WC_H(HrGetOneProp(
						lpMDB,
						PR_MAILBOX_OWNER_NAME,
						&lpProp));
				}
			}
			hRes = S_OK;

			//Set the title
			// Shouldn't have to check lpProp for non-NULL since CheckString does it, but prefast is complaining
			if (lpProp && CheckStringProp(lpProp,PT_TSTRING))
			{
				m_szTitle = lpProp->Value.LPSZ;
			}
			MAPIFreeBuffer(lpProp);

			if (!m_lpContainer)
			{
				// Open root container.
				EC_H(CallOpenEntry(
					lpMDB,
					NULL,
					NULL,
					NULL,
					NULL,//open root container
					NULL,
					MAPI_BEST_ACCESS,
					NULL,
					(LPUNKNOWN*)&m_lpContainer));
			}
		}
	}

	CreateDialogAndMenu(IDR_MENU_MESSAGESTORE);
}

CMsgStoreDlg::~CMsgStoreDlg()
{
	TRACE_DESTRUCTOR(CLASS);
}

BEGIN_MESSAGE_MAP(CMsgStoreDlg, CHierarchyTableDlg)
//{{AFX_MSG_MAP(CMsgStoreDlg)
	ON_COMMAND(ID_PASTE_RULES, OnPasteRules)
	ON_COMMAND(ID_CREATESUBFOLDER, OnCreateSubFolder)
	ON_COMMAND(ID_DISPLAYASSOCIATEDCONTENTS, OnDisplayAssociatedContents)
	ON_COMMAND(ID_DISPLAYCALENDAR, OnDisplayCalendarFolder)
	ON_COMMAND(ID_DISPLAYCONTACTS, OnDisplayContactsFolder)
	ON_COMMAND(ID_DISPLAYDELETEDCONTENTS, OnDisplayDeletedContents)
	ON_COMMAND(ID_DISPLAYDELETEDSUBFOLDERS, OnDisplayDeletedSubFolders)
	ON_COMMAND(ID_DISPLAYINBOX, OnDisplayInbox)
	ON_COMMAND(ID_DISPLAYMAILBOXTABLE, OnDisplayMailboxTable)
	ON_COMMAND(ID_DISPLAYOUTGOINGQUEUE, OnDisplayOutgoingQueueTable)
	ON_COMMAND(ID_DISPLAYRECEIVEFOLDERTABLE, OnDisplayReceiveFolderTable)
	ON_COMMAND(ID_DISPLAYRULESTABLE, OnDisplayRulesTable)
	ON_COMMAND(ID_DISPLAYACLTABLE, OnDisplayACLTable)
	ON_COMMAND(ID_DISPLAYTASKS, OnDisplayTasksFolder)
	ON_COMMAND(ID_EMPTYFOLDER, OnEmptyFolder)
	ON_COMMAND(ID_DELETESELECTEDITEM, OnDeleteSelectedItem)
	ON_COMMAND(ID_OPENFORMCONTAINER, OnOpenFormContainer)
	ON_COMMAND(ID_SELECTFORM, OnSelectForm)
	ON_COMMAND(ID_RESENDALLMESSAGES, OnResendAllMessages)
	ON_COMMAND(ID_RESETPERMISSIONSONITEMS, OnResetPermissionsOnItems)
	ON_COMMAND(ID_RESTOREDELETEDFOLDER, OnRestoreDeletedFolder)
	ON_COMMAND(ID_SAVEFOLDERCONTENTSASMSG, OnSaveFolderContentsAsMSG)
	ON_COMMAND(ID_SAVEFOLDERCONTENTSASTEXTFILES, OnSaveFolderContentsAsTextFiles)
	ON_COMMAND(ID_SETRECEIVEFOLDER, OnSetReceiveFolder)

	ON_COMMAND(ID_VALIDATEIPMSUBTREE, OnValidateIPMSubtree)
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

void CMsgStoreDlg::OnInitMenu(CMenu* pMenu)
{
	if (!pMenu) return;

	LPMDB	lpMDB = NULL;
	BOOL	bItemSelected = m_lpHierarchyTableTreeCtrl && m_lpHierarchyTableTreeCtrl->m_bItemSelected;
	if (m_lpMapiObjects)
		lpMDB = m_lpMapiObjects->GetMDB();//do not release

	if (m_lpMapiObjects)
	{
		ULONG ulStatus = m_lpMapiObjects->GetBufferStatus();
		pMenu->EnableMenuItem(ID_PASTE,DIM((ulStatus != BUFFER_EMPTY) && bItemSelected));
		pMenu->EnableMenuItem(ID_PASTE_RULES,DIM((ulStatus & BUFFER_FOLDER) && bItemSelected));
	}
	pMenu->EnableMenuItem(ID_DISPLAYASSOCIATEDCONTENTS,DIM(bItemSelected));
	pMenu->EnableMenuItem(ID_DISPLAYDELETEDCONTENTS,DIM(bItemSelected));
	pMenu->EnableMenuItem(ID_DISPLAYDELETEDSUBFOLDERS,DIM(bItemSelected));

	pMenu->EnableMenuItem(ID_CREATESUBFOLDER,DIM(bItemSelected));
	pMenu->EnableMenuItem(ID_DISPLAYACLTABLE,DIM(bItemSelected));
	pMenu->EnableMenuItem(ID_DISPLAYMAILBOXTABLE,DIM(lpMDB));
	pMenu->EnableMenuItem(ID_DISPLAYOUTGOINGQUEUE,DIM(lpMDB));
	pMenu->EnableMenuItem(ID_DISPLAYRECEIVEFOLDERTABLE,DIM(lpMDB));
	pMenu->EnableMenuItem(ID_DISPLAYRULESTABLE,DIM(bItemSelected));
	pMenu->EnableMenuItem(ID_EMPTYFOLDER,DIM(bItemSelected));
	pMenu->EnableMenuItem(ID_OPENFORMCONTAINER,DIM(bItemSelected));
	pMenu->EnableMenuItem(ID_SELECTFORM,DIM(bItemSelected));
	pMenu->EnableMenuItem(ID_SAVEFOLDERCONTENTSASMSG,DIM(bItemSelected));
	pMenu->EnableMenuItem(ID_SAVEFOLDERCONTENTSASTEXTFILES,DIM(bItemSelected));
	pMenu->EnableMenuItem(ID_SETRECEIVEFOLDER,DIM(bItemSelected));

	pMenu->EnableMenuItem(ID_RESENDALLMESSAGES,DIM(bItemSelected));
	pMenu->EnableMenuItem(ID_RESETPERMISSIONSONITEMS,DIM(bItemSelected));
	pMenu->EnableMenuItem(ID_RESTOREDELETEDFOLDER,DIM(bItemSelected && mfcmapiSHOW_DELETED_ITEMS == m_bShowingDeletedFolders));

	pMenu->EnableMenuItem(ID_COPY,DIM(bItemSelected));
	pMenu->EnableMenuItem(ID_DELETESELECTEDITEM,DIM(bItemSelected));

	pMenu->EnableMenuItem(ID_DISPLAYINBOX,DIM(lpMDB));
	pMenu->EnableMenuItem(ID_DISPLAYCALENDAR,DIM(lpMDB));
	pMenu->EnableMenuItem(ID_DISPLAYCONTACTS,DIM(lpMDB));
	pMenu->EnableMenuItem(ID_DISPLAYTASKS,DIM(lpMDB));

	CHierarchyTableDlg::OnInitMenu(pMenu);
}

/////////////////////////////////////////////////////////////////////////////////////
//  Menu Commands

void CMsgStoreDlg::OnDisplayInbox()
{
	HRESULT			hRes = S_OK;
	LPMAPIFOLDER	lpInbox = NULL;

	if (!m_lpMapiObjects) return;

	LPMDB lpMDB = m_lpMapiObjects->GetMDB();//do not release
	if (!lpMDB) return;

	EC_H(GetInbox(lpMDB,&lpInbox));

	if (lpInbox)
	{
		EC_H(DisplayObject(
			lpInbox,
			NULL,
			otHierarchy,
			this));

		lpInbox->Release();
	}
}//CMsgStoreDlg::OnDisplayInbox

void CMsgStoreDlg::OnDisplaySpecialFolder(ULONG ulPropTag)
{
	HRESULT			hRes = S_OK;
	LPMAPIFOLDER	lpFolder = NULL;

	if (!m_lpMapiObjects) return;

	LPMDB lpMDB = m_lpMapiObjects->GetMDB();//do not release
	if (!lpMDB) return;

	EC_H(GetSpecialFolder(lpMDB,ulPropTag,&lpFolder));

	if (lpFolder)
	{
		EC_H(DisplayObject(
			lpFolder,
			NULL,
			otHierarchy,
			this));

		lpFolder->Release();
	}
	return;
}

//See Q171670 INFO: Entry IDs of Outlook Special Folders for more info on these tags
void CMsgStoreDlg::OnDisplayCalendarFolder()
{
	OnDisplaySpecialFolder(PR_IPM_APPOINTMENT_ENTRYID);
}

void CMsgStoreDlg::OnDisplayContactsFolder()
{
	OnDisplaySpecialFolder(PR_IPM_CONTACT_ENTRYID);
}

void CMsgStoreDlg::OnDisplayTasksFolder()
{
	OnDisplaySpecialFolder(PR_IPM_TASK_ENTRYID);
}

void CMsgStoreDlg::OnDisplayReceiveFolderTable()
{
	HRESULT			hRes = S_OK;
	LPMAPITABLE		lpMAPITable = NULL;

	if (!m_lpMapiObjects) return;

	LPMDB lpMDB = m_lpMapiObjects->GetMDB();//do not release
	if (!lpMDB) return;

	EC_H(lpMDB->GetReceiveFolderTable(
		fMapiUnicode,
		&lpMAPITable));
	if (lpMAPITable)
	{
		EC_H(DisplayTable(
			lpMAPITable,
			otReceive,
			this));
		lpMAPITable->Release();
	}
	return;
}//CMsgStoreDlg::OnDisplayReceiveFolderTable

void CMsgStoreDlg::OnDisplayOutgoingQueueTable()
{
	HRESULT			hRes = S_OK;
	LPMAPITABLE		lpMAPITable = NULL;

	if (!m_lpMapiObjects) return;

	LPMDB lpMDB = m_lpMapiObjects->GetMDB();//do not release
	if (!lpMDB) return;

	EC_H(lpMDB->GetOutgoingQueue(
		NULL,
		&lpMAPITable));

	if (lpMAPITable)
	{
		EC_H(DisplayTable(
			lpMAPITable,
			otDefault,
			this));
		lpMAPITable->Release();
	}
	return;
}//CMsgStoreDlg::OnDisplayOutgoingQueueTable

void CMsgStoreDlg::OnDisplayRulesTable()
{
	HRESULT			hRes = S_OK;

	if (!m_lpHierarchyTableTreeCtrl) return;

	LPMAPIFOLDER lpMAPIFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	if (lpMAPIFolder)
	{
		EC_H(DisplayExchangeTable(
			lpMAPIFolder,
			PR_RULES_TABLE,
			otRules,
			this));
		lpMAPIFolder->Release();
	}
	return;
}//CMsgStoreDlg::OnDisplayRulesTable

void CMsgStoreDlg::OnSelectForm()
{
	HRESULT			hRes = S_OK;
	LPMAPIFORMMGR	lpMAPIFormMgr = NULL;
	LPMAPIFORMINFO	lpMAPIFormInfo = NULL;

	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl || !m_lpPropDisplay) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession();//do not release
	if (!lpMAPISession) return;

	LPMAPIFOLDER lpMAPIFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	if (lpMAPIFolder)
	{
		EC_H(MAPIOpenFormMgr(lpMAPISession,&lpMAPIFormMgr));

		if (lpMAPIFormMgr)
		{
			// Apparently, SelectForm doesn't support unicode
			// CString doesn't provide a way to extract just ANSI strings, so we do this manually
			CHAR szTitle[256];
			int iRet = NULL;
			EC_D(iRet,LoadStringA(GetModuleHandle(NULL),
				IDS_SELECTFORMPROPS,
				szTitle,
				sizeof(szTitle)/sizeof(CHAR)));
#pragma warning(push)
#pragma warning(disable:4616)
#pragma warning(disable:6276)
			EC_H_CANCEL(lpMAPIFormMgr->SelectForm(
				(ULONG_PTR)m_hWnd,
				0,//fMapiUnicode,
				(LPCTSTR) szTitle,
				lpMAPIFolder,
				&lpMAPIFormInfo));
#pragma warning(pop)

			if (lpMAPIFormInfo)
			{
				EC_H(m_lpPropDisplay->SetDataSource(lpMAPIFormInfo,NULL,false));
				DebugPrintFormInfo(DBGForms,lpMAPIFormInfo);
				lpMAPIFormInfo->Release();
			}
			lpMAPIFormMgr->Release();
		}
		lpMAPIFolder->Release();
	}
}//CMsgStoreDlg::OnSelectForm

void CMsgStoreDlg::OnOpenFormContainer()
{
	HRESULT			hRes = S_OK;
	LPMAPIFORMMGR	lpMAPIFormMgr = NULL;
	LPMAPIFORMCONTAINER	lpMAPIFormContainer = NULL;

	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl || !m_lpPropDisplay) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession();//do not release
	if (!lpMAPISession) return;

	LPMAPIFOLDER lpMAPIFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	if (lpMAPIFolder)
	{
		EC_H(MAPIOpenFormMgr(lpMAPISession,&lpMAPIFormMgr));

		if (lpMAPIFormMgr)
		{
			EC_H(lpMAPIFormMgr->OpenFormContainer(
				HFRMREG_FOLDER,
				lpMAPIFolder,
				&lpMAPIFormContainer));

			if (lpMAPIFormContainer)
			{
				new CFormContainerDlg(
					m_lpParent,
					m_lpMapiObjects,
					lpMAPIFormContainer);

				lpMAPIFormContainer->Release();
			}
			lpMAPIFormMgr->Release();
		}
		lpMAPIFolder->Release();
	}
}//CMsgStoreDlg::OnOpenFormContainer

/////////////////////////////////////////////////////////////////////////////////////
//  Menu Commands

//newstyle copy folder
BOOL CMsgStoreDlg::HandleCopy()
{
	HRESULT hRes = S_OK;
	CWaitCursor	Wait;//Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric,CLASS,_T("OnCopyItems"),_T("\n"));
	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return false;

	//not needed - no case where we don't handle copy
//	if (CBaseDialog::HandleCopy()) return true;

	LPMAPIFOLDER lpMAPISourceFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	LPMDB lpMDB = m_lpMapiObjects->GetMDB();//do not release
	LPMAPIFOLDER lpSrcParentFolder = NULL;
	WC_H(GetParentFolder(lpMAPISourceFolder,lpMDB,&lpSrcParentFolder));

	m_lpMapiObjects->SetFolderToCopy(lpMAPISourceFolder,lpSrcParentFolder);

	if (lpSrcParentFolder) lpSrcParentFolder->Release();
	if (lpMAPISourceFolder) lpMAPISourceFolder->Release();

	return true;
}

BOOL CMsgStoreDlg::HandlePaste()
{
	if (CBaseDialog::HandlePaste()) return true;

	HRESULT		hRes = S_OK;
	CWaitCursor	Wait;//Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric,CLASS,_T("HandlePaste"),_T("\n"));
	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return false;

	ULONG ulStatus = m_lpMapiObjects->GetBufferStatus();

	//Get the destination Folder
	LPMAPIFOLDER	lpMAPIDestFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	if (lpMAPIDestFolder && (ulStatus & BUFFER_MESSAGES) && (ulStatus & BUFFER_PARENTFOLDER))
	{
		OnPasteMessages();
	}
	else if (lpMAPIDestFolder && (ulStatus & BUFFER_FOLDER))
	{
		CEditor MyData(
			this,
			IDS_PASTEFOLDER,
			IDS_PASTEFOLDERPROMPT,
			1,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

		MyData.InitCheck(0,IDS_PASTEFOLDERCONTENTS,false,false);
		WC_H(MyData.DisplayDialog());
		if (S_OK == hRes)
		{
			BOOL bPasteContents = MyData.GetCheck(0);
			if (bPasteContents) OnPasteFolderContents();
			else OnPasteFolder();
		}
	}
	if (lpMAPIDestFolder) lpMAPIDestFolder->Release();
	return true;
}//CMsgStoreDlg::HandlePaste

void CMsgStoreDlg::OnPasteMessages()
{
	HRESULT			hRes = S_OK;
	CWaitCursor	Wait;//Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric,CLASS,_T("OnPasteMessages"),_T("\n"));
	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return;

	//Get the source Messages
	LPENTRYLIST		lpEIDs = m_lpMapiObjects->GetMessagesToCopy();
	LPMAPIFOLDER	lpMAPISourceFolder = m_lpMapiObjects->GetSourceParentFolder();
	//Get the destination Folder
	LPMAPIFOLDER	lpMAPIDestFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	if (lpMAPIDestFolder && lpMAPISourceFolder && lpEIDs)
	{
		CEditor MyData(
			this,
			IDS_COPYMESSAGE,
			IDS_COPYMESSAGEPROMPT,
			1,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

		MyData.InitCheck(0,IDS_MESSAGEMOVE,false,false);
		WC_H(MyData.DisplayDialog());
		if (S_OK == hRes)
		{
			ULONG ulMoveMessage = MyData.GetCheck(0)?MESSAGE_MOVE:0;

			LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IMAPIFolder::CopyMessages"), m_hWnd);// STRING_OK

			if(lpProgress)
				ulMoveMessage |= MESSAGE_DIALOG;

			EC_H(lpMAPISourceFolder->CopyMessages(
				lpEIDs,
				&IID_IMAPIFolder,
				lpMAPIDestFolder,
				lpProgress ? (ULONG_PTR)m_hWnd : NULL,
				lpProgress,
				ulMoveMessage));

			if(lpProgress)
				lpProgress->Release();

			lpProgress = NULL;
		}
	}
	if (lpMAPIDestFolder) lpMAPIDestFolder->Release();
	if (lpMAPISourceFolder) lpMAPISourceFolder->Release();
}//CMsgStoreDlg::OnPasteMessages

void CMsgStoreDlg::OnPasteFolder()
{
	HRESULT			hRes = S_OK;
	ULONG			cProps;
	LPSPropValue	lpProps = NULL;

	if (!m_lpMapiObjects) return;

	enum {NAME,EID,NUM_COLS};
	SizedSPropTagArray(NUM_COLS,sptaSrcFolder) = { NUM_COLS, {
			PR_DISPLAY_NAME,
			PR_ENTRYID}
	};

	DebugPrintEx(DBGGeneric,CLASS,_T("OnPasteFolder"),_T("\n"));

	//Get the source folder
	LPMAPIFOLDER lpMAPISourceFolder = m_lpMapiObjects->GetFolderToCopy();
	LPMAPIFOLDER lpSrcParentFolder = m_lpMapiObjects->GetSourceParentFolder();
	//Get the Destination Folder
	LPMAPIFOLDER lpMAPIDestFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	if (lpMAPISourceFolder && lpMAPIDestFolder)
	{
		DebugPrint(DBGGeneric,_T("Folder Source Object = 0x%08X\n"),lpMAPISourceFolder);
		DebugPrint(DBGGeneric,_T("Folder Source Object Parent = 0x%08X\n"),lpSrcParentFolder);
		DebugPrint(DBGGeneric,_T("Folder Destination Object = 0x%08X\n"),lpMAPIDestFolder);

		CEditor MyData(
			this,
			IDS_PASTEFOLDER,
			IDS_PASTEFOLDERNEWNAMEPROMPT,
			3,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
		MyData.InitSingleLine(0,IDS_FOLDERNAME,NULL,false);
		MyData.InitCheck(1,IDS_COPYSUBFOLDERS,false,false);
		MyData.InitCheck(2,IDS_FOLDERMOVE,false,false);

		// Get required properties from the source folder
		EC_H_GETPROPS(lpMAPISourceFolder->GetProps(
			(LPSPropTagArray) &sptaSrcFolder,
			fMapiUnicode,
			&cProps,
			&lpProps));

		if (lpProps)
		{
			if (CheckStringProp(&lpProps[NAME],PT_TSTRING))
			{
				DebugPrint(DBGGeneric,_T("Folder Source Name = \"%s\"\n"),lpProps[NAME].Value.LPSZ);
				MyData.SetString(0,lpProps[NAME].Value.LPSZ);
			}
		}

		WC_H(MyData.DisplayDialog());

		LPMAPIFOLDER lpCopyRoot = lpSrcParentFolder;
		if (!lpSrcParentFolder) lpCopyRoot = lpMAPIDestFolder;

		if (S_OK == hRes)
		{
			CWaitCursor		Wait;//Change the mouse to an hourglass while we work.

			LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IMAPIFolder::CopyFolder"), m_hWnd);// STRING_OK

			ULONG ulCopyFlags = fMapiUnicode;
			if(MyData.GetCheck(1))
				ulCopyFlags |= COPY_SUBFOLDERS;
			if(MyData.GetCheck(2))
				ulCopyFlags |= FOLDER_MOVE;
			if(lpProgress)
				ulCopyFlags |= FOLDER_DIALOG;

			hRes = lpCopyRoot->CopyFolder(
				lpProps[EID].Value.bin.cb,
				(LPENTRYID) lpProps[EID].Value.bin.lpb,
				&IID_IMAPIFolder,
				lpMAPIDestFolder,
				MyData.GetString(0),
				lpProgress ? (ULONG_PTR)m_hWnd : NULL,//UI
				lpProgress,//Progress
				ulCopyFlags
				);
			if (MAPI_E_COLLISION == hRes)
			{
				ErrDialog(__FILE__,__LINE__,IDS_EDDUPEFOLDER);
				hRes = S_OK;
			}
			else CHECKHRESMSG(hRes,IDS_COPYFOLDERFAILED);

			if(lpProgress)
				lpProgress->Release();

			lpProgress = NULL;
		}
		MAPIFreeBuffer(lpProps);
	}
	if (lpMAPIDestFolder) lpMAPIDestFolder->Release();
	if (lpMAPISourceFolder) lpMAPISourceFolder->Release();
	if (lpSrcParentFolder) lpSrcParentFolder->Release();
	return;
}

void CMsgStoreDlg::OnPasteFolderContents()
{
	HRESULT			hRes = S_OK;

	DebugPrintEx(DBGGeneric,CLASS,_T("OnPasteFolderContents"),_T("\n"));

	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return;

	//Get the Source Folder
	LPMAPIFOLDER lpMAPISourceFolder = m_lpMapiObjects->GetFolderToCopy();
	//Get the Destination Folder
	LPMAPIFOLDER lpMAPIDestFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	if (lpMAPISourceFolder && lpMAPIDestFolder)
	{
		DebugPrint(DBGGeneric,_T("Folder Source Object = 0x%08X\n"),lpMAPISourceFolder);
		DebugPrint(DBGGeneric,_T("Folder Destination Object = 0x%08X\n"),lpMAPIDestFolder);

		CEditor MyData(
			this,
			IDS_COPYFOLDERCONTENTS,
			IDS_PICKOPTIONSPROMPT,
			3,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
		MyData.InitCheck(0,IDS_COPYASSOCIATEDITEMS,false,false);
		MyData.InitCheck(1,IDS_MOVEMESSAGES,false,false);
		MyData.InitCheck(2,IDS_SINGLECALLCOPY,false,false);
		WC_H(MyData.DisplayDialog());

		if (S_OK == hRes)
		{
			CWaitCursor		Wait;//Change the mouse to an hourglass while we work.

			EC_H(CopyFolderContents(
				lpMAPISourceFolder,
				lpMAPIDestFolder,
				MyData.GetCheck(0),//associated contents
				MyData.GetCheck(1),//move
				MyData.GetCheck(2),//Single CopyMessages call
				m_hWnd));
		}

	}
	if (lpMAPIDestFolder) lpMAPIDestFolder->Release();
	if (lpMAPISourceFolder) lpMAPISourceFolder->Release();
	return;
}//CMsgStoreDlg::OnPasteFolderContents

void CMsgStoreDlg::OnPasteRules()
{
	HRESULT			hRes = S_OK;

	DebugPrintEx(DBGGeneric,CLASS,_T("OnPasteRules"),_T("\n"));

	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return;

	//Get the Source Folder
	LPMAPIFOLDER lpMAPISourceFolder = m_lpMapiObjects->GetFolderToCopy();
	//Get the Destination Folder
	LPMAPIFOLDER lpMAPIDestFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	if (lpMAPISourceFolder && lpMAPIDestFolder)
	{
		DebugPrint(DBGGeneric,_T("Folder Source Object = 0x%08X\n"),lpMAPISourceFolder);
		DebugPrint(DBGGeneric,_T("Folder Destination Object = 0x%08X\n"),lpMAPIDestFolder);

		CEditor MyData(
			this,
			IDS_COPYFOLDERRULES,
			IDS_COPYFOLDERRULESPROMPT,
			1,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
		MyData.InitCheck(0,IDS_REPLACERULES,false,false);
		WC_H(MyData.DisplayDialog());

		if (S_OK == hRes)
		{
			CWaitCursor		Wait;//Change the mouse to an hourglass while we work.

			EC_H(CopyFolderRules(
				lpMAPISourceFolder,
				lpMAPIDestFolder,
				MyData.GetCheck(0)));//move
		}

	}
	if (lpMAPIDestFolder) lpMAPIDestFolder->Release();
	if (lpMAPISourceFolder) lpMAPISourceFolder->Release();
	return;
}//CMsgStoreDlg::OnPasteRules

void CMsgStoreDlg::OnCreateSubFolder()
{
	HRESULT			hRes = S_OK;
	LPMAPIFOLDER	lpMAPISubFolder = NULL;

	CEditor MyData(
		this,
		IDS_ADDSUBFOLDER,
		IDS_ADDSUBFOLDERPROMPT,
		4,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
	MyData.SetPromptPostFix(AllFlagsToString(PROP_ID(PR_FOLDER_TYPE),true));
	MyData.InitSingleLine(0,IDS_FOLDERNAME,IDS_FOLDERNAMEVALUE,false);
	MyData.InitSingleLine(1,IDS_FOLDERTYPE,NULL,false);
	MyData.SetHex(1,FOLDER_GENERIC);
	CString szProduct;
	CString szFolderComment;
	szProduct.LoadString(IDS_PRODUCT_NAME);
	szFolderComment.FormatMessage(IDS_FOLDERCOMMENTVALUE,szProduct);
	MyData.InitSingleLineSz(2,IDS_FOLDERCOMMENT,szFolderComment,false);
	MyData.InitCheck(3,IDS_PASSOPENIFEXISTS,false,false);

	LPMAPIFOLDER lpMAPIFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(
		mfcmapiREQUEST_MODIFY);

	if (lpMAPIFolder)
	{
		WC_H(MyData.DisplayDialog());

		EC_H(lpMAPIFolder->CreateFolder(
			MyData.GetHex(1),
			MyData.GetString(0),
			MyData.GetString(2),
			NULL,//interface
			fMapiUnicode
			| (MyData.GetCheck(3)?OPEN_IF_EXISTS:0),
			&lpMAPISubFolder));

		if (lpMAPISubFolder) lpMAPISubFolder->Release();
		lpMAPIFolder->Release();
	}
	return;
}//CMsgStoreDlg::OnCreateSubFolder

void CMsgStoreDlg::OnDisplayACLTable()
{
	HRESULT			hRes = S_OK;

	if (!m_lpHierarchyTableTreeCtrl) return;

	LPMAPIFOLDER lpMAPIFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	if (lpMAPIFolder)
	{
		EC_H(DisplayExchangeTable(
			lpMAPIFolder,
			PR_ACL_TABLE,
			otACL,
			this));
		lpMAPIFolder->Release();
	}
	return;
}//CMsgStoreDlg::OnDisplayACLTable

void CMsgStoreDlg::OnDisplayAssociatedContents()
{
	if (!m_lpHierarchyTableTreeCtrl) return;

	//Find the highlighted item
	LPMAPIFOLDER lpMAPIFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	if (lpMAPIFolder)
	{
		DisplayObject(
			lpMAPIFolder,
			NULL,
			otAssocContents,
			this);

		lpMAPIFolder->Release();
	}
	return;
}//CMsgStoreDlg::OnDisplayAssociatedContents

void CMsgStoreDlg::OnDisplayDeletedContents()
{
	if (!m_lpHierarchyTableTreeCtrl || !m_lpMapiObjects) return;

	//Find the highlighted item
	HRESULT		hRes = S_OK;
	LPSBinary	lpItemEID = NULL;
	lpItemEID = m_lpHierarchyTableTreeCtrl->GetSelectedItemEID();

	if (lpItemEID)
	{
		LPMDB lpMDB = m_lpMapiObjects->GetMDB();//do not release
		if (lpMDB)
		{
			LPMAPIFOLDER	lpMAPIFolder = NULL;
			WC_H(CallOpenEntry(
				lpMDB,
				NULL,
				NULL,
				NULL,
				lpItemEID->cb,
				(LPENTRYID) lpItemEID->lpb,
				NULL,
				MAPI_BEST_ACCESS | SHOW_SOFT_DELETES | MAPI_NO_CACHE,
				NULL,
				(LPUNKNOWN*)&lpMAPIFolder));
			if (lpMAPIFolder)
			{
				//call the dialog
				new CFolderDlg(
					m_lpParent,
					m_lpMapiObjects,
					lpMAPIFolder,
					mfcmapiSHOW_NORMAL_CONTENTS,
					mfcmapiSHOW_DELETED_ITEMS
					);
				lpMAPIFolder->Release();
			}
		}
	}
}//CMsgStoreDlg::OnDisplayDeletedContents

void CMsgStoreDlg::OnDisplayDeletedSubFolders()
{
	if (!m_lpHierarchyTableTreeCtrl) return;

	//Must open the folder with MODIFY permissions if I'm going to restore the folder!
	LPMAPIFOLDER lpFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	if (lpFolder)
	{
//		if (ContainerSupportsDeletedItems(lpFolder))
		{
			new CMsgStoreDlg(
				m_lpParent,
				m_lpMapiObjects,
				lpFolder,
				mfcmapiSHOW_DELETED_ITEMS);
		}
//		else ErrDialog(__FILE__,__LINE__,_T("Folder does not support SHOW_SOFT_DELETES!"));// STRING_OK
		lpFolder->Release();
	}
	return;

}//CMsgStoreDlg::OnDisplayDeletedSubFolders

void CMsgStoreDlg::OnDisplayMailboxTable()
{
	if (!m_lpParent || !m_lpMapiObjects) return;

	DisplayMailboxTable(m_lpParent,m_lpMapiObjects);
}

void CMsgStoreDlg::OnEmptyFolder()
{
	HRESULT		hRes = S_OK;
	ULONG		ulDelAssociated = NULL;

	if (!m_lpHierarchyTableTreeCtrl) return;

	//Find the highlighted item
	LPMAPIFOLDER lpMAPIFolderToDelete = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	if (lpMAPIFolderToDelete)
	{
		CEditor MyData(
			this,
			IDS_DELETEITEMSANDSUB,
			IDS_DELETEITEMSANDSUBPROMPT,
			1,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
		MyData.InitCheck(0,IDS_DELASSOCIATED,false,false);

		WC_H(MyData.DisplayDialog());
		if (S_OK == hRes)
		{
			ulDelAssociated = MyData.GetCheck(0)?DEL_ASSOCIATED:0;

			DebugPrintEx(DBGGeneric,CLASS,_T("OnEmptyFolder"),_T("Calling EmptyFolder on 0x%08X.\n"),lpMAPIFolderToDelete);

			LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IMAPIFolder::EmptyFolder"), m_hWnd);// STRING_OK

			if(lpProgress)
				ulDelAssociated |= FOLDER_DIALOG;

			EC_H(lpMAPIFolderToDelete->EmptyFolder(
				lpProgress ? (ULONG_PTR)m_hWnd : NULL,
				lpProgress,
				ulDelAssociated));

			if(lpProgress)
				lpProgress->Release();

			lpProgress = NULL;
		}
		lpMAPIFolderToDelete->Release();
	}
	return;
}//CMsgStoreDlg::OnEmptyFolder

void CMsgStoreDlg::OnDeleteSelectedItem()
{
	HRESULT			hRes = S_OK;
	LPSBinary		lpItemEID = NULL;
	ULONG			ulFlags = NULL;

	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return;

	ULONG bShiftPressed = GetKeyState(VK_SHIFT) <0;

	lpItemEID = m_lpHierarchyTableTreeCtrl->GetSelectedItemEID();//never free this!!!!!
	if (!lpItemEID) return;

	LPMDB lpMDB = m_lpMapiObjects->GetMDB();//do not release
	if (!lpMDB) return;

	LPMAPIFOLDER lpFolderToDelete = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(
		mfcmapiDO_NOT_REQUEST_MODIFY);

	if (lpFolderToDelete)
	{
		LPMAPIFOLDER lpParentFolder = NULL;
		EC_H(GetParentFolder(lpFolderToDelete,lpMDB,&lpParentFolder));
		if (lpParentFolder)
		{
			CEditor MyData(
				this,
				IDS_DELETEFOLDER,
				IDS_DELETEFOLDERPROMPT,
				1,
				CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
			MyData.InitCheck(0,IDS_HARDDELETION,false,false);
			if (!bShiftPressed) WC_H(MyData.DisplayDialog());
			if (S_OK == hRes)
			{
				ulFlags = DEL_FOLDERS | DEL_MESSAGES;
				ulFlags |= (bShiftPressed || MyData.GetCheck(0))?DELETE_HARD_DELETE:0;

				DebugPrintEx(DBGDeleteSelectedItem,CLASS,_T("OnDeleteSelectedItem"),_T("Calling DeleteFolder on folder. ulFlags = 0x%08X.\n"),ulFlags);
				DebugPrintBinary(DBGGeneric,lpItemEID);

				LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IMAPIFolder::DeleteFolder"), m_hWnd);// STRING_OK

				if(lpProgress)
					ulFlags |= FOLDER_DIALOG;

				EC_H(lpParentFolder->DeleteFolder(
					lpItemEID->cb,
					(LPENTRYID) lpItemEID->lpb,
					lpProgress ? (ULONG_PTR)m_hWnd : NULL,
					lpProgress,
					ulFlags));

				if(lpProgress)
					lpProgress->Release();

				lpProgress = NULL;
			}
			lpParentFolder->Release();
		}
		lpFolderToDelete->Release();
	}
	return;
}//CMsgStoreDlg::OnDeleteSelectedItem()

void CMsgStoreDlg::OnSaveFolderContentsAsMSG()
{
	HRESULT			hRes = S_OK;
	TCHAR			szFilePath[MAX_PATH];

	if (!m_lpHierarchyTableTreeCtrl) return;

	DebugPrintEx(DBGGeneric,CLASS,_T("OnSaveFolderContentsAsMSG"),_T("\n"));

	//Find the highlighted item
	LPMAPIFOLDER lpMAPIFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiDO_NOT_REQUEST_MODIFY);
	if (!lpMAPIFolder) return;

	DebugPrint(DBGGeneric,_T("Saving items from 0x%08X\n"),lpMAPIFolder);

	CEditor MyData(
		this,
		IDS_SAVEFOLDERASMSG,
		IDS_SAVEFOLDERASMSGPROMPT,
		2,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
	MyData.InitCheck(0,IDS_SAVEASSOCIATED,false,false);
	MyData.InitCheck(1,IDS_SAVEUNICODE,false,false);
	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		WC_H(GetDirectoryPath(szFilePath));

		if (S_OK == hRes && szFilePath[0])
		{
			CWaitCursor		Wait;//Change the mouse to an hourglass while we work.

			EC_H(SaveFolderContentsToMSG(
				lpMAPIFolder,
				szFilePath,
				MyData.GetCheck(0)?TRUE:FALSE,
				MyData.GetCheck(1)?TRUE:FALSE,
				m_hWnd));
		}
	}

	lpMAPIFolder->Release();
	return;
}//CMsgStoreDlg::OnSaveFolderContentsAsMSG()

void CMsgStoreDlg::OnSaveFolderContentsAsTextFiles()
{
	HRESULT			hRes = S_OK;
	TCHAR			szPathName[MAX_PATH];

	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return;

	LPMDB lpMDB = m_lpMapiObjects->GetMDB();//do not release
	if (!lpMDB) return;

	LPMAPIFOLDER lpFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiDO_NOT_REQUEST_MODIFY);

	if (lpFolder)
	{
		CEditor MyData(
			this,
			IDS_SAVEFOLDERASPROPFILES,
			IDS_PICKOPTIONSPROMPT,
			3,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
		MyData.InitCheck(0,IDS_RECURSESUBFOLDERS,false,false);
		MyData.InitCheck(1,IDS_SAVEREGULARCONTENTS,true,false);
		MyData.InitCheck(2,IDS_SAVEASSOCIATEDCONTENTS,true,false);

		WC_H(MyData.DisplayDialog());

		if (S_OK == hRes)
		{
			WC_H(GetDirectoryPath(szPathName));

			if (S_OK == hRes && szPathName[0])
			{
				CWaitCursor		Wait;//Change the mouse to an hourglass while we work.

				CDumpStore MyDumpStore;
				MyDumpStore.InitMDB(lpMDB);
				MyDumpStore.InitFolder(lpFolder);
				MyDumpStore.InitFolderPathRoot(szPathName);
				MyDumpStore.ProcessFolders(
					MyData.GetCheck(1),
					MyData.GetCheck(2),
					MyData.GetCheck(0));
			}
		}
		lpFolder->Release();
	}
	return;
}//CMsgStoreDlg::OnSaveFolderContentsAsTextFiles

void CMsgStoreDlg::OnSetReceiveFolder()
{
	HRESULT			hRes = S_OK;

	CWaitCursor		Wait;//Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return;

	LPMDB lpMDB = m_lpMapiObjects->GetMDB();//do not release
	if (!lpMDB) return;

	CEditor MyData(
		this,
		IDS_SETRECFOLDER,
		IDS_SETRECFOLDERPROMPT,
		2,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
	MyData.InitSingleLine(0,IDS_CLASS,NULL,false);
	MyData.InitCheck(1,IDS_DELETEASSOCIATION,false,false);

	//Find the highlighted item
	LPSBinary lpEID = m_lpHierarchyTableTreeCtrl->GetSelectedItemEID();

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		if (MyData.GetCheck(1))
		{
			EC_H(lpMDB->SetReceiveFolder(
				MyData.GetString(0),
				fMapiUnicode,
				NULL,
				NULL));
		}
		else if (lpEID)
		{
			EC_H(lpMDB->SetReceiveFolder(
				MyData.GetString(0),
				fMapiUnicode,
				lpEID->cb,
				(LPENTRYID) lpEID->lpb));
		}
	}
	return;
}

void CMsgStoreDlg::OnResendAllMessages()
{
	HRESULT			hRes = S_OK;
	CWaitCursor		Wait;//Change the mouse to an hourglass while we work.

	if (!m_lpHierarchyTableTreeCtrl) return;

	//Find the highlighted item
	LPMAPIFOLDER lpMAPIFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	if (lpMAPIFolder)
	{
		EC_H(ResendMessages(lpMAPIFolder, m_hWnd));

		lpMAPIFolder->Release();
	}
	return;
}//CMsgStoreDlg::OnResendAllMessages

//Iterate through items in the selected folder and attempt to delete PR_NT_SECURITY_DESCRIPTOR
void CMsgStoreDlg::OnResetPermissionsOnItems()
{
	HRESULT			hRes = S_OK;
	CWaitCursor		Wait;//Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return;

	LPMDB lpMDB = m_lpMapiObjects->GetMDB();//do not release

	//Find the highlighted item
	LPMAPIFOLDER lpMAPIFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	if (lpMAPIFolder)
	{
		EC_H(ResetPermissionsOnItems(lpMDB,lpMAPIFolder));
		lpMAPIFolder->Release();
	}
	return;
}//CMsgStoreDlg::OnResetPermissionsOnItems

//Copy selected folder back to the land of the living
void CMsgStoreDlg::OnRestoreDeletedFolder()
{
	HRESULT			hRes = S_OK;
	ULONG			cProps;
	LPSPropValue	lpProps = NULL;

	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return;

	enum {NAME,EID,NUM_COLS};
	SizedSPropTagArray(NUM_COLS,sptaSrcFolder) = { NUM_COLS, {
			PR_DISPLAY_NAME,
			PR_ENTRYID}
	};

	LPMDB lpMDB = m_lpMapiObjects->GetMDB();//do not release
	if (!lpMDB) return;

	LPMAPIFOLDER lpSrcFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	if (lpSrcFolder)
	{
		LPMAPIFOLDER lpSrcParentFolder = NULL;
		WC_H(GetParentFolder(lpSrcFolder,lpMDB,&lpSrcParentFolder));
		hRes = S_OK;

		// Get required properties from the source folder
		EC_H_GETPROPS(lpSrcFolder->GetProps(
			(LPSPropTagArray) &sptaSrcFolder,
			fMapiUnicode,
			&cProps,
			&lpProps));

		CEditor MyData(
			this,
			IDS_RESTOREDELFOLD,
			IDS_RESTOREDELFOLDPROMPT,
			2,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
		MyData.InitSingleLine(0,IDS_FOLDERNAME,NULL,false);
		MyData.InitCheck(1,IDS_COPYSUBFOLDERS,false,false);

		if (lpProps)
		{
			if (CheckStringProp(&lpProps[NAME],PT_TSTRING))
			{
				DebugPrint(DBGGeneric,_T("Folder Source Name = \"%s\"\n"),lpProps[NAME].Value.LPSZ);
				MyData.SetString(0,lpProps[NAME].Value.LPSZ);
			}
		}

		WC_H(MyData.DisplayDialog());

		if (S_OK == hRes)
		{
			//Restore the folder up under m_lpContainer
			CWaitCursor		Wait;//Change the mouse to an hourglass while we work.

			DebugPrint(DBGGeneric,_T("Restoring 0x%X to 0x%X as \n"),lpSrcFolder,m_lpContainer);

			LPMAPIFOLDER lpCopyRoot = lpSrcParentFolder;
			if (!lpSrcParentFolder) lpCopyRoot = (LPMAPIFOLDER) m_lpContainer;

			LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IMAPIFolder::CopyFolder"), m_hWnd);// STRING_OK

			ULONG ulCopyFlags = fMapiUnicode | (MyData.GetCheck(1)?COPY_SUBFOLDERS:0);

			if(lpProgress)
				ulCopyFlags |= FOLDER_DIALOG;

			hRes = lpCopyRoot->CopyFolder(
				lpProps[EID].Value.bin.cb,
				(LPENTRYID) lpProps[EID].Value.bin.lpb,
				&IID_IMAPIFolder,
				(LPMAPIFOLDER) m_lpContainer,
				MyData.GetString(0),
				lpProgress ? (ULONG_PTR)m_hWnd : NULL,
				lpProgress,
				ulCopyFlags);
			if (MAPI_E_COLLISION == hRes)
			{
				ErrDialog(__FILE__,__LINE__,IDS_EDDUPEFOLDER);
				hRes = S_OK;
			}
			else if (MAPI_W_PARTIAL_COMPLETION == hRes)
			{
				ErrDialog(__FILE__,__LINE__,IDS_EDRESTOREFAILED);
			}
			else CHECKHRESMSG(hRes,IDS_COPYFOLDERFAILED);

			if(lpProgress)
				lpProgress->Release();

			lpProgress = NULL;
		}

		MAPIFreeBuffer(lpProps);
		if (lpSrcParentFolder) lpSrcParentFolder->Release();
		lpSrcFolder->Release();
	}
	return;
}//CMsgStoreDlg::OnRestoreDeletedFolder

void CMsgStoreDlg::OnValidateIPMSubtree()
{
	HRESULT		hRes = S_OK;
	ULONG		ulFlags = NULL;
	ULONG		ulValues = 0;
	LPSPropValue lpProps = NULL;
	LPMAPIERROR lpErr = NULL;

	CEditor MyData(
		this,
		IDS_VALIDATEIPMSUB,
		IDS_PICKOPTIONSPROMPT,
		2,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
	MyData.InitCheck(0,IDS_MAPIFORCECREATE,false,false);
	MyData.InitCheck(1,IDS_MAPIFULLIPMTREE,false,false);

	if (!m_lpMapiObjects) return;

	LPMDB lpMDB = m_lpMapiObjects->GetMDB();//do not release
	if (!lpMDB) return;

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		ulFlags = (MyData.GetCheck(0)?MAPI_FORCE_CREATE:0) |
			(MyData.GetCheck(1)?MAPI_FULL_IPM_TREE:0);

		DebugPrintEx(DBGGeneric,CLASS,_T("OnValidateIPMSubtree"),_T("ulFlags = 0x%08X\n"),ulFlags);

		EC_H(HrValidateIPMSubtree(
			lpMDB,
			ulFlags,
			&ulValues,
			&lpProps,
			&lpErr));
		EC_MAPIERR(fMapiUnicode,lpErr);
		MAPIFreeBuffer(lpErr);

		if (ulValues > 0 && lpProps)
		{
			DebugPrintEx(DBGGeneric,CLASS,_T("OnValidateIPMSubtree"),_T("HrValidateIPMSubtree returned 0x%08X properties:\n"),ulValues);
			DebugPrintProperties(DBGGeneric,ulValues,lpProps,lpMDB);
		}

		MAPIFreeBuffer(lpProps);
	}
	return;
}

void CMsgStoreDlg::HandleAddInMenuSingle(
									   LPADDINMENUPARAMS lpParams,
									   LPMAPIPROP /*lpMAPIProp*/,
									   LPMAPICONTAINER lpContainer)
{
	if (lpParams)
	{
		lpParams->lpFolder = (LPMAPIFOLDER) lpContainer; // GetSelectedContainer returns LPMAPIFOLDER
	}

	InvokeAddInMenu(lpParams);
} // CMsgStoreDlg::HandleAddInMenuSingle
