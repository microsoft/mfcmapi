// FolderDlg.cpp : implementation file
// Displays the contents of a folder

#include "stdafx.h"
#include "FolderDlg.h"
#include "ContentsTableListCtrl.h"
#include "MapiObjects.h"
#include "MAPIFunctions.h"
#include "MAPIStoreFunctions.h"
#include "MAPIABFunctions.h"
#include "SingleMAPIPropListCtrl.h"
#include "ColumnTags.h"
#include "MFCUtilityFunctions.h"
#include "Editor.h"
#include "DumpStore.h"
#include "File.h"
#include "AttachmentsDlg.h"
#include "RecipientsDlg.h"
#include "MAPIFormFunctions.h"
#include "TagArrayEditor.h"
#include "InterpretProp.h"
#include "FileDialogEx.h"
#include "ExtraPropTags.h"
#include "PropertyTagEditor.h"
#include "MAPIProgress.h"
#include "MAPIMime.h"
#include "InterpretProp2.h"

static TCHAR* CLASS = _T("CFolderDlg");

/////////////////////////////////////////////////////////////////////////////
// CFolderDlg dialog

CFolderDlg::CFolderDlg(
					   _In_ CParentWnd* pParentWnd,
					   _In_ CMapiObjects* lpMapiObjects,
					   _In_ LPMAPIFOLDER lpMAPIFolder,
					   ULONG ulDisplayFlags
					   ):
CContentsTableDlg(
				  pParentWnd,
				  lpMapiObjects,
				  IDS_FOLDER,
				  mfcmapiDO_NOT_CALL_CREATE_DIALOG,
				  NULL,
				  (LPSPropTagArray) &sptMSGCols,
				  NUMMSGCOLUMNS,
				  MSGColumns,
				  IDR_MENU_FOLDER_POPUP,
				  MENU_CONTEXT_FOLDER_CONTENTS)
{
	TRACE_CONSTRUCTOR(CLASS);
	m_ulDisplayFlags = ulDisplayFlags;

	m_lpContainer = lpMAPIFolder;
	if (m_lpContainer) m_lpContainer->AddRef();

	CreateDialogAndMenu(IDR_MENU_FOLDER);
} // CFolderDlg::CFolderDlg

CFolderDlg::~CFolderDlg()
{
	TRACE_DESTRUCTOR(CLASS);
} // CFolderDlg::~CFolderDlg

_Check_return_ bool CFolderDlg::HandleMenu(WORD wMenuSelect)
{
	DebugPrint(DBGMenu,_T("CFolderDlg::HandleMenu wMenuSelect = 0x%X = %d\n"),wMenuSelect,wMenuSelect);
	HRESULT hRes = S_OK;
	switch (wMenuSelect)
	{
	case ID_DISPLAYACLTABLE:
		if (m_lpContainer)
		{
			EC_H(DisplayExchangeTable(
				m_lpContainer,
				PR_ACL_TABLE,
				otACL,
				this));
		}
		return true;
	case ID_LOADFROMMSG: OnLoadFromMSG(); return true;
	case ID_LOADFROMTNEF: OnLoadFromTNEF(); return true;
	case ID_LOADFROMEML: OnLoadFromEML(); return true;
	case ID_RESOLVEMESSAGECLASS: OnResolveMessageClass(); return true;
	case ID_SELECTFORM: OnSelectForm(); return true;
	case ID_MANUALRESOLVE: OnManualResolve(); return true;
	case ID_SAVEFOLDERCONTENTSASTEXTFILES: OnSaveFolderContentsAsTextFiles(); return true;
	case ID_SENDBULKMAIL: OnSendBulkMail(); return true;
	case ID_NEW_CUSTOMFORM: OnNewCustomForm(); return true;
	case ID_NEW_MESSAGE: OnNewMessage(); return true;
	case ID_NEW_APPOINTMENT:
	case ID_NEW_CONTACT:
	case ID_NEW_IPMNOTE:
	case ID_NEW_IPMPOST:
	case ID_NEW_TASK:
	case ID_NEW_STICKYNOTE:
		NewSpecialItem(wMenuSelect);
		return true;
	case ID_EXECUTEVERBONFORM: OnExecuteVerbOnForm(); return true;
	case ID_GETPROPSUSINGLONGTERMEID: OnGetPropsUsingLongTermEID(); return true;
	}

	if (MultiSelectSimple(wMenuSelect)) return true;

	if (MultiSelectComplex(wMenuSelect)) return true;

	return CContentsTableDlg::HandleMenu(wMenuSelect);
} // CFolderDlg::HandleMenu

typedef HRESULT (CFolderDlg::* LPSIMPLEMULTI)
(
	int				iItem,
	SortListData*	lpData
);

_Check_return_ bool CFolderDlg::MultiSelectSimple(WORD wMenuSelect)
{
	LPSIMPLEMULTI	lpFunc = NULL;
	HRESULT			hRes = S_OK;

	if (m_lpContentsTableListCtrl)
	{
		switch (wMenuSelect)
		{
		case ID_ATTACHMENTPROPERTIES:
			lpFunc = &CFolderDlg::OnAttachmentProperties;
			break;
		case ID_OPENMODAL:
			lpFunc = &CFolderDlg::OnOpenModal;
			break;
		case ID_OPENNONMODAL:
			lpFunc = &CFolderDlg::OnOpenNonModal;
			break;
		case ID_RESENDMESSAGE:
			lpFunc = &CFolderDlg::OnResendSelectedItem;
			break;
		case ID_SAVEATTACHMENTS:
			lpFunc = &CFolderDlg::OnSaveAttachments;
			break;
		case ID_GETMESSAGESTATUS:
			lpFunc = &CFolderDlg::OnGetMessageStatus;
			break;
		case ID_SUBMITMESSAGE:
			lpFunc = &CFolderDlg::OnSubmitMessage;
			break;
		case ID_ABORTSUBMIT:
			lpFunc = &CFolderDlg::OnAbortSubmit;
			break;
		case ID_RECIPIENTPROPERTIES:
			lpFunc = &CFolderDlg::OnRecipientProperties;
			break;
		}
		if (lpFunc)
		{
			SortListData*	lpData = 0;
			int iItem = m_lpContentsTableListCtrl->GetNextItem(
				-1,
				LVNI_SELECTED);
			while (-1 != iItem)
			{
				lpData = (SortListData*) m_lpContentsTableListCtrl->GetItemData(iItem);
				WC_H((this->*lpFunc)(iItem,lpData));
				iItem = m_lpContentsTableListCtrl->GetNextItem(
					iItem,
					LVNI_SELECTED);
				if (S_OK != hRes && -1 != iItem)
				{
					if (bShouldCancel(this,hRes)) break;
					hRes = S_OK;
				}
			}
			return true;
		}
	}
	return false;
} // CFolderDlg::MultiSelectSimple

_Check_return_ bool CFolderDlg::MultiSelectComplex(WORD wMenuSelect)
{
	switch (wMenuSelect)
	{
	case ID_ADDTESTADDRESS: OnAddOneOffAddress(); return true;
	case ID_DELETESELECTEDITEM: OnDeleteSelectedItem(); return true;
	case ID_REMOVEONEOFF: OnRemoveOneOff(); return true;
	case ID_RTFSYNC: OnRTFSync(); return true;
	case ID_SAVEMESSAGETOFILE: OnSaveMessageToFile(); return true;
	case ID_SETREADFLAG: OnSetReadFlag(); return true;
	case ID_SETMESSAGESTATUS: OnSetMessageStatus(); return true;
	case ID_GETMESSAGEOPTIONS: OnGetMessageOptions(); return true;
	case ID_DELETEATTACHMENTS: OnDeleteAttachments(); return true;
	}
	return false;
} // CFolderDlg::MultiSelectComplex

/////////////////////////////////////////////////////////////////////////////
// CFolderDlg message handlers

void CFolderDlg::OnDisplayItem()
{
	(void) HandleMenu(ID_OPENNONMODAL);
} // CFolderDlg::OnDisplayItem

void CFolderDlg::OnInitMenu(_In_ CMenu* pMenu)
{
	if (pMenu && m_lpContentsTableListCtrl)
	{
		LPMAPISESSION	lpMAPISession = NULL;
		int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
		if (m_lpMapiObjects)
		{
			ULONG ulStatus = m_lpMapiObjects->GetBufferStatus();
			pMenu->EnableMenuItem(ID_PASTE,DIM(ulStatus & BUFFER_MESSAGES));
			lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		}

		pMenu->EnableMenuItem(ID_ADDTESTADDRESS,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_ATTACHMENTPROPERTIES,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_COPY,DIMMSOK(iNumSel));

		pMenu->EnableMenuItem(ID_DELETEATTACHMENTS,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_DELETESELECTEDITEM,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_MANUALRESOLVE,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_OPENMODAL,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_OPENNONMODAL,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_REMOVEONEOFF,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_RESENDMESSAGE,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_RTFSYNC,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_SAVEATTACHMENTS,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_SAVEMESSAGETOFILE,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_SETREADFLAG,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_SETMESSAGESTATUS,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_GETMESSAGESTATUS,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_GETMESSAGEOPTIONS,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_SUBMITMESSAGE,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_ABORTSUBMIT,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_RECIPIENTPROPERTIES,DIMMSOK(iNumSel));

		pMenu->EnableMenuItem(ID_GETPROPSUSINGLONGTERMEID,DIMMSNOK(iNumSel));
		pMenu->EnableMenuItem(ID_EXECUTEVERBONFORM,DIMMSNOK(iNumSel));

		pMenu->EnableMenuItem(ID_RESOLVEMESSAGECLASS,DIM(lpMAPISession));
		pMenu->EnableMenuItem(ID_SELECTFORM,DIM(lpMAPISession));

		pMenu->EnableMenuItem(ID_DISPLAYACLTABLE,DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_LOADFROMEML,DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_LOADFROMMSG,DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_LOADFROMTNEF,DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_NEW_APPOINTMENT,DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_NEW_CONTACT,DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_NEW_CUSTOMFORM,DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_NEW_IPMNOTE,DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_NEW_IPMPOST,DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_NEW_STICKYNOTE,DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_NEW_TASK,DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_NEW_MESSAGE,DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_SENDBULKMAIL,DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_SAVEFOLDERCONTENTSASTEXTFILES,DIM(m_lpContainer));

	}
	CContentsTableDlg::OnInitMenu(pMenu);
} // CFolderDlg::OnInitMenu

// Checks flags on add-in menu items to ensure they should be enabled
// Override to support context sensitive scenarios
void CFolderDlg::EnableAddInMenus(_In_ CMenu* pMenu, ULONG ulMenu, _In_ LPMENUITEM lpAddInMenu, UINT uiEnable)
{
	if (lpAddInMenu)
	{
		if (lpAddInMenu->ulFlags & MENU_FLAGS_FOLDER_ASSOC)
		{
			if (m_ulDisplayFlags & dfAssoc) uiEnable = MF_GRAYED;
		}
		else if (lpAddInMenu->ulFlags & MENU_FLAGS_DELETED)
		{
			if (m_ulDisplayFlags & dfDeleted) uiEnable = MF_GRAYED;
		}
		else if (lpAddInMenu->ulFlags & MENU_FLAGS_FOLDER_REG)
		{
			if (!(m_ulDisplayFlags == dfNormal)) uiEnable = MF_GRAYED;
		}
	}
	if (pMenu) pMenu->EnableMenuItem(ulMenu,uiEnable);
} // CFolderDlg::EnableAddInMenus

/////////////////////////////////////////////////////////////////////////////////////
//  Menu Commands

void CFolderDlg::OnAddOneOffAddress()
{
	HRESULT			hRes = S_OK;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;
	LPMAPISESSION	lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	CEditor MyData(
		this,
		IDS_ADDONEOFF,
		IDS_ADDONEOFFPROMPT,
		4,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	MyData.InitSingleLine(0,IDS_DISPLAYNAME,IDS_DISPLAYNAMEVALUE,false);
	MyData.InitSingleLineSz(1,IDS_ADDRESSTYPE,_T("EX"),false); // STRING_OK
	MyData.InitSingleLine(2,IDS_ADDRESS,IDS_ADDRESSVALUE,false);
	MyData.InitSingleLine(3,IDS_RECIPTYPE,NULL,false);
	MyData.SetHex(3,MAPI_TO);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		LPMESSAGE lpMessage = NULL;
		int iItem = m_lpContentsTableListCtrl->GetNextItem(
			-1,
			LVNI_SELECTED);
		while (iItem != -1)
		{
			EC_H(OpenItemProp(
				iItem,
				mfcmapiREQUEST_MODIFY,
				(LPMAPIPROP*)&lpMessage));

			if (lpMessage)
			{
				EC_H(AddOneOffAddress(
					lpMAPISession,
					lpMessage,
					MyData.GetString(0),
					MyData.GetString(1),
					MyData.GetString(2),
					MyData.GetHex(3)));

				lpMessage->Release();
				lpMessage = NULL;
			}
			iItem = m_lpContentsTableListCtrl->GetNextItem(
				iItem,
				LVNI_SELECTED);
			if (S_OK != hRes && -1 != iItem)
			{
				if (bShouldCancel(this,hRes)) break;
				hRes = S_OK;
			}
		}
	}
} // CFolderDlg::OnAddOneOffAddress

_Check_return_ HRESULT CFolderDlg::OnAttachmentProperties(int iItem, _In_ SortListData* /*lpData*/)
{
	HRESULT			hRes = S_OK;
	LPMESSAGE		lpMessage = NULL;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;

	EC_H(OpenItemProp(
		iItem,
		mfcmapiREQUEST_MODIFY,
		(LPMAPIPROP*)&lpMessage));

	if (lpMessage)
	{
		EC_H(OpenAttachmentsFromMessage(lpMessage, true));

		lpMessage->Release();
	}

	return hRes;
} // CFolderDlg::OnAttachmentProperties

void CFolderDlg::HandleCopy()
{
	HRESULT			hRes = S_OK;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric,CLASS,_T("HandleCopy"),_T("\n"));
	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

	LPENTRYLIST lpEIDs = NULL;

	EC_H(m_lpContentsTableListCtrl->GetSelectedItemEIDs(&lpEIDs));

	// m_lpMapiObjects takes over ownership of lpEIDs - don't free now
	m_lpMapiObjects->SetMessagesToCopy(lpEIDs,(LPMAPIFOLDER) m_lpContainer);
} // CFolderDlg::HandleCopy

_Check_return_ bool CFolderDlg::HandlePaste()
{
	if (CBaseDialog::HandlePaste()) return true;

	HRESULT		hRes = S_OK;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric,CLASS,_T("HandlePaste"),_T("\n"));
	if (!m_lpMapiObjects || !m_lpContainer) return false;

	CEditor MyData(
		this,
		IDS_COPYMESSAGE,
		IDS_COPYMESSAGEPROMPT,
		2,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	MyData.InitCheck(0,IDS_MESSAGEMOVE,false,false);
	UINT uidDropDown[] = {
		IDS_DDCOPYMESSAGES,
		IDS_DDCOPYTO
	};
	MyData.InitDropDown(1,IDS_COPYINTERFACE,_countof(uidDropDown),uidDropDown,true);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		LPENTRYLIST	lpEIDs = m_lpMapiObjects->GetMessagesToCopy();
		LPMAPIFOLDER lpMAPISourceFolder = m_lpMapiObjects->GetSourceParentFolder();
		ULONG		ulMoveMessage = MyData.GetCheck(0)?MESSAGE_MOVE:0;

		if (lpEIDs && lpMAPISourceFolder)
		{
			if (0 == MyData.GetDropDown(1))
			{ // CopyMessages
				LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IMAPIFolder::CopyMessages"), m_hWnd); // STRING_OK

				if(lpProgress)
					ulMoveMessage |= MESSAGE_DIALOG;

				EC_H(lpMAPISourceFolder->CopyMessages(
					lpEIDs,
					&IID_IMAPIFolder,
					m_lpContainer,
					lpProgress ? (ULONG_PTR)m_hWnd : NULL,
					lpProgress,
					ulMoveMessage));

				if(lpProgress)
					lpProgress->Release();

				lpProgress = NULL;
			}
			else
			{ // CopyTo
				// Specify properties to exclude in the copy operation. These are
				// the properties that Exchange excludes to save bits and time.
				// Should not be necessary to exclude these, but speeds the process
				// when a lot of messages are being copied.
				static const SizedSPropTagArray (7, excludeTags) =
				{
					7,
					PR_ACCESS,
					PR_BODY,
					PR_RTF_SYNC_BODY_COUNT,
					PR_RTF_SYNC_BODY_CRC,
					PR_RTF_SYNC_BODY_TAG,
					PR_RTF_SYNC_PREFIX_COUNT,
					PR_RTF_SYNC_TRAILING_COUNT
				};

				CTagArrayEditor MyEditor(
					this,
					IDS_TAGSTOEXCLUDE,
					IDS_TAGSTOEXCLUDEPROMPT,
					(LPSPropTagArray)&excludeTags,
					false,
					m_lpContainer);
				WC_H(MyEditor.DisplayDialog());

				if (S_OK == hRes)
				{
					ULONG i = 0;
					for (i = 0 ; i < lpEIDs->cValues ; i++)
					{
						LPMESSAGE lpMessage = NULL;

						EC_H(CallOpenEntry(
							NULL,
							NULL,
							lpMAPISourceFolder,
							NULL,
							lpEIDs->lpbin[i].cb,
							(LPENTRYID) lpEIDs->lpbin[i].lpb,
							NULL,
							MyData.GetCheck(0)?MAPI_MODIFY :0, // only need write access for moves
							NULL,
							(LPUNKNOWN*)&lpMessage));
						if (lpMessage)
						{
							LPMESSAGE lpNewMessage = NULL;
							EC_H(((LPMAPIFOLDER) m_lpContainer)->CreateMessage(NULL,(m_ulDisplayFlags & dfAssoc)?MAPI_ASSOCIATED:NULL,&lpNewMessage));
							if (lpNewMessage)
							{
								LPSPropProblemArray lpProblems = NULL;

								LPSPropTagArray lpTagsToExclude = (LPSPropTagArray)&excludeTags;
								LPSPropTagArray lpTagArray = MyEditor.DetachModifiedTagArray();
								if (lpTagArray)
								{
									lpTagsToExclude = lpTagArray;
								}
								// copy message properties to IMessage object opened on top of IStorage.
								LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IMAPIProp::CopyTo"), m_hWnd); // STRING_OK

								if(lpProgress)
									ulMoveMessage |= MAPI_DIALOG;

								EC_H(lpMessage->CopyTo(
									0,
									NULL, // TODO: interfaces?
									lpTagsToExclude,
									lpProgress ? (ULONG_PTR)m_hWnd : NULL, // UI param
									lpProgress, // progress
									(LPIID)&IID_IMessage,
									lpNewMessage,
									ulMoveMessage,
									&lpProblems));

								if(lpProgress)
									lpProgress->Release();

								lpProgress = NULL;

								EC_PROBLEMARRAY(lpProblems);
								MAPIFreeBuffer(lpProblems);
								MAPIFreeBuffer(lpTagArray);

								// save changes to IMessage object.
								EC_H(lpNewMessage->SaveChanges(KEEP_OPEN_READWRITE));
								lpNewMessage->Release();

								if (MyData.GetCheck(0)) // if we moved, save changes on original message
								{
									EC_H(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
								}
							}

							lpMessage->Release();
							lpMessage = NULL;
						}
					}
				}
			}
		}

		if (lpMAPISourceFolder) lpMAPISourceFolder->Release();
	}
	return true;
} // CFolderDlg::HandlePaste

void CFolderDlg::OnDeleteAttachments()
{
	HRESULT			hRes = S_OK;

	CEditor MyData(
		this,
		IDS_DELETEATTACHMENTS,
		IDS_DELETEATTACHMENTSPROMPT,
		1,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	MyData.InitSingleLine(0,IDS_FILENAME,NULL, false);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

		CString szAttName = MyData.GetString(0);
		LPCTSTR lpszAttName = NULL;

		if (!szAttName.IsEmpty())
		{
			lpszAttName = (LPCTSTR)szAttName;
		}

		int iItem = m_lpContentsTableListCtrl->GetNextItem(
			-1,
			LVNI_SELECTED);

		while (-1 != iItem)
		{
			LPMESSAGE		lpMessage = NULL;

			EC_H(OpenItemProp(
				iItem,
				mfcmapiREQUEST_MODIFY,
				(LPMAPIPROP*)&lpMessage));

			if (lpMessage)
			{
				EC_H(DeleteAttachments(
					lpMessage, lpszAttName, m_hWnd));

				lpMessage->Release();
			}

			iItem = m_lpContentsTableListCtrl->GetNextItem(
				iItem,
				LVNI_SELECTED);
		}
	}
} // CFolderDlg::OnDeleteAttachments

void CFolderDlg::OnDeleteSelectedItem()
{
	HRESULT			hRes = S_OK;
	LPMDB			lpMDB = NULL;
	CWaitCursor		Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl || !m_lpContainer) return;

	LPMAPISESSION	lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	EC_H(OpenDefaultMessageStore(lpMAPISession, &lpMDB));
	if (!lpMDB) return;

	bool	bMove = false;
	ULONG	ulFlag = MESSAGE_DIALOG;

	if (m_ulDisplayFlags & dfDeleted)
	{
		ulFlag |= DELETE_HARD_DELETE;
	}
	else
	{
		bool bShift = !(GetKeyState(VK_SHIFT) < 0);

		CEditor MyData(
			this,
			IDS_DELETEITEM,
			IDS_DELETEITEMPROMPT,
			1,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
		UINT uidDropDown[] = {
			IDS_DDDELETETODELETED,
			IDS_DDDELETETORETENTION,
			IDS_DDDELETEHARDDELETE
		};

		if (bShift)
			MyData.InitDropDown(0,IDS_DELSTYLE,_countof(uidDropDown),uidDropDown,true);
		else
			MyData.InitDropDown(0,IDS_DELSTYLE,_countof(uidDropDown) - 1,&uidDropDown[1],true);

		WC_H(MyData.DisplayDialog());

		if (bShift)
		{
			switch(MyData.GetDropDown(0))
			{
			case 0:
				bMove = true;
				break;
			case 2:
				ulFlag |= DELETE_HARD_DELETE;
				break;
			}
		}
		else
		{
			if (1 == MyData.GetDropDown(0)) ulFlag |= DELETE_HARD_DELETE;
		}
	}

	if (S_OK == hRes)
	{
		LPENTRYLIST lpEIDs = NULL;

		EC_H(m_lpContentsTableListCtrl->GetSelectedItemEIDs(&lpEIDs));

		if (bMove)
		{
			EC_H(DeleteToDeletedItems(
				lpMDB,
				(LPMAPIFOLDER) m_lpContainer,
				lpEIDs,
				m_hWnd));
		}
		else
		{
			LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IMAPIFolder::DeleteMessages"), m_hWnd); // STRING_OK

			if(lpProgress)
				ulFlag |= MESSAGE_DIALOG;

			EC_H(((LPMAPIFOLDER) m_lpContainer)->DeleteMessages(
				lpEIDs, // list of messages to delete
				lpProgress ? (ULONG_PTR)m_hWnd : NULL,
				lpProgress,
				ulFlag));

			if(lpProgress)
				lpProgress->Release();

			lpProgress = NULL;
		}
		MAPIFreeBuffer(lpEIDs);
	}

	lpMDB->Release();
} // CFolderDlg::OnDeleteSelectedItem

void CFolderDlg::OnGetPropsUsingLongTermEID()
{
	HRESULT			hRes = S_OK;
	LPSBinary		lpMessageEID = NULL;
	int				iItem = -1;
	SortListData*	lpListData = NULL;
	CWaitCursor		Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl) return;

	hRes = S_OK;
	lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);

	if (lpListData)
	{
		lpMessageEID = lpListData->data.Contents.lpLongtermID;

		if (lpMessageEID)
		{
			LPMAPIPROP lpMAPIProp = NULL;

			LPMDB		lpMDB = m_lpMapiObjects->GetMDB(); // do not release
			if (lpMDB)
			{
				WC_H(CallOpenEntry(
					lpMDB,
					NULL,
					NULL,
					NULL,
					lpMessageEID->cb,
					(LPENTRYID) lpMessageEID->lpb,
					NULL,
					MAPI_BEST_ACCESS,
					NULL,
					(LPUNKNOWN*)&lpMAPIProp));
			}

			OnUpdateSingleMAPIPropListCtrl(lpMAPIProp, NULL);

			if (lpMAPIProp) lpMAPIProp->Release();
		}
	}
} // CFolderDlg::OnGetPropsUsingLongTermEID

// Use CFileDialogExW to locate a .MSG file to load,
// Pass the file name and a message to load in to LoadFromMsg to do the work.
void CFolderDlg::OnLoadFromMSG()
{
	if (!m_lpContainer) return;

	HRESULT			hRes = S_OK;
	LPMESSAGE		lpNewMessage = NULL;
	INT_PTR			iDlgRet = IDOK;

	CStringW szFileSpec;
	EC_B(szFileSpec.LoadString(IDS_MSGFILES));

	CFileDialogExW dlgFilePicker;
	EC_D_DIALOG(dlgFilePicker.DisplayDialog(
		true,
		L"msg", // STRING_OK
		NULL,
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT | OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT,
		szFileSpec,
		this));

	if (iDlgRet == IDOK)
	{
		CEditor MyData(
			this,
			IDS_LOADMSG,
			IDS_LOADMSGPROMPT,
			1,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

		UINT uidDropDown[] = {
			IDS_DDLOADTOFOLDER,
			IDS_DDDISPLAYPROPSONLY
		};
		MyData.InitDropDown(0,IDS_LOADSTYLE,_countof(uidDropDown),uidDropDown,true);

		WC_H(MyData.DisplayDialog());
		if (S_OK == hRes)
		{
			LPWSTR lpszPath = NULL;
			while (NULL != (lpszPath = dlgFilePicker.GetNextFileName()))
			{
				hRes = S_OK;
				switch (MyData.GetDropDown(0))
				{
				case 0:
					EC_H(((LPMAPIFOLDER)m_lpContainer)->CreateMessage(
						NULL,
						(m_ulDisplayFlags & dfAssoc)?MAPI_ASSOCIATED:NULL,
						&lpNewMessage));

					if (lpNewMessage)
					{
						EC_H(LoadFromMSG(
							lpszPath,
							lpNewMessage, m_hWnd));
					}
					break;
				case 1:
					if (m_lpPropDisplay)
					{
						EC_H(LoadMSGToMessage(
							lpszPath,
							&lpNewMessage));

						if (lpNewMessage)
						{
							EC_H(m_lpPropDisplay->SetDataSource(lpNewMessage,NULL,false));
						}
					}
					break;
				}

				if (lpNewMessage)
				{
					lpNewMessage->Release();
					lpNewMessage = NULL;
				}
			}
		}
	}
} // CFolderDlg::OnLoadFromMSG

void CFolderDlg::OnResolveMessageClass()
{
	if (!m_lpMapiObjects) return;

	LPMAPIFORMINFO lpMAPIFormInfo = NULL;
	ResolveMessageClass(m_lpMapiObjects, (LPMAPIFOLDER)m_lpContainer, &lpMAPIFormInfo);
	if (lpMAPIFormInfo)
	{
		OnUpdateSingleMAPIPropListCtrl(lpMAPIFormInfo, NULL);
		lpMAPIFormInfo->Release();
	}
} // CFolderDlg::OnResolveMessageClass

void CFolderDlg::OnSelectForm()
{
	LPMAPIFORMINFO lpMAPIFormInfo = NULL;

	if (!m_lpMapiObjects) return;

	SelectForm(m_lpMapiObjects, (LPMAPIFOLDER)m_lpContainer, &lpMAPIFormInfo);
	if (lpMAPIFormInfo)
	{
		OnUpdateSingleMAPIPropListCtrl(lpMAPIFormInfo, NULL);
		// TODO: Put some code in here which works with the returned Form Info pointer
		lpMAPIFormInfo->Release();
	}
} // CFolderDlg::OnSelectForm

void CFolderDlg::OnManualResolve()
{
	HRESULT			hRes = S_OK;
	LPMESSAGE		lpMessage = NULL;
	int				iItem = -1;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

	LPMAPISESSION	lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	CPropertyTagEditor MyPropertyTag(
		IDS_MANUALRESOLVE,
		NULL, // prompt
		NULL,
		m_bIsAB,
		m_lpContainer,
		this);

	WC_H(MyPropertyTag.DisplayDialog());
	if (S_OK == hRes)
	{
		CEditor MyData(
			this,
			IDS_MANUALRESOLVE,
			IDS_MANUALRESOLVEPROMPT,
			1,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

		MyData.InitSingleLine(0,IDS_DISPLAYNAME,NULL,false);

		WC_H(MyData.DisplayDialog());
		if (S_OK == hRes)
		{
			do
			{
				hRes = S_OK;
				EC_H(m_lpContentsTableListCtrl->OpenNextSelectedItemProp(
					&iItem,
					mfcmapiREQUEST_MODIFY,
					(LPMAPIPROP*)&lpMessage));

				if (lpMessage)
				{
					EC_H(ManualResolve(
						lpMAPISession,
						lpMessage,
						MyData.GetString(0),
						CHANGE_PROP_TYPE(MyPropertyTag.GetPropertyTag(),PT_TSTRING)));

					lpMessage->Release();
					lpMessage = NULL;
				}
			}
			while (iItem != -1);
		}
	}
} // CFolderDlg::OnManualResolve

void CFolderDlg::NewSpecialItem(WORD wMenuSelect)
{
	HRESULT			hRes = S_OK;
	LPMAPIFOLDER	lpFolder = NULL; // DO NOT RELEASE

	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

	LPMDB lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release

	if (lpMAPISession)
	{
		ULONG ulTag = NULL;
		LPCSTR szClass = NULL;

		switch(wMenuSelect)
		{
		case ID_NEW_APPOINTMENT:
			ulTag = PR_IPM_APPOINTMENT_ENTRYID;
			szClass = "IPM.APPOINTMENT"; // STRING_OK
			break;
		case ID_NEW_CONTACT:
			ulTag = PR_IPM_CONTACT_ENTRYID;
			szClass = "IPM.CONTACT"; // STRING_OK
			break;
		case ID_NEW_IPMNOTE:
			szClass = "IPM.NOTE"; // STRING_OK
			break;
		case ID_NEW_IPMPOST:
			szClass = "IPM.POST"; // STRING_OK
			break;
		case ID_NEW_TASK:
			ulTag = PR_IPM_TASK_ENTRYID;
			szClass = "IPM.TASK"; // STRING_OK
			break;
		case ID_NEW_STICKYNOTE:
			ulTag = PR_IPM_NOTE_ENTRYID;
			szClass = "IPM.STICKYNOTE"; // STRING_OK
			break;

		}
		LPMAPIFOLDER lpSpecialFolder = NULL;
		if (ulTag)
		{
			EC_H(GetSpecialFolder(lpMDB,ulTag,&lpSpecialFolder));
			lpFolder = lpSpecialFolder;
		}
		else
		{
			lpFolder = (LPMAPIFOLDER) m_lpContainer;
		}

		if (lpFolder)
		{
			EC_H(CreateAndDisplayNewMailInFolder(
				m_hWnd,
				lpMDB,
				lpMAPISession,
				m_lpContentsTableListCtrl,
				-1,
				szClass,
				lpFolder));
		}
		if (lpSpecialFolder) lpSpecialFolder->Release();
	}
} // CFolderDlg::NewSpecialItem

void CFolderDlg::OnNewMessage()
{
	HRESULT	hRes = S_OK;
	LPMESSAGE lpMessage = NULL;

	EC_H(((LPMAPIFOLDER)m_lpContainer)->CreateMessage(
		NULL,
		m_ulDisplayFlags & dfAssoc?MAPI_ASSOCIATED:0,
		&lpMessage));

	if (lpMessage)
	{
		EC_H(lpMessage->SaveChanges(NULL));
		lpMessage->Release();
	}
} // CFolderDlg::OnNewMessage

void CFolderDlg::OnNewCustomForm()
{
	HRESULT			hRes = S_OK;

	if (!m_lpMapiObjects || !m_lpContainer || !m_lpContentsTableListCtrl) return;

	LPMDB lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release

	if (lpMAPISession)
	{
		CEditor MyPrompt1(
			this,
			IDS_NEWCUSTOMFORM,
			IDS_NEWCUSTOMFORMPROMPT1,
			1,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
		UINT uidDropDown[] = {
			IDS_DDENTERFORMCLASS,
			IDS_DDFOLDERFORMLIBRARY,
			IDS_DDORGFORMLIBRARY
		};
		MyPrompt1.InitDropDown(0,IDS_LOCATIONOFFORM,_countof(uidDropDown),uidDropDown,true);

		WC_H(MyPrompt1.DisplayDialog());

		if (S_OK == hRes)
		{
			LPCSTR			szClass = NULL;
			LPSPropValue	lpProp = NULL;

			CEditor MyClass(
				this,
				IDS_NEWCUSTOMFORM,
				IDS_NEWCUSTOMFORMPROMPT2,
				1,
				CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
			MyClass.InitSingleLineSz(0,IDS_FORMTYPE,_T("IPM.Note"),false); // STRING_OK

			switch (MyPrompt1.GetDropDown(0))
			{
			case 0:
				WC_H(MyClass.DisplayDialog());

				if (S_OK == hRes)
				{
					szClass = MyClass.GetStringA(0);
				}
				break;
			case 1:
			case 2:
				{
					LPMAPIFORMMGR	lpMAPIFormMgr = NULL;
					LPMAPIFORMINFO	lpMAPIFormInfo = NULL;
					EC_H(MAPIOpenFormMgr(lpMAPISession,&lpMAPIFormMgr));

					if (lpMAPIFormMgr)
					{
						LPMAPIFOLDER lpFormSource = 0;
						if (1 == MyPrompt1.GetDropDown(0)) lpFormSource = (LPMAPIFOLDER) m_lpContainer;
						// Apparently, SelectForm doesn't support unicode
						// CString doesn't provide a way to extract just ANSI strings, so we do this manually
						CHAR szTitle[256];
						int iRet = NULL;
						EC_D(iRet,LoadStringA(GetModuleHandle(NULL),
							IDS_SELECTFORMCREATE,
							szTitle,
							_countof(szTitle)));
#pragma warning(push)
#pragma warning(disable:4616)
#pragma warning(disable:6276)
						EC_H_CANCEL(lpMAPIFormMgr->SelectForm(
							(ULONG_PTR)m_hWnd,
							0, // fMapiUnicode,
							(LPCTSTR) szTitle,
							lpFormSource,
							&lpMAPIFormInfo));
#pragma warning(pop)

						if (lpMAPIFormInfo)
						{
							EC_H(HrGetOneProp(
								lpMAPIFormInfo,
								PR_MESSAGE_CLASS_A,
								&lpProp));
							if (CheckStringProp(lpProp,PT_STRING8))
							{
								szClass = lpProp->Value.lpszA;
							}
							lpMAPIFormInfo->Release();
						}
						lpMAPIFormMgr->Release();
					}
				}
				break;
			}

			EC_H(CreateAndDisplayNewMailInFolder(
				m_hWnd,
				lpMDB,
				lpMAPISession,
				m_lpContentsTableListCtrl,
				-1,
				szClass,
				(LPMAPIFOLDER) m_lpContainer));

			MAPIFreeBuffer(lpProp);
		}
	}
} // CFolderDlg::OnNewCustomForm

_Check_return_ HRESULT CFolderDlg::OnOpenModal(int iItem, _In_ SortListData* /*lpData*/)
{
	HRESULT			hRes = S_OK;
	LPMESSAGE		lpMessage = NULL;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;
	if (!m_lpMapiObjects || !m_lpContainer) return MAPI_E_INVALID_PARAMETER;

	LPMDB lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return MAPI_E_CALL_FAILED;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (lpMAPISession)
	{
		// Before we open the message, make sure the MAPI Form Manager is implemented
		LPMAPIFORMMGR lpFormMgr = NULL;
		WC_H(MAPIOpenFormMgr(lpMAPISession,&lpFormMgr));
		hRes = S_OK; // Ditch the error if we got one

		if (lpFormMgr)
		{
			EC_H(OpenItemProp(
				iItem,
				mfcmapiREQUEST_MODIFY,
				(LPMAPIPROP*)&lpMessage));
			if (lpMessage)
			{
				EC_H(OpenMessageModal(
					(LPMAPIFOLDER) m_lpContainer,
					lpMAPISession,
					lpMDB,
					lpMessage));

				lpMessage->Release();
			}
			lpFormMgr->Release();
		}
	}
	return hRes;
} // CFolderDlg::OnOpenModal

_Check_return_ HRESULT CFolderDlg::OnOpenNonModal(int iItem, _In_ SortListData* /*lpData*/)
{
	HRESULT			hRes = S_OK;
	LPMESSAGE		lpMessage = NULL;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;
	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl || !m_lpContainer) return MAPI_E_INVALID_PARAMETER;

	LPMDB lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return MAPI_E_CALL_FAILED;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (lpMAPISession)
	{
		// Before we open the message, make sure the MAPI Form Manager is implemented
		LPMAPIFORMMGR lpFormMgr = NULL;
		WC_H(MAPIOpenFormMgr(lpMAPISession,&lpFormMgr));
		hRes = S_OK; // Ditch the error if we got one

		if (lpFormMgr)
		{
			EC_H(OpenItemProp(
				iItem,
				mfcmapiREQUEST_MODIFY,
				(LPMAPIPROP*)&lpMessage));

			if (lpMessage)
			{
				WC_H(OpenMessageNonModal(
					m_hWnd,
					lpMDB,
					lpMAPISession,
					(LPMAPIFOLDER) m_lpContainer,
					m_lpContentsTableListCtrl,
					iItem,
					lpMessage,
					EXCHIVERB_OPEN,
					NULL));

				lpMessage->Release();
			}
			lpFormMgr->Release();
		}
	}
	return hRes;
} // CFolderDlg::OnOpenNonModal

void CFolderDlg::OnExecuteVerbOnForm()
{
	HRESULT			hRes = S_OK;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

	LPMDB lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (lpMAPISession)
	{
		CEditor MyData(
			this,
			IDS_EXECUTEVERB,
			IDS_EXECUTEVERBPROMPT,
			1,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
		MyData.SetPromptPostFix(AllFlagsToString(PROP_ID(PR_LAST_VERB_EXECUTED),false));

		MyData.InitSingleLine(0,IDS_VERB,NULL,false);
		MyData.SetDecimal(0,EXCHIVERB_OPEN);

		WC_H(MyData.DisplayDialog());
		if (S_OK == hRes)
		{
			LPMESSAGE lpMessage = NULL;
			int iItem = m_lpContentsTableListCtrl->GetNextItem(
				-1,
				LVNI_SELECTED);
			if (iItem != -1)
			{
				EC_H(OpenItemProp(
					iItem,
					mfcmapiREQUEST_MODIFY,
					(LPMAPIPROP*)&lpMessage));

				if (lpMessage)
				{
					EC_H(OpenMessageNonModal(
						m_hWnd,
						lpMDB,
						lpMAPISession,
						(LPMAPIFOLDER) m_lpContainer,
						m_lpContentsTableListCtrl,
						iItem,
						lpMessage,
						MyData.GetDecimal(0),
						NULL));

					lpMessage->Release();
					lpMessage = NULL;
				}
			}
		}
	}
} // CFolderDlg::OnExecuteVerbOnForm

_Check_return_ HRESULT CFolderDlg::OnResendSelectedItem(int /*iItem*/, _In_ SortListData* lpData)
{
	HRESULT			hRes = S_OK;
	CWaitCursor		Wait; // Change the mouse to an hourglass while we work.

	if (!lpData || !m_lpContainer) return MAPI_E_INVALID_PARAMETER;

	if (lpData->data.Contents.lpEntryID)
	{
		EC_H(ResendSingleMessage(
			(LPMAPIFOLDER) m_lpContainer,
			lpData->data.Contents.lpEntryID,
			m_hWnd));
	}
	return hRes;
} // CFolderDlg::OnResendSelectedItem

_Check_return_ HRESULT CFolderDlg::OnRecipientProperties(int iItem, _In_ SortListData* /*lpData*/)
{
	HRESULT			hRes = S_OK;
	LPMESSAGE		lpMessage = NULL;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;

	EC_H(OpenItemProp(
		iItem,
		mfcmapiREQUEST_MODIFY,
		(LPMAPIPROP*)&lpMessage));

	if (lpMessage)
	{
		EC_H(OpenRecipientsFromMessage(lpMessage));

		lpMessage->Release();
	}

	return hRes;
} // CFolderDlg::OnRecipientProperties

void CFolderDlg::OnRemoveOneOff()
{
	HRESULT			hRes = S_OK;
	CWaitCursor		Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl) return;

	CEditor MyData(
		this,
		IDS_REMOVEONEOFF,
		IDS_REMOVEONEOFFPROMPT,
		1,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
	MyData.InitCheck(0,IDS_REMOVEPROPDEFSTREAM,true,false);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		LPMESSAGE	lpMessage = NULL;

		int iItem = m_lpContentsTableListCtrl->GetNextItem(
			-1,
			LVNI_SELECTED);
		while (iItem != -1)
		{
			EC_H(OpenItemProp(
				iItem,
				mfcmapiREQUEST_MODIFY,
				(LPMAPIPROP*)&lpMessage));

			if (lpMessage)
			{
				DebugPrint(DBGGeneric, _T("Calling RemoveOneOff on %p, %sremoving property definition stream\n"),lpMessage,MyData.GetCheck(0)?_T(""):_T("not "));
				EC_H(RemoveOneOff(
					lpMessage,
					MyData.GetCheck(0)));

				lpMessage->Release();
				lpMessage = NULL;
			}
			iItem = m_lpContentsTableListCtrl->GetNextItem(
				iItem,
				LVNI_SELECTED);
			if (S_OK != hRes && -1 != iItem)
			{
				if (bShouldCancel(this,hRes)) break;
				hRes = S_OK;
			}
		}
	}
} // CFolderDlg::OnRemoveOneOff

#define RTF_SYNC_HTML_CHANGED ((ULONG) 0x00000004)

void CFolderDlg::OnRTFSync()
{
	HRESULT			hRes = S_OK;
	CWaitCursor		Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl) return;

	CEditor MyData(
		this,
		IDS_CALLRTFSYNC,
		IDS_CALLRTFSYNCPROMPT,
		1,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	MyData.InitSingleLine(0,IDS_FLAGS,NULL,false);
	MyData.SetHex(0,RTF_SYNC_RTF_CHANGED);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		LPMESSAGE	lpMessage = NULL;
		BOOL		bMessageUpdated = false;

		int iItem = m_lpContentsTableListCtrl->GetNextItem(
			-1,
			LVNI_SELECTED);
		while (iItem != -1)
		{
			EC_H(OpenItemProp(
				iItem,
				mfcmapiREQUEST_MODIFY,
				(LPMAPIPROP*)&lpMessage));

			if (lpMessage)
			{
				DebugPrint(DBGGeneric, _T("Calling RTFSync on %p with flags 0x%X\n"),lpMessage,MyData.GetHex(0));
				EC_H(RTFSync(
					lpMessage,
					MyData.GetHex(0),
					&bMessageUpdated));

				DebugPrint(DBGGeneric, _T("RTFSync returned %d\n"),bMessageUpdated);

				EC_H(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));

				lpMessage->Release();
				lpMessage = NULL;
			}
			iItem = m_lpContentsTableListCtrl->GetNextItem(
				iItem,
				LVNI_SELECTED);
			if (S_OK != hRes && -1 != iItem)
			{
				if (bShouldCancel(this,hRes)) break;
				hRes = S_OK;
			}
		}
	}
} // CFolderDlg::OnRTFSync

_Check_return_ HRESULT CFolderDlg::OnSaveAttachments(int iItem, _In_ SortListData* /*lpData*/)
{
	HRESULT			hRes = S_OK;
	LPMESSAGE		lpMessage = NULL;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;

	EC_H(OpenItemProp(
		iItem,
		mfcmapiREQUEST_MODIFY,
		(LPMAPIPROP*)&lpMessage));

	if (lpMessage)
	{
		EC_H(WriteAttachmentsToFile(
			lpMessage, m_hWnd));

		lpMessage->Release();
	}
	return hRes;
} // CFolderDlg::OnSaveAttachments

void CFolderDlg::OnSaveFolderContentsAsTextFiles()
{
	HRESULT			hRes = S_OK;
	WCHAR			szDir[MAX_PATH];

	if (!m_lpMapiObjects || !m_lpContainer) return;

	LPMDB lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return;

	szDir[0] = 0;
	WC_H(GetDirectoryPath(szDir));

	if (S_OK == hRes && szDir[0])
	{
		CWaitCursor		Wait; // Change the mouse to an hourglass while we work.

		CDumpStore MyDumpStore;
		MyDumpStore.InitMDB(lpMDB);
		MyDumpStore.InitFolder((LPMAPIFOLDER) m_lpContainer);
		MyDumpStore.InitFolderPathRoot(szDir);
		MyDumpStore.ProcessFolders(
			(m_ulDisplayFlags & dfAssoc)?false:true,
			(m_ulDisplayFlags & dfAssoc)?true:false,
			false);
	}
} // CFolderDlg::OnSaveFolderContentsAsTextFiles

void CFolderDlg::OnSaveMessageToFile()
{
	HRESULT	hRes = S_OK;

	DebugPrintEx(DBGGeneric,CLASS,_T("OnSaveMessageToFile"),_T("\n"));

	CEditor MyData(
		this,
		IDS_SAVEMESSAGETOFILE,
		IDS_SAVEMESSAGETOFILEPROMPT,
		1,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	UINT uidDropDown[] = {
		IDS_DDTEXTFILE,
		IDS_DDMSGFILEANSI,
		IDS_DDMSGFILEUNICODE,
		IDS_DDEMLFILE,
		IDS_DDEMLFILEUSINGICONVERTERSESSION,
		IDS_DDTNEFFILE
	};
	MyData.InitDropDown(0,IDS_FORMATTOSAVEMESSAGE,_countof(uidDropDown),uidDropDown,true);
	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		LPCWSTR		szExt = NULL;
		LPCWSTR		szDotExt = NULL;
		ULONG		ulDotExtLen = NULL;
		CStringW	szFilter;
		LPADRBOOK	lpAddrBook = NULL;
		switch(MyData.GetDropDown(0))
		{
		case 0:
			szExt = L"xml"; // STRING_OK
			szDotExt = L".xml"; // STRING_OK
			ulDotExtLen = 4;
			EC_B(szFilter.LoadString(IDS_XMLFILES));
			break;
		case 1:
		case 2:
			szExt = L"msg"; // STRING_OK
			szDotExt = L".msg"; // STRING_OK
			ulDotExtLen = 4;
			EC_B(szFilter.LoadString(IDS_MSGFILES));
			break;
		case 3:
		case 4:
			szExt = L"eml"; // STRING_OK
			szDotExt = L".eml"; // STRING_OK
			ulDotExtLen = 4;
			EC_B(szFilter.LoadString(IDS_EMLFILES));
			break;
		case 5:
			szExt = L"tnef"; // STRING_OK
			szDotExt = L".tnef"; // STRING_OK
			ulDotExtLen = 5;
			EC_B(szFilter.LoadString(IDS_TNEFFILES));

			lpAddrBook = m_lpMapiObjects->GetAddrBook(true); // do not release
			break;
		default:
			break;
		}
		int iItem = m_lpContentsTableListCtrl->GetNextItem(
			-1,
			LVNI_SELECTED);
		while (-1 != iItem)
		{
			LPMESSAGE	lpMessage = NULL;
			WCHAR		szFileName[MAX_PATH] = {0};
			INT_PTR		iDlgRet = IDOK;

			EC_H(OpenItemProp(
				iItem,
				mfcmapiREQUEST_MODIFY,
				(LPMAPIPROP*)&lpMessage));

			if (lpMessage)
			{
				WC_H(BuildFileName(szFileName,_countof(szFileName),szDotExt,ulDotExtLen,lpMessage));

				CFileDialogExW dlgFilePicker;

				EC_D_DIALOG(dlgFilePicker.DisplayDialog(
					false,
					szExt,
					szFileName,
					OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
					szFilter,
					this));

				if (iDlgRet == IDOK)
				{
					switch(MyData.GetDropDown(0))
					{
					case 0:
						// Idea is to capture anything that may be important about this message to disk so it can be analyzed.
						{
							CDumpStore MyDumpStore;
							MyDumpStore.InitMessagePath(dlgFilePicker.GetFileName());
							// Just assume this message might have attachments
							MyDumpStore.ProcessMessage(lpMessage,true,NULL);
						}
						break;
					case 1:
						EC_H(SaveToMSG(lpMessage,dlgFilePicker.GetFileName(),false,m_hWnd));
						break;
					case 2:
						EC_H(SaveToMSG(lpMessage,dlgFilePicker.GetFileName(),true,m_hWnd));
						break;
					case 3:
						EC_H(SaveToEML(lpMessage,dlgFilePicker.GetFileName()));
						break;
					case 4:
						{
							ULONG ulConvertFlags = CCSF_SMTP;
							ENCODINGTYPE et = IET_UNKNOWN;
							MIMESAVETYPE mst = USE_DEFAULT_SAVETYPE;
							ULONG ulWrapLines = USE_DEFAULT_WRAPPING;
							bool bDoAdrBook = false;

							EC_H(GetConversionToEMLOptions(this,&ulConvertFlags,&et,&mst,&ulWrapLines,&bDoAdrBook));
							if (S_OK == hRes)
							{
								LPADRBOOK lpAdrBook = NULL;
								if (bDoAdrBook) lpAdrBook = m_lpMapiObjects->GetAddrBook(true); // do not release

								EC_H(ExportIMessageToEML(
									lpMessage,
									dlgFilePicker.GetFileName(),
									ulConvertFlags,
									et,
									mst,
									ulWrapLines,
									lpAdrBook));
							}
						}
						break;
					case 5:
						EC_H(SaveToTNEF(lpMessage,lpAddrBook,dlgFilePicker.GetFileName()));
						break;
					default:
						break;
					}
				}
				else if (iDlgRet == IDCANCEL)
				{
					hRes = MAPI_E_USER_CANCEL;
				}

				lpMessage->Release();
			}
			else
			{
				hRes = MAPI_E_USER_CANCEL;
				CHECKHRESMSG(hRes,IDS_OPENMSGFAILED);
			}

			iItem = m_lpContentsTableListCtrl->GetNextItem(
				iItem,
				LVNI_SELECTED);
			if ((IDOK != iDlgRet || S_OK != hRes) && -1 != iItem)
			{
				if (bShouldCancel(this,hRes)) break;
				hRes = S_OK;
			}
		}
	}
} // CFolderDlg::OnSaveMessageToFile

// Use CFileDialogExW to locate a .DAT or .TNEF file to load,
// Pass the file name and a message to load in to LoadFromTNEF to do the work.
void CFolderDlg::OnLoadFromTNEF()
{
	HRESULT			hRes = S_OK;
	LPMESSAGE		lpNewMessage = NULL;
	INT_PTR			iDlgRet = IDOK;

	if (!m_lpContainer) return;

	LPADRBOOK lpAddrBook = m_lpMapiObjects->GetAddrBook(true); // do not release
	if (lpAddrBook)
	{
		CStringW szFileSpec;
		EC_B(szFileSpec.LoadString(IDS_TNEFFILES));

		CFileDialogExW dlgFilePicker;
		EC_D_DIALOG(dlgFilePicker.DisplayDialog(
			true,
			L"tnef", // STRING_OK
			NULL,
			OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT | OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT,
			szFileSpec,
			this));

		if (iDlgRet == IDOK)
		{
			LPWSTR lpszPath = NULL;
			while (NULL != (lpszPath = dlgFilePicker.GetNextFileName()))
			{
				hRes = S_OK;
				EC_H(((LPMAPIFOLDER)m_lpContainer)->CreateMessage(
					NULL,
					(m_ulDisplayFlags & dfAssoc)?MAPI_ASSOCIATED:0,
					&lpNewMessage));

				if (lpNewMessage)
				{
					EC_H(LoadFromTNEF(
						lpszPath,
						lpAddrBook,
						lpNewMessage));

					lpNewMessage->Release();
					lpNewMessage = NULL;
				}
			}
		}
	}

	if (lpNewMessage) lpNewMessage->Release();
} // CFolderDlg::OnLoadFromTNEF

// Use CFileDialogExW to locate a .EML file to load,
// Pass the file name and a message to load in to ImportEMLToIMessage to do the work.
void CFolderDlg::OnLoadFromEML()
{
	if (!m_lpContainer) return;

	HRESULT			hRes = S_OK;
	LPMESSAGE		lpNewMessage = NULL;
	INT_PTR			iDlgRet = IDOK;

	ULONG ulConvertFlags = CCSF_SMTP;
	bool bDoAdrBook = false;
	bool bDoApply = false;
	HCHARSET hCharSet = NULL;
	CSETAPPLYTYPE cSetApplyType = CSET_APPLY_UNTAGGED;
	WC_H(GetConversionFromEMLOptions(this,&ulConvertFlags,&bDoAdrBook,&bDoApply,&hCharSet,&cSetApplyType,NULL));
	if (S_OK == hRes)
	{
		LPADRBOOK lpAdrBook = NULL;
		if (bDoAdrBook) lpAdrBook = m_lpMapiObjects->GetAddrBook(true); // do not release

		CFileDialogExW dlgFilePicker;
		CStringW szFileSpec;

		EC_B(szFileSpec.LoadString(IDS_EMLFILES));
		EC_D_DIALOG(dlgFilePicker.DisplayDialog(
			true,
			L"eml", // STRING_OK
			NULL,
			OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT | OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT,
			szFileSpec,
			this));

		if (iDlgRet == IDOK)
		{
			LPWSTR lpszPath = NULL;
			while (NULL != (lpszPath = dlgFilePicker.GetNextFileName()))
			{
				hRes = S_OK;
				EC_H(((LPMAPIFOLDER)m_lpContainer)->CreateMessage(
					NULL,
					(m_ulDisplayFlags & dfAssoc)?MAPI_ASSOCIATED:0,
					&lpNewMessage));

				if (lpNewMessage)
				{
					EC_H(ImportEMLToIMessage(
						lpszPath,
						lpNewMessage,
						ulConvertFlags,
						bDoApply,
						hCharSet,
						cSetApplyType,
						lpAdrBook));

					lpNewMessage->Release();
					lpNewMessage = NULL;
				}
			}
		}
	}
} // CFolderDlg::OnLoadFromEML

void CFolderDlg::OnSendBulkMail()
{
	HRESULT			hRes = S_OK;

	CEditor MyData(
		this,
		IDS_SENDBULKMAIL,
		IDS_SENDBULKMAILPROMPT,
		4,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
	MyData.InitSingleLine(0,IDS_NUMMESSAGES,NULL,false);
	MyData.InitSingleLine(1,IDS_RECIPNAME,NULL,false);
	MyData.InitSingleLine(2,IDS_SUBJECT,NULL,false);
	MyData.InitMultiLine(3,IDS_BODY,NULL,false);

	if (!m_lpContainer) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		ULONG ulNumMessages = 0;
		ULONG i =0;
		LPTSTR szSubject = NULL;

		ulNumMessages = MyData.GetDecimal(0);
		szSubject = MyData.GetString(2);

		for (i = 0 ; i < ulNumMessages ; i++)
		{
			hRes = S_OK;
			CString szTestSubject;
			szTestSubject.FormatMessage(IDS_TESTSUBJECT,szSubject,i);

			EC_H(SendTestMessage(
				lpMAPISession,
				(LPMAPIFOLDER) m_lpContainer,
				MyData.GetString(1),
				MyData.GetString(3),
				szTestSubject));
			if (FAILED(hRes))
			{
				CHECKHRESMSG(hRes,IDS_ERRORSENDINGMSGS);
				break;
			}
		}
	}
} // CFolderDlg::OnSendBulkMail

void CFolderDlg::OnSetReadFlag()
{
	HRESULT			hRes = S_OK;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl) return;

	DebugPrintEx(DBGGeneric,CLASS,_T("OnSetReadFlag"),_T("\n"));

	CEditor MyFlags(
		this,
		IDS_SETREADFLAG,
		IDS_SETREADFLAGPROMPT,
		1,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	MyFlags.InitSingleLine(0,IDS_FLAGSINHEX,NULL,false);
	MyFlags.SetHex(0,CLEAR_READ_FLAG);

	WC_H(MyFlags.DisplayDialog());
	if (S_OK == hRes)
	{
		int iNumSelected = m_lpContentsTableListCtrl->GetSelectedCount();

		if (1 == iNumSelected)
		{
			LPMESSAGE lpMessage = NULL;

			EC_H(m_lpContentsTableListCtrl->OpenNextSelectedItemProp(
				NULL,
				mfcmapiREQUEST_MODIFY,
				(LPMAPIPROP*)&lpMessage));

			if (lpMessage)
			{
				EC_H(lpMessage->SetReadFlag(MyFlags.GetHex(0)));
				lpMessage->Release();
				lpMessage = NULL;
			}
		}
		else if (iNumSelected > 1)
		{
			LPENTRYLIST lpEIDs = NULL;

			EC_H(m_lpContentsTableListCtrl->GetSelectedItemEIDs(&lpEIDs));

			LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IMAPIFolder::SetReadFlags"), m_hWnd); // STRING_OK

			ULONG ulFlags = MyFlags.GetHex(0);

			if(lpProgress)
				ulFlags |= MESSAGE_DIALOG;

			EC_H(((LPMAPIFOLDER)m_lpContainer)->SetReadFlags(
				lpEIDs,
				lpProgress ? (ULONG_PTR)m_hWnd : NULL,
				lpProgress,
				ulFlags));

			if(lpProgress)
				lpProgress->Release();

			lpProgress = NULL;

			MAPIFreeBuffer(lpEIDs);
		}
	}
} // CFolderDlg::OnSetReadFlag

void CFolderDlg::OnGetMessageOptions()
{
	HRESULT			hRes = S_OK;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	DebugPrintEx(DBGGeneric,CLASS,_T("OnGetMessageOptions"),_T("\n"));

	CEditor MyAddress(
		this,
		IDS_MESSAGEOPTIONS,
		IDS_ADDRESSTYPEPROMPT,
		1,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
	MyAddress.InitSingleLineSz(0,IDS_ADDRESSTYPE,_T("EX"),false); // STRING_OK
	WC_H(MyAddress.DisplayDialog());

	if (S_OK == hRes)
	{
		LPMESSAGE lpMessage = NULL;
		int iItem = m_lpContentsTableListCtrl->GetNextItem(
			-1,
			LVNI_SELECTED);
		while (iItem != -1)
		{
			EC_H(OpenItemProp(
				iItem,
				mfcmapiREQUEST_MODIFY,
				(LPMAPIPROP*)&lpMessage));

			if (lpMessage)
			{
				EC_H(lpMAPISession->MessageOptions(
					(ULONG_PTR) m_hWnd,
					NULL, // API doesn't like Unicode
					(LPTSTR)MyAddress.GetStringA(0),
					lpMessage));

				lpMessage->Release();
				lpMessage = NULL;
			}
			iItem = m_lpContentsTableListCtrl->GetNextItem(
				iItem,
				LVNI_SELECTED);
			if (S_OK != hRes && -1 != iItem)
			{
				if (bShouldCancel(this,hRes)) break;
				hRes = S_OK;
			}
		}
	}
} // CFolderDlg::OnGetMessageOptions

_Check_return_ HRESULT CFolderDlg::OnGetMessageStatus(int /*iItem*/, _In_ SortListData* lpData)
{
	HRESULT			hRes = S_OK;
	LPSBinary		lpMessageEID = NULL;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (!lpData || !m_lpContainer) return MAPI_E_INVALID_PARAMETER;

	DebugPrintEx(DBGGeneric,CLASS,_T("OnGetMessageStatus"),_T("\n"));

	ULONG ulMessageStatus = NULL;

	lpMessageEID = lpData->data.Contents.lpEntryID;

	if (lpMessageEID)
	{
		EC_H(((LPMAPIFOLDER)m_lpContainer)->GetMessageStatus(
			lpMessageEID->cb,
			(LPENTRYID) lpMessageEID->lpb,
			NULL,
			&ulMessageStatus));

		CEditor MyStatus(
			this,
			IDS_MESSAGESTATUS,
			IDS_MESSAGESTATUS,
			1,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
		MyStatus.InitSingleLine(0,IDS_MESSAGESTATUS,NULL,true);
		MyStatus.SetHex(0,ulMessageStatus);

		WC_H(MyStatus.DisplayDialog());
	}
	return hRes;
} // CFolderDlg::OnGetMessageStatus

void CFolderDlg::OnSetMessageStatus()
{
	HRESULT			hRes = S_OK;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl || !m_lpContainer) return;

	CEditor MyData(
		this,
		IDS_SETMSGSTATUS,
		IDS_SETMSGSTATUSPROMPT,
		2,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	MyData.InitSingleLine(0,IDS_STATUSINHEX,NULL,false);
	MyData.InitSingleLine(1,IDS_MASKINHEX,NULL,false);

	DebugPrintEx(DBGGeneric,CLASS,_T("OnSetMessageStatus"),_T("\n"));

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		SortListData*	lpListData = NULL;
		LPSBinary		lpMessageEID = NULL;
		int iItem = m_lpContentsTableListCtrl->GetNextItem(
			-1,
			LVNI_SELECTED);
		while (iItem != -1)
		{
			lpListData = (SortListData*) m_lpContentsTableListCtrl->GetItemData(iItem);
			if (lpListData)
			{
				lpMessageEID = lpListData->data.Contents.lpEntryID;

				if (lpMessageEID)
				{
					ULONG ulOldStatus = NULL;

					EC_H(((LPMAPIFOLDER)m_lpContainer)->SetMessageStatus(
						lpMessageEID->cb,
						(LPENTRYID) lpMessageEID->lpb,
						MyData.GetHex(0),
						MyData.GetHex(1),
						&ulOldStatus));
				}
			}
			iItem = m_lpContentsTableListCtrl->GetNextItem(
				iItem,
				LVNI_SELECTED);
			if (S_OK != hRes && -1 != iItem)
			{
				if (bShouldCancel(this,hRes)) break;
				hRes = S_OK;
			}
		}
	}
} // CFolderDlg::OnSetMessageStatus

_Check_return_ HRESULT CFolderDlg::OnSubmitMessage(int iItem, _In_ SortListData* /*lpData*/)
{
	HRESULT			hRes = S_OK;
	LPMESSAGE		lpMessage = NULL;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric,CLASS,_T("OnSubmitMesssage"),_T("\n"));

	if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;

	EC_H(OpenItemProp(
		iItem,
		mfcmapiREQUEST_MODIFY,
		(LPMAPIPROP*)&lpMessage));

	if (lpMessage)
	{
		// Get subject line of message to copy.
		// This will be used as the new file name.
		EC_H(lpMessage->SubmitMessage(NULL));

		lpMessage->Release();
	}
	return hRes;
} // CFolderDlg::OnSubmitMessage

_Check_return_ HRESULT CFolderDlg::OnAbortSubmit(int iItem, _In_ SortListData* lpData)
{
	HRESULT			hRes = S_OK;
	LPMDB			lpMDB = NULL;
	LPSBinary		lpMessageEID = NULL;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric,CLASS,_T("OnSubmitMesssage"),_T("\n"));

	if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;
	if (!m_lpMapiObjects || !lpData) return MAPI_E_INVALID_PARAMETER;
	if (SORTLIST_CONTENTS != lpData->ulSortDataType) return MAPI_E_INVALID_PARAMETER;

	lpMDB = m_lpMapiObjects->GetMDB(); // do not release

	lpMessageEID = lpData->data.Contents.lpEntryID;

	if (lpMDB && lpMessageEID)
	{
		EC_H(lpMDB->AbortSubmit(
			lpMessageEID->cb,
			(LPENTRYID)lpMessageEID->lpb,
			NULL));
	}

	return hRes;
} // CFolderDlg::OnAbortSubmit

void CFolderDlg::HandleAddInMenuSingle(
									   _In_ LPADDINMENUPARAMS lpParams,
									   _In_ LPMAPIPROP lpMAPIProp,
									   _In_ LPMAPICONTAINER /*lpContainer*/)
{
	if (lpParams)
	{
		lpParams->lpFolder = (LPMAPIFOLDER) m_lpContainer; // m_lpContainer is an LPMAPIFOLDER
		lpParams->lpMessage = (LPMESSAGE) lpMAPIProp; // OpenItemProp returns LPMESSAGE
		// Add appropriate flag to context
		if (m_ulDisplayFlags & dfAssoc)
			lpParams->ulCurrentFlags |= MENU_FLAGS_FOLDER_ASSOC;
		if (m_ulDisplayFlags & dfDeleted)
			lpParams->ulCurrentFlags |= MENU_FLAGS_DELETED;
	}

	InvokeAddInMenu(lpParams);
} // CFolderDlg::HandleAddInMenuSingle