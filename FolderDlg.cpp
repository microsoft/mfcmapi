// FolderDlg.cpp : Displays the contents of a folder

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
#include "MAPIFormFunctions.h"
#include "InterpretProp.h"
#include "FileDialogEx.h"
#include "ExtraPropTags.h"
#include "PropertyTagEditor.h"
#include "MAPIProgress.h"
#include "MAPIMime.h"
#include "InterpretProp2.h"
#include "SortList/ContentsData.h"

static wstring CLASS = L"CFolderDlg";

CFolderDlg::CFolderDlg(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects,
	_In_ LPMAPIFOLDER lpMAPIFolder,
	ULONG ulDisplayFlags
) :
	CContentsTableDlg(
		pParentWnd,
		lpMapiObjects,
		IDS_FOLDER,
		mfcmapiDO_NOT_CALL_CREATE_DIALOG,
		nullptr,
		LPSPropTagArray(&sptMSGCols),
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
}

CFolderDlg::~CFolderDlg()
{
	TRACE_DESTRUCTOR(CLASS);
}

_Check_return_ bool CFolderDlg::HandleMenu(WORD wMenuSelect)
{
	DebugPrint(DBGMenu, L"CFolderDlg::HandleMenu wMenuSelect = 0x%X = %u\n", wMenuSelect, wMenuSelect);
	auto hRes = S_OK;
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
	case ID_HIERARCHY:
	case ID_CONTENTS:
	case ID_HIDDENCONTENTS:
		OnDisplayFolder(wMenuSelect); return true;
	}

	if (MultiSelectSimple(wMenuSelect)) return true;

	if (MultiSelectComplex(wMenuSelect)) return true;

	return CContentsTableDlg::HandleMenu(wMenuSelect);
}

typedef HRESULT(CFolderDlg::* LPSIMPLEMULTI)
(
	int iItem,
	SortListData* lpData
	);

_Check_return_ bool CFolderDlg::MultiSelectSimple(WORD wMenuSelect)
{
	LPSIMPLEMULTI lpFunc = nullptr;
	auto hRes = S_OK;

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
			auto iItem = m_lpContentsTableListCtrl->GetNextItem(
				-1,
				LVNI_SELECTED);
			while (-1 != iItem)
			{
				auto lpData = reinterpret_cast<SortListData*>(m_lpContentsTableListCtrl->GetItemData(iItem));
				WC_H((this->*lpFunc)(iItem, lpData));
				iItem = m_lpContentsTableListCtrl->GetNextItem(
					iItem,
					LVNI_SELECTED);
				if (S_OK != hRes && -1 != iItem)
				{
					if (bShouldCancel(this, hRes)) break;
					hRes = S_OK;
				}
			}

			return true;
		}
	}

	return false;
}

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
}

void CFolderDlg::OnDisplayItem()
{
	auto hRes = S_OK;
	LPMAPIPROP lpMAPIProp = nullptr;
	auto iItem = -1;
	CWaitCursor Wait;

	do
	{
		EC_H(m_lpContentsTableListCtrl->OpenNextSelectedItemProp(
			&iItem,
			mfcmapiREQUEST_MODIFY,
			&lpMAPIProp));

		if (lpMAPIProp)
		{
			EC_H(DisplayObject(
				lpMAPIProp,
				NULL,
				otDefault,
				this));
			lpMAPIProp->Release();
			lpMAPIProp = nullptr;
		}

		hRes = S_OK;
	} while (iItem != -1);
}

void CFolderDlg::OnInitMenu(_In_ CMenu* pMenu)
{
	if (pMenu && m_lpContentsTableListCtrl)
	{
		LPMAPISESSION lpMAPISession = nullptr;
		int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
		if (m_lpMapiObjects)
		{
			auto ulStatus = m_lpMapiObjects->GetBufferStatus();
			pMenu->EnableMenuItem(ID_PASTE, DIM(ulStatus & BUFFER_MESSAGES));
			lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		}

		pMenu->EnableMenuItem(ID_ADDTESTADDRESS, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_ATTACHMENTPROPERTIES, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_COPY, DIMMSOK(iNumSel));

		pMenu->EnableMenuItem(ID_DELETEATTACHMENTS, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_DELETESELECTEDITEM, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_MANUALRESOLVE, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_OPENMODAL, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_OPENNONMODAL, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_REMOVEONEOFF, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_RESENDMESSAGE, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_RTFSYNC, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_SAVEATTACHMENTS, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_SAVEMESSAGETOFILE, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_SETREADFLAG, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_SETMESSAGESTATUS, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_GETMESSAGESTATUS, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_GETMESSAGEOPTIONS, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_SUBMITMESSAGE, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_ABORTSUBMIT, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_RECIPIENTPROPERTIES, DIMMSOK(iNumSel));

		pMenu->EnableMenuItem(ID_GETPROPSUSINGLONGTERMEID, DIMMSNOK(iNumSel));
		pMenu->EnableMenuItem(ID_EXECUTEVERBONFORM, DIMMSNOK(iNumSel));

		pMenu->EnableMenuItem(ID_RESOLVEMESSAGECLASS, DIM(lpMAPISession));
		pMenu->EnableMenuItem(ID_SELECTFORM, DIM(lpMAPISession));

		pMenu->EnableMenuItem(ID_DISPLAYACLTABLE, DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_LOADFROMEML, DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_LOADFROMMSG, DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_LOADFROMTNEF, DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_NEW_APPOINTMENT, DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_NEW_CONTACT, DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_NEW_CUSTOMFORM, DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_NEW_IPMNOTE, DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_NEW_IPMPOST, DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_NEW_STICKYNOTE, DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_NEW_TASK, DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_NEW_MESSAGE, DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_SENDBULKMAIL, DIM(m_lpContainer));
		pMenu->EnableMenuItem(ID_SAVEFOLDERCONTENTSASTEXTFILES, DIM(m_lpContainer));

		pMenu->EnableMenuItem(ID_CONTENTS, DIM(m_lpContainer && !(m_ulDisplayFlags == dfNormal)));
		pMenu->EnableMenuItem(ID_HIDDENCONTENTS, DIM(m_lpContainer && !(m_ulDisplayFlags & dfAssoc)));
	}

	CContentsTableDlg::OnInitMenu(pMenu);
}

// Checks flags on add-in menu items to ensure they should be enabled
// Override to support context sensitive scenarios
void CFolderDlg::EnableAddInMenus(_In_ HMENU hMenu, ULONG ulMenu, _In_ LPMENUITEM lpAddInMenu, UINT uiEnable)
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

	if (hMenu) ::EnableMenuItem(hMenu, ulMenu, uiEnable);
}

void CFolderDlg::OnAddOneOffAddress()
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;
	auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	CEditor MyData(
		this,
		IDS_ADDONEOFF,
		NULL,
		5,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	MyData.InitPane(0, CreateSingleLinePaneID(IDS_DISPLAYNAME, IDS_DISPLAYNAMEVALUE, false));
	MyData.InitPane(1, CreateSingleLinePane(IDS_ADDRESSTYPE, wstring(L"EX"), false)); // STRING_OK
	MyData.InitPane(2, CreateSingleLinePaneID(IDS_ADDRESS, IDS_ADDRESSVALUE, false));
	MyData.InitPane(3, CreateSingleLinePane(IDS_RECIPTYPE, false));
	MyData.SetHex(3, MAPI_TO);
	MyData.InitPane(4, CreateSingleLinePane(IDS_RECIPCOUNT, false));
	MyData.SetDecimal(4, 1);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		wstring displayName = MyData.GetStringW(0);
		wstring addressType = MyData.GetStringW(1);
		wstring address = MyData.GetStringW(2);
		auto recipientType = MyData.GetHex(3);
		auto count = MyData.GetDecimal(4);
		LPMESSAGE lpMessage = nullptr;
		auto iItem = m_lpContentsTableListCtrl->GetNextItem(
			-1,
			LVNI_SELECTED);
		while (iItem != -1)
		{
			EC_H(OpenItemProp(
				iItem,
				mfcmapiREQUEST_MODIFY,
				reinterpret_cast<LPMAPIPROP*>(&lpMessage)));

			if (lpMessage)
			{
				if (count <= 1)
				{
					EC_H(AddOneOffAddress(
						lpMAPISession,
						lpMessage,
						displayName,
						addressType,
						address,
						recipientType));
				}
				else
				{
					for (ULONG i = 0; i < count; i++)
					{
						auto countedDisplayName = displayName + format(L"%d", i); // STRING_OK
						EC_H(AddOneOffAddress(
							lpMAPISession,
							lpMessage,
							countedDisplayName,
							addressType,
							address,
							recipientType));
					}
				}

				lpMessage->Release();
				lpMessage = nullptr;
			}

			iItem = m_lpContentsTableListCtrl->GetNextItem(
				iItem,
				LVNI_SELECTED);
			if (S_OK != hRes && -1 != iItem)
			{
				if (bShouldCancel(this, hRes)) break;
				hRes = S_OK;
			}
		}
	}
}

_Check_return_ HRESULT CFolderDlg::OnAttachmentProperties(int iItem, _In_ SortListData* /*lpData*/)
{
	auto hRes = S_OK;
	LPMESSAGE lpMessage = nullptr;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;

	EC_H(OpenItemProp(
		iItem,
		mfcmapiREQUEST_MODIFY,
		reinterpret_cast<LPMAPIPROP*>(&lpMessage)));

	if (lpMessage)
	{
		EC_H(OpenAttachmentsFromMessage(lpMessage));

		lpMessage->Release();
	}

	return hRes;
}

void CFolderDlg::HandleCopy()
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric, CLASS, L"HandleCopy", L"\n");
	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

	LPENTRYLIST lpEIDs = nullptr;

	EC_H(m_lpContentsTableListCtrl->GetSelectedItemEIDs(&lpEIDs));

	// m_lpMapiObjects takes over ownership of lpEIDs - don't free now
	m_lpMapiObjects->SetMessagesToCopy(lpEIDs, static_cast<LPMAPIFOLDER>(m_lpContainer));
}

_Check_return_ bool CFolderDlg::HandlePaste()
{
	if (CBaseDialog::HandlePaste()) return true;

	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric, CLASS, L"HandlePaste", L"\n");
	if (!m_lpMapiObjects || !m_lpContainer) return false;

	CEditor MyData(
		this,
		IDS_COPYMESSAGE,
		IDS_COPYMESSAGEPROMPT,
		2,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	MyData.InitPane(0, CreateCheckPane(IDS_MESSAGEMOVE, false, false));
	UINT uidDropDown[] = {
	IDS_DDCOPYMESSAGES,
	IDS_DDCOPYTO
	};
	MyData.InitPane(1, CreateDropDownPane(IDS_COPYINTERFACE, _countof(uidDropDown), uidDropDown, true));

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		auto lpEIDs = m_lpMapiObjects->GetMessagesToCopy();
		auto lpMAPISourceFolder = m_lpMapiObjects->GetSourceParentFolder();
		auto ulMoveMessage = MyData.GetCheck(0) ? MESSAGE_MOVE : 0;

		if (lpEIDs && lpMAPISourceFolder)
		{
			if (0 == MyData.GetDropDown(1))
			{ // CopyMessages
				LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IMAPIFolder::CopyMessages", m_hWnd); // STRING_OK

				if (lpProgress)
					ulMoveMessage |= MESSAGE_DIALOG;

				EC_MAPI(lpMAPISourceFolder->CopyMessages(
					lpEIDs,
					&IID_IMAPIFolder,
					m_lpContainer,
					lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL,
					lpProgress,
					ulMoveMessage));

				if (lpProgress)
					lpProgress->Release();
			}
			else
			{ // CopyTo
			// Specify properties to exclude in the copy operation. These are
			// the properties that Exchange excludes to save bits and time.
			// Should not be necessary to exclude these, but speeds the process
			// when a lot of messages are being copied.
				static const SizedSPropTagArray(7, excludeTags) =
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

				auto lpTagsToExclude = GetExcludedTags(LPSPropTagArray(&excludeTags), m_lpContainer, m_bIsAB);

				if (lpTagsToExclude)
				{
					for (ULONG i = 0; i < lpEIDs->cValues; i++)
					{
						LPMESSAGE lpMessage = nullptr;

						EC_H(CallOpenEntry(
							NULL,
							NULL,
							lpMAPISourceFolder,
							NULL,
							lpEIDs->lpbin[i].cb,
							reinterpret_cast<LPENTRYID>(lpEIDs->lpbin[i].lpb),
							NULL,
							MyData.GetCheck(0) ? MAPI_MODIFY : 0, // only need write access for moves
							NULL,
							reinterpret_cast<LPUNKNOWN*>(&lpMessage)));
						if (lpMessage)
						{
							LPMESSAGE lpNewMessage = nullptr;
							EC_MAPI((static_cast<LPMAPIFOLDER>(m_lpContainer))->CreateMessage(NULL, (m_ulDisplayFlags & dfAssoc) ? MAPI_ASSOCIATED : NULL, &lpNewMessage));
							if (lpNewMessage)
							{
								LPSPropProblemArray lpProblems = nullptr;

								// copy message properties to IMessage object opened on top of IStorage.
								LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IMAPIProp::CopyTo", m_hWnd); // STRING_OK

								if (lpProgress)
									ulMoveMessage |= MAPI_DIALOG;

								EC_MAPI(lpMessage->CopyTo(
									0,
									NULL, // ODO: interfaces?
									lpTagsToExclude,
									lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL, // UI param
									lpProgress, // progress
									const_cast<LPIID>(&IID_IMessage),
									lpNewMessage,
									ulMoveMessage,
									&lpProblems));

								if (lpProgress)
									lpProgress->Release();

								EC_PROBLEMARRAY(lpProblems);
								MAPIFreeBuffer(lpProblems);

								// save changes to IMessage object.
								EC_MAPI(lpNewMessage->SaveChanges(KEEP_OPEN_READWRITE));
								lpNewMessage->Release();

								if (MyData.GetCheck(0)) // if we moved, save changes on original message
								{
									EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
								}
							}

							lpMessage->Release();
							lpMessage = nullptr;
						}
					}

					MAPIFreeBuffer(lpTagsToExclude);
				}
			}
		}

		if (lpMAPISourceFolder) lpMAPISourceFolder->Release();
	}

	return true;
}

void CFolderDlg::OnDeleteAttachments()
{
	auto hRes = S_OK;

	CEditor MyData(
		this,
		IDS_DELETEATTACHMENTS,
		IDS_DELETEATTACHMENTSPROMPT,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	MyData.InitPane(0, CreateSingleLinePane(IDS_FILENAME, false));

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		wstring szAttName = MyData.GetStringW(0);

		if (!szAttName.empty())
		{
			auto iItem = m_lpContentsTableListCtrl->GetNextItem(
				-1,
				LVNI_SELECTED);

			while (-1 != iItem)
			{
				LPMESSAGE lpMessage = nullptr;

				EC_H(OpenItemProp(
					iItem,
					mfcmapiREQUEST_MODIFY,
					reinterpret_cast<LPMAPIPROP*>(&lpMessage)));

				if (lpMessage)
				{
					EC_H(DeleteAttachments(lpMessage, szAttName, m_hWnd));

					lpMessage->Release();
				}

				iItem = m_lpContentsTableListCtrl->GetNextItem(
					iItem,
					LVNI_SELECTED);
			}
		}
	}
}

void CFolderDlg::OnDeleteSelectedItem()
{
	auto hRes = S_OK;
	LPMDB lpMDB = nullptr;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl || !m_lpContainer) return;

	auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	EC_H(OpenDefaultMessageStore(lpMAPISession, &lpMDB));
	if (!lpMDB) return;

	auto bMove = false;
	auto ulFlag = MESSAGE_DIALOG;

	if (m_ulDisplayFlags & dfDeleted)
	{
		ulFlag |= DELETE_HARD_DELETE;
	}
	else
	{
		auto bShift = !(GetKeyState(VK_SHIFT) < 0);

		CEditor MyData(
			this,
			IDS_DELETEITEM,
			IDS_DELETEITEMPROMPT,
			1,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		UINT uidDropDown[] = {
		IDS_DDDELETETODELETED,
		IDS_DDDELETETORETENTION,
		IDS_DDDELETEHARDDELETE
		};

		if (bShift)
			MyData.InitPane(0, CreateDropDownPane(IDS_DELSTYLE, _countof(uidDropDown), uidDropDown, true));
		else
			MyData.InitPane(0, CreateDropDownPane(IDS_DELSTYLE, _countof(uidDropDown) - 1, &uidDropDown[1], true));

		WC_H(MyData.DisplayDialog());

		if (bShift)
		{
			switch (MyData.GetDropDown(0))
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
		LPENTRYLIST lpEIDs = nullptr;

		EC_H(m_lpContentsTableListCtrl->GetSelectedItemEIDs(&lpEIDs));

		if (bMove)
		{
			EC_H(DeleteToDeletedItems(
				lpMDB,
				static_cast<LPMAPIFOLDER>(m_lpContainer),
				lpEIDs,
				m_hWnd));
		}
		else
		{
			LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IMAPIFolder::DeleteMessages", m_hWnd); // STRING_OK

			if (lpProgress)
				ulFlag |= MESSAGE_DIALOG;

			EC_MAPI((static_cast<LPMAPIFOLDER>(m_lpContainer))->DeleteMessages(
				lpEIDs, // list of messages to delete
				lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL,
				lpProgress,
				ulFlag));

			if (lpProgress)
				lpProgress->Release();
		}

		MAPIFreeBuffer(lpEIDs);
	}

	lpMDB->Release();
}

void CFolderDlg::OnGetPropsUsingLongTermEID()
{
	auto hRes = S_OK;
	auto iItem = -1;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl) return;

	auto lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);

	if (lpListData && lpListData->Contents())
	{
		auto lpMessageEID = lpListData->Contents()->m_lpLongtermID;

		if (lpMessageEID)
		{
			LPMAPIPROP lpMAPIProp = nullptr;

			auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
			if (lpMDB)
			{
				WC_H(CallOpenEntry(
					lpMDB,
					NULL,
					NULL,
					NULL,
					lpMessageEID->cb,
					reinterpret_cast<LPENTRYID>(lpMessageEID->lpb),
					NULL,
					MAPI_BEST_ACCESS,
					NULL,
					reinterpret_cast<LPUNKNOWN*>(&lpMAPIProp)));
			}

			OnUpdateSingleMAPIPropListCtrl(lpMAPIProp, nullptr);

			if (lpMAPIProp) lpMAPIProp->Release();
		}
	}
}

// Use CFileDialogExW to locate a .MSG file to load,
// Pass the file name and a message to load in to LoadFromMsg to do the work.
void CFolderDlg::OnLoadFromMSG()
{
	if (!m_lpContainer) return;

	auto hRes = S_OK;
	LPMESSAGE lpNewMessage = nullptr;
	INT_PTR iDlgRet = IDOK;

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
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		UINT uidDropDown[] = {
		IDS_DDLOADTOFOLDER,
		IDS_DDDISPLAYPROPSONLY
		};
		MyData.InitPane(0, CreateDropDownPane(IDS_LOADSTYLE, _countof(uidDropDown), uidDropDown, true));

		WC_H(MyData.DisplayDialog());
		if (S_OK == hRes)
		{
			LPWSTR lpszPath = nullptr;
			while (NULL != (lpszPath = dlgFilePicker.GetNextFileName()))
			{
				hRes = S_OK;
				switch (MyData.GetDropDown(0))
				{
				case 0:
					EC_MAPI((static_cast<LPMAPIFOLDER>(m_lpContainer))->CreateMessage(
						NULL,
						(m_ulDisplayFlags & dfAssoc) ? MAPI_ASSOCIATED : NULL,
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
							EC_H(m_lpPropDisplay->SetDataSource(lpNewMessage, NULL, false));
						}
					}

					break;
				}

				if (lpNewMessage)
				{
					lpNewMessage->Release();
					lpNewMessage = nullptr;
				}
			}
		}
	}
}

void CFolderDlg::OnResolveMessageClass()
{
	if (!m_lpMapiObjects) return;

	LPMAPIFORMINFO lpMAPIFormInfo = nullptr;
	ResolveMessageClass(m_lpMapiObjects, static_cast<LPMAPIFOLDER>(m_lpContainer), &lpMAPIFormInfo);
	if (lpMAPIFormInfo)
	{
		OnUpdateSingleMAPIPropListCtrl(lpMAPIFormInfo, nullptr);
		lpMAPIFormInfo->Release();
	}
}

void CFolderDlg::OnSelectForm()
{
	LPMAPIFORMINFO lpMAPIFormInfo = nullptr;

	if (!m_lpMapiObjects) return;

	SelectForm(m_hWnd, m_lpMapiObjects, static_cast<LPMAPIFOLDER>(m_lpContainer), &lpMAPIFormInfo);
	if (lpMAPIFormInfo)
	{
		OnUpdateSingleMAPIPropListCtrl(lpMAPIFormInfo, nullptr);
		// TODO: Put some code in here which works with the returned Form Info pointer
		lpMAPIFormInfo->Release();
	}
}

void CFolderDlg::OnManualResolve()
{
	auto hRes = S_OK;
	LPMESSAGE lpMessage = nullptr;
	auto iItem = -1;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

	auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
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
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.InitPane(0, CreateSingleLinePane(IDS_DISPLAYNAME, false));

		WC_H(MyData.DisplayDialog());
		if (S_OK == hRes)
		{
			do
			{
				hRes = S_OK;
				EC_H(m_lpContentsTableListCtrl->OpenNextSelectedItemProp(
					&iItem,
					mfcmapiREQUEST_MODIFY,
					reinterpret_cast<LPMAPIPROP*>(&lpMessage)));

				if (lpMessage)
				{
					wstring name = MyData.GetStringW(0);
					EC_H(ManualResolve(
						lpMAPISession,
						lpMessage,
						name,
						CHANGE_PROP_TYPE(MyPropertyTag.GetPropertyTag(), PT_UNICODE)));

					lpMessage->Release();
					lpMessage = nullptr;
				}
			} while (iItem != -1);
		}
	}
}

void CFolderDlg::NewSpecialItem(WORD wMenuSelect) const
{
	auto hRes = S_OK;
	LPMAPIFOLDER lpFolder = nullptr; // DO NOT RELEASE

	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

	auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return;

	auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release

	if (lpMAPISession)
	{
		ULONG ulFolder = NULL;
		wstring szClass;

		switch (wMenuSelect)
		{
		case ID_NEW_APPOINTMENT:
			ulFolder = DEFAULT_CALENDAR;
			szClass = L"IPM.APPOINTMENT"; // STRING_OK
			break;
		case ID_NEW_CONTACT:
			ulFolder = DEFAULT_CONTACTS;
			szClass = L"IPM.CONTACT"; // STRING_OK
			break;
		case ID_NEW_IPMNOTE:
			szClass = L"IPM.NOTE"; // STRING_OK
			break;
		case ID_NEW_IPMPOST:
			szClass = L"IPM.POST"; // STRING_OK
			break;
		case ID_NEW_TASK:
			ulFolder = DEFAULT_TASKS;
			szClass = L"IPM.TASK"; // STRING_OK
			break;
		case ID_NEW_STICKYNOTE:
			ulFolder = DEFAULT_NOTES;
			szClass = L"IPM.STICKYNOTE"; // STRING_OK
			break;

		}

		LPMAPIFOLDER lpSpecialFolder = nullptr;
		if (ulFolder)
		{
			EC_H(OpenDefaultFolder(ulFolder, lpMDB, &lpSpecialFolder));
			lpFolder = lpSpecialFolder;
		}
		else
		{
			lpFolder = static_cast<LPMAPIFOLDER>(m_lpContainer);
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
}

void CFolderDlg::OnNewMessage()
{
	auto hRes = S_OK;
	LPMESSAGE lpMessage = nullptr;

	EC_MAPI((static_cast<LPMAPIFOLDER>(m_lpContainer))->CreateMessage(
		NULL,
		m_ulDisplayFlags & dfAssoc ? MAPI_ASSOCIATED : 0,
		&lpMessage));

	if (lpMessage)
	{
		EC_MAPI(lpMessage->SaveChanges(NULL));
		lpMessage->Release();
	}
}

void CFolderDlg::OnNewCustomForm()
{
	auto hRes = S_OK;

	if (!m_lpMapiObjects || !m_lpContainer || !m_lpContentsTableListCtrl) return;

	auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return;

	auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release

	if (lpMAPISession)
	{
		CEditor MyPrompt1(
			this,
			IDS_NEWCUSTOMFORM,
			IDS_NEWCUSTOMFORMPROMPT1,
			1,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		UINT uidDropDown[] = {
		IDS_DDENTERFORMCLASS,
		IDS_DDFOLDERFORMLIBRARY,
		IDS_DDORGFORMLIBRARY
		};
		MyPrompt1.InitPane(0, CreateDropDownPane(IDS_LOCATIONOFFORM, _countof(uidDropDown), uidDropDown, true));

		WC_H(MyPrompt1.DisplayDialog());

		if (S_OK == hRes)
		{
			wstring szClass;
			LPSPropValue lpProp = nullptr;

			CEditor MyClass(
				this,
				IDS_NEWCUSTOMFORM,
				IDS_NEWCUSTOMFORMPROMPT2,
				1,
				CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyClass.InitPane(0, CreateSingleLinePane(IDS_FORMTYPE, wstring(L"IPM.Note"), false)); // STRING_OK

			switch (MyPrompt1.GetDropDown(0))
			{
			case 0:
				WC_H(MyClass.DisplayDialog());

				if (S_OK == hRes)
				{
					szClass = MyClass.GetStringW(0);
				}

				break;
			case 1:
			case 2:
			{
				LPMAPIFORMMGR lpMAPIFormMgr = nullptr;
				LPMAPIFORMINFO lpMAPIFormInfo = nullptr;
				EC_MAPI(MAPIOpenFormMgr(lpMAPISession, &lpMAPIFormMgr));

				if (lpMAPIFormMgr)
				{
					LPMAPIFOLDER lpFormSource = nullptr;
					if (1 == MyPrompt1.GetDropDown(0)) lpFormSource = static_cast<LPMAPIFOLDER>(m_lpContainer);
					auto szTitle = loadstring(IDS_SELECTFORMCREATE);

					// Apparently, SelectForm doesn't support unicode
					EC_H_CANCEL(lpMAPIFormMgr->SelectForm(
						reinterpret_cast<ULONG_PTR>(m_hWnd),
						0,
						reinterpret_cast<LPCTSTR>(wstringTostring(szTitle).c_str()),
						lpFormSource,
						&lpMAPIFormInfo));

					if (lpMAPIFormInfo)
					{
						EC_MAPI(HrGetOneProp(
							lpMAPIFormInfo,
							PR_MESSAGE_CLASS_W,
							&lpProp));
						if (CheckStringProp(lpProp, PT_UNICODE))
						{
							szClass = lpProp->Value.lpszW;
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
				static_cast<LPMAPIFOLDER>(m_lpContainer)));

			MAPIFreeBuffer(lpProp);
		}
	}
}

_Check_return_ HRESULT CFolderDlg::OnOpenModal(int iItem, _In_ SortListData* /*lpData*/)
{
	auto hRes = S_OK;
	LPMESSAGE lpMessage = nullptr;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;
	if (!m_lpMapiObjects || !m_lpContainer) return MAPI_E_INVALID_PARAMETER;

	auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return MAPI_E_CALL_FAILED;

	auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (lpMAPISession)
	{
		// Before we open the message, make sure the MAPI Form Manager is implemented
		LPMAPIFORMMGR lpFormMgr = nullptr;
		WC_MAPI(MAPIOpenFormMgr(lpMAPISession, &lpFormMgr));
		hRes = S_OK; // Ditch the error if we got one

		if (lpFormMgr)
		{
			EC_H(OpenItemProp(
				iItem,
				mfcmapiREQUEST_MODIFY,
				reinterpret_cast<LPMAPIPROP*>(&lpMessage)));
			if (lpMessage)
			{
				EC_H(OpenMessageModal(
					static_cast<LPMAPIFOLDER>(m_lpContainer),
					lpMAPISession,
					lpMDB,
					lpMessage));

				lpMessage->Release();
			}

			lpFormMgr->Release();
		}
	}

	return hRes;
}

_Check_return_ HRESULT CFolderDlg::OnOpenNonModal(int iItem, _In_ SortListData* /*lpData*/)
{
	auto hRes = S_OK;
	LPMESSAGE lpMessage = nullptr;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;
	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl || !m_lpContainer) return MAPI_E_INVALID_PARAMETER;

	auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return MAPI_E_CALL_FAILED;

	auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (lpMAPISession)
	{
		// Before we open the message, make sure the MAPI Form Manager is implemented
		LPMAPIFORMMGR lpFormMgr = nullptr;
		WC_MAPI(MAPIOpenFormMgr(lpMAPISession, &lpFormMgr));
		hRes = S_OK; // Ditch the error if we got one

		if (lpFormMgr)
		{
			EC_H(OpenItemProp(
				iItem,
				mfcmapiREQUEST_MODIFY,
				reinterpret_cast<LPMAPIPROP*>(&lpMessage)));

			if (lpMessage)
			{
				WC_H(OpenMessageNonModal(
					m_hWnd,
					lpMDB,
					lpMAPISession,
					static_cast<LPMAPIFOLDER>(m_lpContainer),
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
}

void CFolderDlg::OnExecuteVerbOnForm()
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

	auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return;

	auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (lpMAPISession)
	{
		CEditor MyData(
			this,
			IDS_EXECUTEVERB,
			IDS_EXECUTEVERBPROMPT,
			1,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.SetPromptPostFix(AllFlagsToString(PROP_ID(PR_LAST_VERB_EXECUTED), false));

		MyData.InitPane(0, CreateSingleLinePane(IDS_VERB, false));
		MyData.SetDecimal(0, EXCHIVERB_OPEN);

		WC_H(MyData.DisplayDialog());
		if (S_OK == hRes)
		{
			LPMESSAGE lpMessage = nullptr;
			auto iItem = m_lpContentsTableListCtrl->GetNextItem(
				-1,
				LVNI_SELECTED);
			if (iItem != -1)
			{
				EC_H(OpenItemProp(
					iItem,
					mfcmapiREQUEST_MODIFY,
					reinterpret_cast<LPMAPIPROP*>(&lpMessage)));

				if (lpMessage)
				{
					EC_H(OpenMessageNonModal(
						m_hWnd,
						lpMDB,
						lpMAPISession,
						static_cast<LPMAPIFOLDER>(m_lpContainer),
						m_lpContentsTableListCtrl,
						iItem,
						lpMessage,
						MyData.GetDecimal(0),
						NULL));

					lpMessage->Release();
					lpMessage = nullptr;
				}
			}
		}
	}
}

_Check_return_ HRESULT CFolderDlg::OnResendSelectedItem(int /*iItem*/, _In_ SortListData* lpData)
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!lpData || !lpData->Contents() || !m_lpContainer) return MAPI_E_INVALID_PARAMETER;

	if (lpData->Contents()->m_lpEntryID)
	{
		EC_H(ResendSingleMessage(
			static_cast<LPMAPIFOLDER>(m_lpContainer),
			lpData->Contents()->m_lpEntryID,
			m_hWnd));
	}

	return hRes;
}

_Check_return_ HRESULT CFolderDlg::OnRecipientProperties(int iItem, _In_ SortListData* /*lpData*/)
{
	auto hRes = S_OK;
	LPMESSAGE lpMessage = nullptr;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;

	EC_H(OpenItemProp(
		iItem,
		mfcmapiREQUEST_MODIFY,
		reinterpret_cast<LPMAPIPROP*>(&lpMessage)));

	if (lpMessage)
	{
		EC_H(OpenRecipientsFromMessage(lpMessage));

		lpMessage->Release();
	}

	return hRes;
}

void CFolderDlg::OnRemoveOneOff()
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl) return;

	CEditor MyData(
		this,
		IDS_REMOVEONEOFF,
		IDS_REMOVEONEOFFPROMPT,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CreateCheckPane(IDS_REMOVEPROPDEFSTREAM, true, false));

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		LPMESSAGE lpMessage = nullptr;

		auto iItem = m_lpContentsTableListCtrl->GetNextItem(
			-1,
			LVNI_SELECTED);
		while (iItem != -1)
		{
			EC_H(OpenItemProp(
				iItem,
				mfcmapiREQUEST_MODIFY,
				reinterpret_cast<LPMAPIPROP*>(&lpMessage)));

			if (lpMessage)
			{
				DebugPrint(DBGGeneric, L"Calling RemoveOneOff on %p, %wsremoving property definition stream\n", lpMessage, MyData.GetCheck(0) ? L"" : L"not ");
				EC_H(RemoveOneOff(
					lpMessage,
					MyData.GetCheck(0)));

				lpMessage->Release();
				lpMessage = nullptr;
			}

			iItem = m_lpContentsTableListCtrl->GetNextItem(
				iItem,
				LVNI_SELECTED);
			if (S_OK != hRes && -1 != iItem)
			{
				if (bShouldCancel(this, hRes)) break;
				hRes = S_OK;
			}
		}
	}
}

#define RTF_SYNC_HTML_CHANGED ((ULONG) 0x00000004)

void CFolderDlg::OnRTFSync()
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl) return;

	CEditor MyData(
		this,
		IDS_CALLRTFSYNC,
		IDS_CALLRTFSYNCPROMPT,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	MyData.InitPane(0, CreateSingleLinePane(IDS_FLAGS, false));
	MyData.SetHex(0, RTF_SYNC_RTF_CHANGED);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		LPMESSAGE lpMessage = nullptr;
		BOOL bMessageUpdated = false;

		auto iItem = m_lpContentsTableListCtrl->GetNextItem(
			-1,
			LVNI_SELECTED);
		while (iItem != -1)
		{
			EC_H(OpenItemProp(
				iItem,
				mfcmapiREQUEST_MODIFY,
				reinterpret_cast<LPMAPIPROP*>(&lpMessage)));

			if (lpMessage)
			{
				DebugPrint(DBGGeneric, L"Calling RTFSync on %p with flags 0x%X\n", lpMessage, MyData.GetHex(0));
				EC_MAPI(RTFSync(
					lpMessage,
					MyData.GetHex(0),
					&bMessageUpdated));

				DebugPrint(DBGGeneric, L"RTFSync returned %d\n", bMessageUpdated);

				EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));

				lpMessage->Release();
				lpMessage = nullptr;
			}

			iItem = m_lpContentsTableListCtrl->GetNextItem(
				iItem,
				LVNI_SELECTED);
			if (S_OK != hRes && -1 != iItem)
			{
				if (bShouldCancel(this, hRes)) break;
				hRes = S_OK;
			}
		}
	}
}

_Check_return_ HRESULT CFolderDlg::OnSaveAttachments(int iItem, _In_ SortListData* /*lpData*/)
{
	auto hRes = S_OK;
	LPMESSAGE lpMessage = nullptr;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;

	EC_H(OpenItemProp(
		iItem,
		mfcmapiREQUEST_MODIFY,
		reinterpret_cast<LPMAPIPROP*>(&lpMessage)));

	if (lpMessage)
	{
		EC_H(WriteAttachmentsToFile(
			lpMessage, m_hWnd));

		lpMessage->Release();
	}

	return hRes;
}

void CFolderDlg::OnSaveFolderContentsAsTextFiles()
{
	if (!m_lpMapiObjects || !m_lpContainer) return;

	auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return;

	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	SaveFolderContentsToTXT(
		lpMDB,
		static_cast<LPMAPIFOLDER>(m_lpContainer),
		(m_ulDisplayFlags & dfAssoc) ? false : true,
		(m_ulDisplayFlags & dfAssoc) ? true : false,
		false,
		m_hWnd);
}

void CFolderDlg::OnSaveMessageToFile()
{
	auto hRes = S_OK;

	DebugPrintEx(DBGGeneric, CLASS, L"OnSaveMessageToFile", L"\n");

	CEditor MyData(
		this,
		IDS_SAVEMESSAGETOFILE,
		IDS_SAVEMESSAGETOFILEPROMPT,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	UINT uidDropDown[] = {
	IDS_DDTEXTFILE,
	IDS_DDMSGFILEANSI,
	IDS_DDMSGFILEUNICODE,
	IDS_DDEMLFILE,
	IDS_DDEMLFILEUSINGICONVERTERSESSION,
	IDS_DDTNEFFILE
	};
	MyData.InitPane(0, CreateDropDownPane(IDS_FORMATTOSAVEMESSAGE, _countof(uidDropDown), uidDropDown, true));
	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		LPCWSTR szExt = nullptr;
		LPCWSTR szDotExt = nullptr;
		ULONG ulDotExtLen = NULL;
		CStringW szFilter;
		LPADRBOOK lpAddrBook = nullptr;
		switch (MyData.GetDropDown(0))
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

		auto iItem = m_lpContentsTableListCtrl->GetNextItem(
			-1,
			LVNI_SELECTED);
		while (-1 != iItem)
		{
			LPMESSAGE lpMessage = nullptr;
			WCHAR szFileName[MAX_PATH] = { 0 };
			INT_PTR iDlgRet = IDOK;

			EC_H(OpenItemProp(
				iItem,
				mfcmapiREQUEST_MODIFY,
				reinterpret_cast<LPMAPIPROP*>(&lpMessage)));

			if (lpMessage)
			{
				WC_H(BuildFileName(szFileName, _countof(szFileName), szDotExt, ulDotExtLen, lpMessage));

				DebugPrint(DBGGeneric, L"BuildFileNameAndPath built file name \"%ws\"\n", szFileName);

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
					switch (MyData.GetDropDown(0))
					{
					case 0:
						// Idea is to capture anything that may be important about this message to disk so it can be analyzed.
					{
						CDumpStore MyDumpStore;
						MyDumpStore.InitMessagePath(dlgFilePicker.GetFileName());
						// Just assume this message might have attachments
						MyDumpStore.ProcessMessage(lpMessage, true, nullptr);
					}

					break;
					case 1:
						EC_H(SaveToMSG(lpMessage, dlgFilePicker.GetFileName(), false, m_hWnd, true));
						break;
					case 2:
						EC_H(SaveToMSG(lpMessage, dlgFilePicker.GetFileName(), true, m_hWnd, true));
						break;
					case 3:
						EC_H(SaveToEML(lpMessage, dlgFilePicker.GetFileName()));
						break;
					case 4:
					{
						ULONG ulConvertFlags = CCSF_SMTP;
						auto et = IET_UNKNOWN;
						auto mst = USE_DEFAULT_SAVETYPE;
						ULONG ulWrapLines = USE_DEFAULT_WRAPPING;
						auto bDoAdrBook = false;

						EC_H(GetConversionToEMLOptions(this, &ulConvertFlags, &et, &mst, &ulWrapLines, &bDoAdrBook));
						if (S_OK == hRes)
						{
							LPADRBOOK lpAdrBook = nullptr;
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
						EC_H(SaveToTNEF(lpMessage, lpAddrBook, dlgFilePicker.GetFileName()));
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
				CHECKHRESMSG(hRes, IDS_OPENMSGFAILED);
			}

			iItem = m_lpContentsTableListCtrl->GetNextItem(
				iItem,
				LVNI_SELECTED);
			if ((IDOK != iDlgRet || S_OK != hRes) && -1 != iItem)
			{
				if (bShouldCancel(this, hRes)) break;
				hRes = S_OK;
			}
		}
	}
}

// Use CFileDialogExW to locate a .DAT or .TNEF file to load,
// Pass the file name and a message to load in to LoadFromTNEF to do the work.
void CFolderDlg::OnLoadFromTNEF()
{
	auto hRes = S_OK;
	LPMESSAGE lpNewMessage = nullptr;
	INT_PTR iDlgRet = IDOK;

	if (!m_lpContainer) return;

	auto lpAddrBook = m_lpMapiObjects->GetAddrBook(true); // do not release
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
			LPWSTR lpszPath = nullptr;
			while (NULL != (lpszPath = dlgFilePicker.GetNextFileName()))
			{
				hRes = S_OK;
				EC_MAPI((static_cast<LPMAPIFOLDER>(m_lpContainer))->CreateMessage(
					NULL,
					(m_ulDisplayFlags & dfAssoc) ? MAPI_ASSOCIATED : 0,
					&lpNewMessage));

				if (lpNewMessage)
				{
					EC_H(LoadFromTNEF(
						lpszPath,
						lpAddrBook,
						lpNewMessage));

					lpNewMessage->Release();
					lpNewMessage = nullptr;
				}
			}
		}
	}

	if (lpNewMessage) lpNewMessage->Release();
}

// Use CFileDialogExW to locate a .EML file to load,
// Pass the file name and a message to load in to ImportEMLToIMessage to do the work.
void CFolderDlg::OnLoadFromEML()
{
	if (!m_lpContainer) return;

	auto hRes = S_OK;
	LPMESSAGE lpNewMessage = nullptr;
	INT_PTR iDlgRet = IDOK;

	ULONG ulConvertFlags = CCSF_SMTP;
	auto bDoAdrBook = false;
	auto bDoApply = false;
	HCHARSET hCharSet = nullptr;
	auto cSetApplyType = CSET_APPLY_UNTAGGED;
	WC_H(GetConversionFromEMLOptions(this, &ulConvertFlags, &bDoAdrBook, &bDoApply, &hCharSet, &cSetApplyType, NULL));
	if (S_OK == hRes)
	{
		LPADRBOOK lpAdrBook = nullptr;
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
			LPWSTR lpszPath = nullptr;
			while (NULL != (lpszPath = dlgFilePicker.GetNextFileName()))
			{
				hRes = S_OK;
				EC_MAPI((static_cast<LPMAPIFOLDER>(m_lpContainer))->CreateMessage(
					NULL,
					(m_ulDisplayFlags & dfAssoc) ? MAPI_ASSOCIATED : 0,
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
					lpNewMessage = nullptr;
				}
			}
		}
	}
}

void CFolderDlg::OnSendBulkMail()
{
	auto hRes = S_OK;

	CEditor MyData(
		this,
		IDS_SENDBULKMAIL,
		IDS_SENDBULKMAILPROMPT,
		5,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CreateSingleLinePane(IDS_NUMMESSAGES, false));
	MyData.InitPane(1, CreateSingleLinePane(IDS_RECIPNAME, false));
	MyData.InitPane(2, CreateSingleLinePane(IDS_SUBJECT, false));
	MyData.InitPane(3, CreateSingleLinePane(IDS_CLASS, wstring(L"IPM.Note"), false)); // STRING_OK
	MyData.InitPane(4, CreateMultiLinePane(IDS_BODY, false));

	if (!m_lpContainer) return;

	auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		auto ulNumMessages = MyData.GetDecimal(0);
		wstring szSubject = MyData.GetStringW(2);

		for (ULONG i = 0; i < ulNumMessages; i++)
		{
			hRes = S_OK;
			auto szTestSubject = formatmessage(IDS_TESTSUBJECT, szSubject.c_str(), i);

			EC_H(SendTestMessage(
				lpMAPISession,
				static_cast<LPMAPIFOLDER>(m_lpContainer),
				MyData.GetStringW(1),
				MyData.GetStringW(4),
				szTestSubject,
				MyData.GetStringW(3)));
			if (FAILED(hRes))
			{
				CHECKHRESMSG(hRes, IDS_ERRORSENDINGMSGS);
				break;
			}
		}
	}
}

void CFolderDlg::OnSetReadFlag()
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl) return;

	DebugPrintEx(DBGGeneric, CLASS, L"OnSetReadFlag", L"\n");

	CEditor MyFlags(
		this,
		IDS_SETREADFLAG,
		IDS_SETREADFLAGPROMPT,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	MyFlags.InitPane(0, CreateSingleLinePane(IDS_FLAGSINHEX, false));
	MyFlags.SetHex(0, CLEAR_READ_FLAG);

	WC_H(MyFlags.DisplayDialog());
	if (S_OK == hRes)
	{
		int iNumSelected = m_lpContentsTableListCtrl->GetSelectedCount();

		if (1 == iNumSelected)
		{
			LPMESSAGE lpMessage = nullptr;

			EC_H(m_lpContentsTableListCtrl->OpenNextSelectedItemProp(
				NULL,
				mfcmapiREQUEST_MODIFY,
				reinterpret_cast<LPMAPIPROP*>(&lpMessage)));

			if (lpMessage)
			{
				EC_MAPI(lpMessage->SetReadFlag(MyFlags.GetHex(0)));
				lpMessage->Release();
				lpMessage = nullptr;
			}
		}
		else if (iNumSelected > 1)
		{
			LPENTRYLIST lpEIDs = nullptr;

			EC_H(m_lpContentsTableListCtrl->GetSelectedItemEIDs(&lpEIDs));

			LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IMAPIFolder::SetReadFlags", m_hWnd); // STRING_OK

			auto ulFlags = MyFlags.GetHex(0);

			if (lpProgress)
				ulFlags |= MESSAGE_DIALOG;

			EC_MAPI((static_cast<LPMAPIFOLDER>(m_lpContainer))->SetReadFlags(
				lpEIDs,
				lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL,
				lpProgress,
				ulFlags));

			if (lpProgress)
				lpProgress->Release();

			MAPIFreeBuffer(lpEIDs);
		}
	}
}

void CFolderDlg::OnGetMessageOptions()
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl) return;

	auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	DebugPrintEx(DBGGeneric, CLASS, L"OnGetMessageOptions", L"\n");

	CEditor MyAddress(
		this,
		IDS_MESSAGEOPTIONS,
		IDS_ADDRESSTYPEPROMPT,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyAddress.InitPane(0, CreateSingleLinePane(IDS_ADDRESSTYPE, wstring(L"EX"), false)); // STRING_OK
	WC_H(MyAddress.DisplayDialog());

	if (S_OK == hRes)
	{
		LPMESSAGE lpMessage = nullptr;
		auto iItem = m_lpContentsTableListCtrl->GetNextItem(
			-1,
			LVNI_SELECTED);
		while (iItem != -1)
		{
			EC_H(OpenItemProp(
				iItem,
				mfcmapiREQUEST_MODIFY,
				reinterpret_cast<LPMAPIPROP*>(&lpMessage)));

			if (lpMessage)
			{
				EC_MAPI(lpMAPISession->MessageOptions(
					reinterpret_cast<ULONG_PTR>(m_hWnd),
					NULL, // API doesn't like Unicode
					LPTSTR(wstringTostring(MyAddress.GetStringW(0)).c_str()),
					lpMessage));

				lpMessage->Release();
				lpMessage = nullptr;
			}

			iItem = m_lpContentsTableListCtrl->GetNextItem(
				iItem,
				LVNI_SELECTED);
			if (S_OK != hRes && -1 != iItem)
			{
				if (bShouldCancel(this, hRes)) break;
				hRes = S_OK;
			}
		}
	}
}

_Check_return_ HRESULT CFolderDlg::OnGetMessageStatus(int /*iItem*/, _In_ SortListData* lpData)
{
	auto hRes = S_OK;
	LPSBinary lpMessageEID = nullptr;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!lpData || !lpData->Contents() || !m_lpContainer) return MAPI_E_INVALID_PARAMETER;

	DebugPrintEx(DBGGeneric, CLASS, L"OnGetMessageStatus", L"\n");

	ULONG ulMessageStatus = NULL;

	lpMessageEID = lpData->Contents()->m_lpEntryID;

	if (lpMessageEID)
	{
		EC_MAPI((static_cast<LPMAPIFOLDER>(m_lpContainer))->GetMessageStatus(
			lpMessageEID->cb,
			reinterpret_cast<LPENTRYID>(lpMessageEID->lpb),
			NULL,
			&ulMessageStatus));

		CEditor MyStatus(
			this,
			IDS_MESSAGESTATUS,
			NULL,
			1,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyStatus.InitPane(0, CreateSingleLinePane(IDS_MESSAGESTATUS, true));
		MyStatus.SetHex(0, ulMessageStatus);

		WC_H(MyStatus.DisplayDialog());
	}

	return hRes;
}

void CFolderDlg::OnSetMessageStatus()
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl || !m_lpContainer) return;

	CEditor MyData(
		this,
		IDS_SETMSGSTATUS,
		IDS_SETMSGSTATUSPROMPT,
		2,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	MyData.InitPane(0, CreateSingleLinePane(IDS_STATUSINHEX, false));
	MyData.InitPane(1, CreateSingleLinePane(IDS_MASKINHEX, false));

	DebugPrintEx(DBGGeneric, CLASS, L"OnSetMessageStatus", L"\n");

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		auto iItem = m_lpContentsTableListCtrl->GetNextItem(
			-1,
			LVNI_SELECTED);
		while (iItem != -1)
		{
			auto lpListData = reinterpret_cast<SortListData*>(m_lpContentsTableListCtrl->GetItemData(iItem));
			if (lpListData && lpListData->Contents())
			{
				auto lpMessageEID = lpListData->Contents()->m_lpEntryID;

				if (lpMessageEID)
				{
					ULONG ulOldStatus = NULL;

					EC_MAPI((static_cast<LPMAPIFOLDER>(m_lpContainer))->SetMessageStatus(
						lpMessageEID->cb,
						reinterpret_cast<LPENTRYID>(lpMessageEID->lpb),
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
				if (bShouldCancel(this, hRes)) break;
				hRes = S_OK;
			}
		}
	}
}

_Check_return_ HRESULT CFolderDlg::OnSubmitMessage(int iItem, _In_ SortListData* /*lpData*/)
{
	auto hRes = S_OK;
	LPMESSAGE lpMessage = nullptr;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric, CLASS, L"OnSubmitMesssage", L"\n");

	if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;

	EC_H(OpenItemProp(
		iItem,
		mfcmapiREQUEST_MODIFY,
		reinterpret_cast<LPMAPIPROP*>(&lpMessage)));

	if (lpMessage)
	{
		// Get subject line of message to copy.
		// This will be used as the new file name.
		EC_MAPI(lpMessage->SubmitMessage(NULL));

		lpMessage->Release();
	}

	return hRes;
}

_Check_return_ HRESULT CFolderDlg::OnAbortSubmit(int iItem, _In_ SortListData* lpData)
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric, CLASS, L"OnSubmitMesssage", L"\n");

	if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;
	if (!m_lpMapiObjects || !lpData || !lpData->Contents()) return MAPI_E_INVALID_PARAMETER;

	auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release

	auto lpMessageEID = lpData->Contents()->m_lpEntryID;

	if (lpMDB && lpMessageEID)
	{
		EC_MAPI(lpMDB->AbortSubmit(
			lpMessageEID->cb,
			reinterpret_cast<LPENTRYID>(lpMessageEID->lpb),
			NULL));
	}

	return hRes;
}

void CFolderDlg::OnDisplayFolder(WORD wMenuSelect)
{
	auto hRes = S_OK;
	auto otType = otDefault;
	switch (wMenuSelect)
	{
	case ID_HIERARCHY:
		otType = otHierarchy; break;
	case ID_CONTENTS:
		otType = otContents; break;
	case ID_HIDDENCONTENTS:
		otType = otAssocContents; break;
	}

	EC_H(DisplayObject(
		m_lpContainer,
		NULL,
		otType,
		this));
}

void CFolderDlg::HandleAddInMenuSingle(
	_In_ LPADDINMENUPARAMS lpParams,
	_In_ LPMAPIPROP lpMAPIProp,
	_In_ LPMAPICONTAINER /*lpContainer*/)
{
	if (lpParams)
	{
		lpParams->lpFolder = static_cast<LPMAPIFOLDER>(m_lpContainer); // m_lpContainer is an LPMAPIFOLDER
		lpParams->lpMessage = static_cast<LPMESSAGE>(lpMAPIProp); // OpenItemProp returns LPMESSAGE
		// Add appropriate flag to context
		if (m_ulDisplayFlags & dfAssoc)
			lpParams->ulCurrentFlags |= MENU_FLAGS_FOLDER_ASSOC;
		if (m_ulDisplayFlags & dfDeleted)
			lpParams->ulCurrentFlags |= MENU_FLAGS_DELETED;
	}

	InvokeAddInMenu(lpParams);
}