// Displays the contents of a folder
#include <StdAfx.h>
#include <UI/Dialogs/ContentsTable/FolderDlg.h>
#include <UI/Controls/ContentsTableListCtrl.h>
#include <MAPI/Cache/MapiObjects.h>
#include <MAPI/MAPIFunctions.h>
#include <MAPI/MAPIStoreFunctions.h>
#include <MAPI/MAPIABFunctions.h>
#include <UI/Controls/SingleMAPIPropListCtrl.h>
#include <MAPI/ColumnTags.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <MAPI/MAPIProcessor/DumpStore.h>
#include <IO/File.h>
#include <UI/Dialogs/ContentsTable/AttachmentsDlg.h>
#include <UI/MAPIFormFunctions.h>
#include <Interpret/InterpretProp.h>
#include <UI/FileDialogEx.h>
#include <Interpret/ExtraPropTags.h>
#include <UI/Dialogs/Editors/PropertyTagEditor.h>
#include <MAPI/MAPIProgress.h>
#include <MAPI/MapiMime.h>
#include <UI/Controls/SortList/ContentsData.h>
#include <MAPI/Cache/GlobalCache.h>
#include <MAPI/MapiMemory.h>

namespace dialog
{
	static std::wstring CLASS = L"CFolderDlg";

	CFolderDlg::CFolderDlg(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ cache::CMapiObjects* lpMapiObjects,
		_In_ LPMAPIPROP lpMAPIFolder,
		ULONG ulDisplayFlags)
		: CContentsTableDlg(
			  pParentWnd,
			  lpMapiObjects,
			  IDS_FOLDER,
			  mfcmapiDO_NOT_CALL_CREATE_DIALOG,
			  lpMAPIFolder,
			  nullptr,
			  LPSPropTagArray(&columns::sptMSGCols),
			  columns::MSGColumns,
			  IDR_MENU_FOLDER_POPUP,
			  MENU_CONTEXT_FOLDER_CONTENTS)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_ulDisplayFlags = ulDisplayFlags;

		m_lpFolder = mapi::safe_cast<LPMAPIFOLDER>(lpMAPIFolder);

		CContentsTableDlg::CreateDialogAndMenu(IDR_MENU_FOLDER);
	}

	CFolderDlg::~CFolderDlg()
	{
		TRACE_DESTRUCTOR(CLASS);
		if (m_lpFolder) m_lpFolder->Release();
	}

	_Check_return_ bool CFolderDlg::HandleMenu(WORD wMenuSelect)
	{
		output::DebugPrint(DBGMenu, L"CFolderDlg::HandleMenu wMenuSelect = 0x%X = %u\n", wMenuSelect, wMenuSelect);
		switch (wMenuSelect)
		{
		case ID_DISPLAYACLTABLE:
			if (m_lpFolder)
			{
				EC_H_S(DisplayExchangeTable(m_lpFolder, PR_ACL_TABLE, otACL, this));
			}

			return true;
		case ID_EXPORTMESSAGES:
			OnExportMessages();
			return true;
		case ID_LOADFROMMSG:
			OnLoadFromMSG();
			return true;
		case ID_LOADFROMTNEF:
			OnLoadFromTNEF();
			return true;
		case ID_LOADFROMEML:
			OnLoadFromEML();
			return true;
		case ID_RESOLVEMESSAGECLASS:
			OnResolveMessageClass();
			return true;
		case ID_SELECTFORM:
			OnSelectForm();
			return true;
		case ID_MANUALRESOLVE:
			OnManualResolve();
			return true;
		case ID_SAVEFOLDERCONTENTSASTEXTFILES:
			OnSaveFolderContentsAsTextFiles();
			return true;
		case ID_SENDBULKMAIL:
			OnSendBulkMail();
			return true;
		case ID_NEW_CUSTOMFORM:
			OnNewCustomForm();
			return true;
		case ID_NEW_MESSAGE:
			OnNewMessage();
			return true;
		case ID_NEW_APPOINTMENT:
		case ID_NEW_CONTACT:
		case ID_NEW_IPMNOTE:
		case ID_NEW_IPMPOST:
		case ID_NEW_TASK:
		case ID_NEW_STICKYNOTE:
			NewSpecialItem(wMenuSelect);
			return true;
		case ID_EXECUTEVERBONFORM:
			OnExecuteVerbOnForm();
			return true;
		case ID_GETPROPSUSINGLONGTERMEID:
			OnGetPropsUsingLongTermEID();
			return true;
		case ID_HIERARCHY:
		case ID_CONTENTS:
		case ID_HIDDENCONTENTS:
			OnDisplayFolder(wMenuSelect);
			return true;
		}

		if (MultiSelectSimple(wMenuSelect)) return true;

		if (MultiSelectComplex(wMenuSelect)) return true;

		return CContentsTableDlg::HandleMenu(wMenuSelect);
	}

	typedef HRESULT (CFolderDlg::*LPSIMPLEMULTI)(int iItem, controls::sortlistdata::SortListData* lpData);

	_Check_return_ bool CFolderDlg::MultiSelectSimple(WORD wMenuSelect)
	{
		LPSIMPLEMULTI lpFunc = nullptr;

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
				auto iItem = m_lpContentsTableListCtrl->GetNextItem(-1, LVNI_SELECTED);
				while (-1 != iItem)
				{
					const auto lpData = m_lpContentsTableListCtrl->GetSortListData(iItem);
					auto hRes = WC_H((this->*lpFunc)(iItem, lpData));
					iItem = m_lpContentsTableListCtrl->GetNextItem(iItem, LVNI_SELECTED);
					if (S_OK != hRes && -1 != iItem)
					{
						if (bShouldCancel(this, hRes)) break;
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
		case ID_ADDTESTADDRESS:
			OnAddOneOffAddress();
			return true;
		case ID_DELETESELECTEDITEM:
			OnDeleteSelectedItem();
			return true;
		case ID_REMOVEONEOFF:
			OnRemoveOneOff();
			return true;
		case ID_RTFSYNC:
			OnRTFSync();
			return true;
		case ID_SAVEMESSAGETOFILE:
			OnSaveMessageToFile();
			return true;
		case ID_SETREADFLAG:
			OnSetReadFlag();
			return true;
		case ID_SETMESSAGESTATUS:
			OnSetMessageStatus();
			return true;
		case ID_GETMESSAGEOPTIONS:
			OnGetMessageOptions();
			return true;
		case ID_DELETEATTACHMENTS:
			OnDeleteAttachments();
			return true;
		case ID_CREATEMESSAGERESTRICTION:
			OnCreateMessageRestriction();
			return true;
		}

		return false;
	}

	void CFolderDlg::OnDisplayItem()
	{
		auto iItem = -1;
		CWaitCursor Wait;

		do
		{
			auto lpMAPIProp = m_lpContentsTableListCtrl->OpenNextSelectedItemProp(&iItem, mfcmapiREQUEST_MODIFY);

			if (lpMAPIProp)
			{
				EC_H_S(DisplayObject(lpMAPIProp, NULL, otDefault, this));
				lpMAPIProp->Release();
			}
		} while (iItem != -1);
	}

	void CFolderDlg::OnInitMenu(_In_ CMenu* pMenu)
	{
		if (pMenu && m_lpContentsTableListCtrl)
		{
			LPMAPISESSION lpMAPISession = nullptr;
			const int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
			const auto ulStatus = cache::CGlobalCache::getInstance().GetBufferStatus();
			pMenu->EnableMenuItem(ID_PASTE, DIM(ulStatus & BUFFER_MESSAGES));

			if (m_lpMapiObjects)
			{
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

			pMenu->EnableMenuItem(ID_DISPLAYACLTABLE, DIM(m_lpFolder));
			pMenu->EnableMenuItem(ID_LOADFROMEML, DIM(m_lpFolder));
			pMenu->EnableMenuItem(ID_LOADFROMMSG, DIM(m_lpFolder));
			pMenu->EnableMenuItem(ID_LOADFROMTNEF, DIM(m_lpFolder));
			pMenu->EnableMenuItem(ID_NEW_APPOINTMENT, DIM(m_lpFolder));
			pMenu->EnableMenuItem(ID_NEW_CONTACT, DIM(m_lpFolder));
			pMenu->EnableMenuItem(ID_NEW_CUSTOMFORM, DIM(m_lpFolder));
			pMenu->EnableMenuItem(ID_NEW_IPMNOTE, DIM(m_lpFolder));
			pMenu->EnableMenuItem(ID_NEW_IPMPOST, DIM(m_lpFolder));
			pMenu->EnableMenuItem(ID_NEW_STICKYNOTE, DIM(m_lpFolder));
			pMenu->EnableMenuItem(ID_NEW_TASK, DIM(m_lpFolder));
			pMenu->EnableMenuItem(ID_NEW_MESSAGE, DIM(m_lpFolder));
			pMenu->EnableMenuItem(ID_SENDBULKMAIL, DIM(m_lpFolder));
			pMenu->EnableMenuItem(ID_SAVEFOLDERCONTENTSASTEXTFILES, DIM(m_lpFolder));
			pMenu->EnableMenuItem(ID_EXPORTMESSAGES, DIM(m_lpFolder));

			pMenu->EnableMenuItem(ID_CONTENTS, DIM(m_lpFolder && !(m_ulDisplayFlags == dfNormal)));
			pMenu->EnableMenuItem(ID_HIDDENCONTENTS, DIM(m_lpFolder && !(m_ulDisplayFlags & dfAssoc)));

			pMenu->EnableMenuItem(
				ID_CREATEMESSAGERESTRICTION,
				DIM(1 == iNumSel && m_lpContentsTableListCtrl->IsContentsTableSet() &&
					m_lpContentsTableListCtrl->GetContainerType() == MAPI_FOLDER));
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
		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		editor::CEditor MyData(this, IDS_ADDONEOFF, NULL, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.InitPane(0, viewpane::TextPane::CreateSingleLinePaneID(IDS_DISPLAYNAME, IDS_DISPLAYNAMEVALUE, false));
		MyData.InitPane(
			1, viewpane::TextPane::CreateSingleLinePane(IDS_ADDRESSTYPE, std::wstring(L"EX"), false)); // STRING_OK
		MyData.InitPane(2, viewpane::TextPane::CreateSingleLinePaneID(IDS_ADDRESS, IDS_ADDRESSVALUE, false));
		MyData.InitPane(3, viewpane::TextPane::CreateSingleLinePane(IDS_RECIPTYPE, false));
		MyData.SetHex(3, MAPI_TO);
		MyData.InitPane(4, viewpane::TextPane::CreateSingleLinePane(IDS_RECIPCOUNT, false));
		MyData.SetDecimal(4, 1);

		if (!MyData.DisplayDialog()) return;

		const auto displayName = MyData.GetStringW(0);
		const auto addressType = MyData.GetStringW(1);
		const auto address = MyData.GetStringW(2);
		const auto recipientType = MyData.GetHex(3);
		const auto count = MyData.GetDecimal(4);
		auto iItem = m_lpContentsTableListCtrl->GetNextItem(-1, LVNI_SELECTED);
		while (iItem != -1)
		{
			auto lpMessage = OpenMessage(iItem, mfcmapiREQUEST_MODIFY);
			if (lpMessage)
			{
				if (count <= 1)
				{
					hRes = EC_H(mapi::ab::AddOneOffAddress(
						lpMAPISession, lpMessage, displayName, addressType, address, recipientType));
				}
				else
				{
					for (ULONG i = 0; i < count; i++)
					{
						const auto countedDisplayName = displayName + std::to_wstring(i);
						hRes = EC_H(mapi::ab::AddOneOffAddress(
							lpMAPISession, lpMessage, countedDisplayName, addressType, address, recipientType));
					}
				}

				lpMessage->Release();
				lpMessage = nullptr;
			}

			iItem = m_lpContentsTableListCtrl->GetNextItem(iItem, LVNI_SELECTED);
			if (S_OK != hRes && -1 != iItem)
			{
				if (bShouldCancel(this, hRes)) break;
			}
		}
	}

	_Check_return_ HRESULT
	CFolderDlg::OnAttachmentProperties(int iItem, _In_ controls::sortlistdata::SortListData* /*lpData*/)
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;

		auto hRes = S_OK;
		auto lpMessage = OpenMessage(iItem, mfcmapiREQUEST_MODIFY);
		if (lpMessage)
		{
			hRes = EC_H(OpenAttachmentsFromMessage(lpMessage));

			lpMessage->Release();
		}

		return hRes;
	}

	void CFolderDlg::HandleCopy()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		output::DebugPrintEx(DBGGeneric, CLASS, L"HandleCopy", L"\n");
		if (!m_lpContentsTableListCtrl) return;

		const auto lpEIDs = m_lpContentsTableListCtrl->GetSelectedItemEIDs();

		// CGlobalCache takes over ownership of lpEIDs - don't free now
		cache::CGlobalCache::getInstance().SetMessagesToCopy(lpEIDs, m_lpFolder);
	}

	_Check_return_ bool CFolderDlg::HandlePaste()
	{
		if (CBaseDialog::HandlePaste()) return true;

		auto hRes = S_OK;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		output::DebugPrintEx(DBGGeneric, CLASS, L"HandlePaste", L"\n");
		if (!m_lpFolder) return false;

		editor::CEditor MyData(this, IDS_COPYMESSAGE, IDS_COPYMESSAGEPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.InitPane(0, viewpane::CheckPane::Create(IDS_MESSAGEMOVE, false, false));
		UINT uidDropDown[] = {IDS_DDCOPYMESSAGES, IDS_DDCOPYTO};
		MyData.InitPane(1, viewpane::DropDownPane::Create(IDS_COPYINTERFACE, _countof(uidDropDown), uidDropDown, true));

		if (!MyData.DisplayDialog()) return false;

		const auto lpEIDs = cache::CGlobalCache::getInstance().GetMessagesToCopy();
		auto lpMAPISourceFolder = cache::CGlobalCache::getInstance().GetSourceParentFolder();
		auto ulMoveMessage = MyData.GetCheck(0) ? MESSAGE_MOVE : 0;

		if (lpEIDs && lpMAPISourceFolder)
		{
			if (0 == MyData.GetDropDown(1))
			{ // CopyMessages
				LPMAPIPROGRESS lpProgress =
					mapi::mapiui::GetMAPIProgress(L"IMAPIFolder::CopyMessages", m_hWnd); // STRING_OK

				if (lpProgress) ulMoveMessage |= MESSAGE_DIALOG;

				EC_MAPI_S(lpMAPISourceFolder->CopyMessages(
					lpEIDs,
					&IID_IMAPIFolder,
					m_lpFolder,
					lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL,
					lpProgress,
					ulMoveMessage));

				if (lpProgress) lpProgress->Release();
			}
			else
			{ // CopyTo
				// Specify properties to exclude in the copy operation. These are
				// the properties that Exchange excludes to save bits and time.
				// Should not be necessary to exclude these, but speeds the process
				// when a lot of messages are being copied.
				static const SizedSPropTagArray(7, excludeTags) = {7,
																   {PR_ACCESS,
																	PR_BODY,
																	PR_RTF_SYNC_BODY_COUNT,
																	PR_RTF_SYNC_BODY_CRC,
																	PR_RTF_SYNC_BODY_TAG,
																	PR_RTF_SYNC_PREFIX_COUNT,
																	PR_RTF_SYNC_TRAILING_COUNT}};

				const auto lpTagsToExclude = mapi::GetExcludedTags(LPSPropTagArray(&excludeTags), m_lpFolder, m_bIsAB);
				if (lpTagsToExclude)
				{
					for (ULONG i = 0; i < lpEIDs->cValues; i++)
					{
						LPMESSAGE lpMessage = nullptr;

						EC_H_S(mapi::CallOpenEntry(
							nullptr,
							nullptr,
							lpMAPISourceFolder,
							nullptr,
							lpEIDs->lpbin[i].cb,
							reinterpret_cast<LPENTRYID>(lpEIDs->lpbin[i].lpb),
							nullptr,
							MyData.GetCheck(0) ? MAPI_MODIFY : 0, // only need write access for moves
							nullptr,
							reinterpret_cast<LPUNKNOWN*>(&lpMessage)));
						if (lpMessage)
						{
							LPMESSAGE lpNewMessage = nullptr;
							EC_MAPI_S(m_lpFolder->CreateMessage(
								nullptr, m_ulDisplayFlags & dfAssoc ? MAPI_ASSOCIATED : NULL, &lpNewMessage));
							if (lpNewMessage)
							{
								LPSPropProblemArray lpProblems = nullptr;

								// copy message properties to IMessage object opened on top of IStorage.
								LPMAPIPROGRESS lpProgress =
									mapi::mapiui::GetMAPIProgress(L"IMAPIProp::CopyTo", m_hWnd); // STRING_OK

								if (lpProgress) ulMoveMessage |= MAPI_DIALOG;

								hRes = EC_MAPI(lpMessage->CopyTo(
									0,
									nullptr, // TODO: interfaces?
									lpTagsToExclude,
									lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL, // UI param
									lpProgress, // progress
									const_cast<LPIID>(&IID_IMessage),
									lpNewMessage,
									ulMoveMessage,
									&lpProblems));

								if (lpProgress) lpProgress->Release();

								EC_PROBLEMARRAY(lpProblems);
								MAPIFreeBuffer(lpProblems);

								if (SUCCEEDED(hRes))
								{
									// save changes to IMessage object.
									hRes = EC_MAPI(lpNewMessage->SaveChanges(KEEP_OPEN_READWRITE));
								}

								lpNewMessage->Release();

								if (SUCCEEDED(hRes) &&
									MyData.GetCheck(0)) // if we moved, save changes on original message
								{
									hRes = EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
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

		return true;
	}

	void CFolderDlg::OnDeleteAttachments()
	{
		editor::CEditor MyData(
			this, IDS_DELETEATTACHMENTS, IDS_DELETEATTACHMENTSPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_FILENAME, false));

		if (!MyData.DisplayDialog()) return;

		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		auto szAttName = MyData.GetStringW(0);
		if (!szAttName.empty())
		{
			auto iItem = m_lpContentsTableListCtrl->GetNextItem(-1, LVNI_SELECTED);

			while (-1 != iItem)
			{
				auto lpMessage = OpenMessage(iItem, mfcmapiREQUEST_MODIFY);
				if (lpMessage)
				{
					EC_H_S(file::DeleteAttachments(lpMessage, szAttName, m_hWnd));

					lpMessage->Release();
				}

				iItem = m_lpContentsTableListCtrl->GetNextItem(iItem, LVNI_SELECTED);
			}
		}
	}

	void CFolderDlg::OnDeleteSelectedItem()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpMapiObjects || !m_lpContentsTableListCtrl || !m_lpFolder) return;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		auto lpMDB = mapi::store::OpenDefaultMessageStore(lpMAPISession);
		if (!lpMDB) return;

		auto bDelete = false;
		auto bMove = false;
		auto ulFlag = MESSAGE_DIALOG;

		if (m_ulDisplayFlags & dfDeleted)
		{
			ulFlag |= DELETE_HARD_DELETE;
			bDelete = true;
		}
		else
		{
			const auto bShift = !(GetKeyState(VK_SHIFT) < 0);

			editor::CEditor MyData(
				this, IDS_DELETEITEM, IDS_DELETEITEMPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			UINT uidDropDown[] = {IDS_DDDELETETODELETED, IDS_DDDELETETORETENTION, IDS_DDDELETEHARDDELETE};

			if (bShift)
				MyData.InitPane(
					0, viewpane::DropDownPane::Create(IDS_DELSTYLE, _countof(uidDropDown), uidDropDown, true));
			else
				MyData.InitPane(
					0, viewpane::DropDownPane::Create(IDS_DELSTYLE, _countof(uidDropDown) - 1, &uidDropDown[1], true));

			if (MyData.DisplayDialog())
			{
				bDelete = true;

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
		}

		if (bDelete)
		{
			const auto lpEIDs = m_lpContentsTableListCtrl->GetSelectedItemEIDs();

			if (bMove)
			{
				EC_H_S(mapi::DeleteToDeletedItems(lpMDB, m_lpFolder, lpEIDs, m_hWnd));
			}
			else
			{
				LPMAPIPROGRESS lpProgress =
					mapi::mapiui::GetMAPIProgress(L"IMAPIFolder::DeleteMessages", m_hWnd); // STRING_OK

				if (lpProgress) ulFlag |= MESSAGE_DIALOG;

				EC_MAPI_S(m_lpFolder->DeleteMessages(
					lpEIDs, // list of messages to delete
					lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL,
					lpProgress,
					ulFlag));

				if (lpProgress) lpProgress->Release();
			}

			MAPIFreeBuffer(lpEIDs);
		}

		lpMDB->Release();
	}

	void CFolderDlg::OnGetPropsUsingLongTermEID() const
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

		const auto lpListData = m_lpContentsTableListCtrl->GetFirstSelectedItemData();
		if (lpListData && lpListData->Contents())
		{
			const auto lpMessageEID = lpListData->Contents()->m_lpLongtermID;

			if (lpMessageEID)
			{
				LPMAPIPROP lpMAPIProp = nullptr;

				const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
				if (lpMDB)
				{
					WC_H_S(mapi::CallOpenEntry(
						lpMDB,
						nullptr,
						nullptr,
						nullptr,
						lpMessageEID->cb,
						reinterpret_cast<LPENTRYID>(lpMessageEID->lpb),
						nullptr,
						MAPI_BEST_ACCESS,
						nullptr,
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
		if (!m_lpFolder) return;

		LPMESSAGE lpNewMessage = nullptr;

		auto files = file::CFileDialogExW::OpenFiles(
			L"msg", // STRING_OK
			strings::emptystring,
			OFN_HIDEREADONLY | OFN_FILEMUSTEXIST,
			strings::loadstring(IDS_MSGFILES),
			this);

		if (!files.empty())
		{
			editor::CEditor MyData(this, IDS_LOADMSG, IDS_LOADMSGPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

			UINT uidDropDown[] = {IDS_DDLOADTOFOLDER, IDS_DDDISPLAYPROPSONLY};
			MyData.InitPane(0, viewpane::DropDownPane::Create(IDS_LOADSTYLE, _countof(uidDropDown), uidDropDown, true));

			if (!MyData.DisplayDialog()) return;

			for (auto& lpszPath : files)
			{
				switch (MyData.GetDropDown(0))
				{
				case 0:
					EC_MAPI_S(m_lpFolder->CreateMessage(
						nullptr, m_ulDisplayFlags & dfAssoc ? MAPI_ASSOCIATED : NULL, &lpNewMessage));

					if (lpNewMessage)
					{
						EC_H_S(file::LoadFromMSG(lpszPath, lpNewMessage, m_hWnd));
					}

					break;
				case 1:
					if (m_lpPropDisplay)
					{
						EC_H_S(file::LoadMSGToMessage(lpszPath, &lpNewMessage));

						if (lpNewMessage)
						{
							EC_H_S(m_lpPropDisplay->SetDataSource(lpNewMessage, nullptr, false));
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

	void CFolderDlg::OnResolveMessageClass()
	{
		if (!m_lpMapiObjects) return;

		LPMAPIFORMINFO lpMAPIFormInfo = nullptr;
		ResolveMessageClass(m_lpMapiObjects, m_lpFolder, &lpMAPIFormInfo);
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

		SelectForm(m_hWnd, m_lpMapiObjects, m_lpFolder, &lpMAPIFormInfo);
		if (lpMAPIFormInfo)
		{
			OnUpdateSingleMAPIPropListCtrl(lpMAPIFormInfo, nullptr);
			// TODO: Put some code in here which works with the returned Form Info pointer
			lpMAPIFormInfo->Release();
		}
	}

	void CFolderDlg::OnManualResolve()
	{
		auto iItem = -1;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		editor::CPropertyTagEditor MyPropertyTag(
			IDS_MANUALRESOLVE,
			NULL, // prompt
			NULL,
			m_bIsAB,
			m_lpFolder,
			this);

		if (!MyPropertyTag.DisplayDialog()) return;

		editor::CEditor MyData(
			this, IDS_MANUALRESOLVE, IDS_MANUALRESOLVEPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_DISPLAYNAME, false));

		if (!MyData.DisplayDialog()) return;
		do
		{
			auto lpMAPIProp = m_lpContentsTableListCtrl->OpenNextSelectedItemProp(&iItem, mfcmapiREQUEST_MODIFY);
			auto lpMessage = mapi::safe_cast<LPMESSAGE>(lpMAPIProp);

			if (lpMessage)
			{
				const auto name = MyData.GetStringW(0);
				EC_H_S(mapi::ab::ManualResolve(
					lpMAPISession, lpMessage, name, CHANGE_PROP_TYPE(MyPropertyTag.GetPropertyTag(), PT_UNICODE)));

				lpMessage->Release();
			}

			if (lpMAPIProp) lpMAPIProp->Release();
		} while (iItem != -1);
	}

	void CFolderDlg::NewSpecialItem(WORD wMenuSelect) const
	{
		LPMAPIFOLDER lpFolder = nullptr; // DO NOT RELEASE

		if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

		const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
		if (!lpMDB) return;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release

		if (lpMAPISession)
		{
			ULONG ulFolder = NULL;
			std::wstring szClass;

			switch (wMenuSelect)
			{
			case ID_NEW_APPOINTMENT:
				ulFolder = mapi::DEFAULT_CALENDAR;
				szClass = L"IPM.APPOINTMENT"; // STRING_OK
				break;
			case ID_NEW_CONTACT:
				ulFolder = mapi::DEFAULT_CONTACTS;
				szClass = L"IPM.CONTACT"; // STRING_OK
				break;
			case ID_NEW_IPMNOTE:
				szClass = L"IPM.NOTE"; // STRING_OK
				break;
			case ID_NEW_IPMPOST:
				szClass = L"IPM.POST"; // STRING_OK
				break;
			case ID_NEW_TASK:
				ulFolder = mapi::DEFAULT_TASKS;
				szClass = L"IPM.TASK"; // STRING_OK
				break;
			case ID_NEW_STICKYNOTE:
				ulFolder = mapi::DEFAULT_NOTES;
				szClass = L"IPM.STICKYNOTE"; // STRING_OK
				break;
			}

			LPMAPIFOLDER lpSpecialFolder = nullptr;
			if (ulFolder)
			{
				lpSpecialFolder = mapi::OpenDefaultFolder(ulFolder, lpMDB);
				lpFolder = lpSpecialFolder;
			}
			else
			{
				lpFolder = m_lpFolder;
			}

			if (lpFolder)
			{
				EC_H_S(mapi::mapiui::CreateAndDisplayNewMailInFolder(
					m_hWnd, lpMDB, lpMAPISession, m_lpContentsTableListCtrl, -1, szClass, lpFolder));
			}

			if (lpSpecialFolder) lpSpecialFolder->Release();
		}
	}

	void CFolderDlg::OnNewMessage()
	{
		LPMESSAGE lpMessage = nullptr;

		auto hRes =
			EC_MAPI(m_lpFolder->CreateMessage(nullptr, m_ulDisplayFlags & dfAssoc ? MAPI_ASSOCIATED : 0, &lpMessage));

		if (SUCCEEDED(hRes) && lpMessage)
		{
			hRes = EC_MAPI(lpMessage->SaveChanges(NULL));
			lpMessage->Release();
		}
	}

	void CFolderDlg::OnNewCustomForm()
	{
		auto hRes = S_OK;

		if (!m_lpMapiObjects || !m_lpFolder || !m_lpContentsTableListCtrl) return;

		const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
		if (!lpMDB) return;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release

		if (lpMAPISession)
		{
			editor::CEditor MyPrompt1(
				this, IDS_NEWCUSTOMFORM, IDS_NEWCUSTOMFORMPROMPT1, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			UINT uidDropDown[] = {IDS_DDENTERFORMCLASS, IDS_DDFOLDERFORMLIBRARY, IDS_DDORGFORMLIBRARY};
			MyPrompt1.InitPane(
				0, viewpane::DropDownPane::Create(IDS_LOCATIONOFFORM, _countof(uidDropDown), uidDropDown, true));

			if (!MyPrompt1.DisplayDialog()) return;
			std::wstring szClass;

			editor::CEditor MyClass(
				this, IDS_NEWCUSTOMFORM, IDS_NEWCUSTOMFORMPROMPT2, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyClass.InitPane(
				0,
				viewpane::TextPane::CreateSingleLinePane(IDS_FORMTYPE, std::wstring(L"IPM.Note"), false)); // STRING_OK

			switch (MyPrompt1.GetDropDown(0))
			{
			case 0:
				if (!MyClass.DisplayDialog()) return;
				szClass = MyClass.GetStringW(0);

				break;
			case 1:
			case 2:
			{
				LPMAPIFORMMGR lpMAPIFormMgr = nullptr;
				LPMAPIFORMINFO lpMAPIFormInfo = nullptr;
				hRes = EC_MAPI(MAPIOpenFormMgr(lpMAPISession, &lpMAPIFormMgr));

				if (lpMAPIFormMgr)
				{
					LPMAPIFOLDER lpFormSource = nullptr;
					if (1 == MyPrompt1.GetDropDown(0)) lpFormSource = m_lpFolder;
					const auto szTitle = strings::loadstring(IDS_SELECTFORMCREATE);

					// Apparently, SelectForm doesn't support unicode
					hRes = EC_H_CANCEL(lpMAPIFormMgr->SelectForm(
						reinterpret_cast<ULONG_PTR>(m_hWnd),
						0,
						reinterpret_cast<LPCTSTR>(strings::wstringTostring(szTitle).c_str()),
						lpFormSource,
						&lpMAPIFormInfo));

					if (lpMAPIFormInfo)
					{
						LPSPropValue lpProp = nullptr;
						hRes = EC_MAPI(HrGetOneProp(lpMAPIFormInfo, PR_MESSAGE_CLASS_W, &lpProp));
						if (mapi::CheckStringProp(lpProp, PT_UNICODE))
						{
							szClass = lpProp->Value.lpszW;
						}

						MAPIFreeBuffer(lpProp);
						lpMAPIFormInfo->Release();
					}

					lpMAPIFormMgr->Release();
				}
			}

			break;
			}

			if (SUCCEEDED(hRes))
			{
				EC_H_S(mapi::mapiui::CreateAndDisplayNewMailInFolder(
					m_hWnd, lpMDB, lpMAPISession, m_lpContentsTableListCtrl, -1, szClass, m_lpFolder));
			}
		}
	}

	_Check_return_ HRESULT CFolderDlg::OnOpenModal(int iItem, _In_ controls::sortlistdata::SortListData* /*lpData*/)
	{
		auto hRes = S_OK;
		LPMESSAGE lpMessage = nullptr;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;
		if (!m_lpMapiObjects || !m_lpFolder) return MAPI_E_INVALID_PARAMETER;

		const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
		if (!lpMDB) return MAPI_E_CALL_FAILED;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (lpMAPISession)
		{
			// Before we open the message, make sure the MAPI Form Manager is implemented
			LPMAPIFORMMGR lpFormMgr = nullptr;
			hRes = WC_MAPI(MAPIOpenFormMgr(lpMAPISession, &lpFormMgr));
			if (lpFormMgr)
			{
				lpMessage = OpenMessage(iItem, mfcmapiREQUEST_MODIFY);
				if (lpMessage)
				{
					hRes = EC_H(mapi::mapiui::OpenMessageModal(m_lpFolder, lpMAPISession, lpMDB, lpMessage));

					lpMessage->Release();
				}

				lpFormMgr->Release();
			}
		}

		return hRes;
	}

	_Check_return_ HRESULT CFolderDlg::OnOpenNonModal(int iItem, _In_ controls::sortlistdata::SortListData* /*lpData*/)
	{
		auto hRes = S_OK;
		LPMESSAGE lpMessage = nullptr;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;
		if (!m_lpMapiObjects || !m_lpContentsTableListCtrl || !m_lpFolder) return MAPI_E_INVALID_PARAMETER;

		const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
		if (!lpMDB) return MAPI_E_CALL_FAILED;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (lpMAPISession)
		{
			// Before we open the message, make sure the MAPI Form Manager is implemented
			LPMAPIFORMMGR lpFormMgr = nullptr;
			hRes = WC_MAPI(MAPIOpenFormMgr(lpMAPISession, &lpFormMgr));

			if (lpFormMgr)
			{
				lpMessage = OpenMessage(iItem, mfcmapiREQUEST_MODIFY);
				if (lpMessage)
				{
					hRes = WC_H(mapi::mapiui::OpenMessageNonModal(
						m_hWnd,
						lpMDB,
						lpMAPISession,
						m_lpFolder,
						m_lpContentsTableListCtrl,
						iItem,
						lpMessage,
						EXCHIVERB_OPEN,
						nullptr));

					lpMessage->Release();
				}

				lpFormMgr->Release();
			}
		}

		return hRes;
	}

	void CFolderDlg::OnExecuteVerbOnForm()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

		const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
		if (!lpMDB) return;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (lpMAPISession)
		{
			editor::CEditor MyData(
				this, IDS_EXECUTEVERB, IDS_EXECUTEVERBPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyData.SetPromptPostFix(interpretprop::AllFlagsToString(PROP_ID(PR_LAST_VERB_EXECUTED), false));

			MyData.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_VERB, false));
			MyData.SetDecimal(0, EXCHIVERB_OPEN);

			if (!MyData.DisplayDialog()) return;

			const auto iItem = m_lpContentsTableListCtrl->GetNextItem(-1, LVNI_SELECTED);
			if (iItem != -1)
			{
				auto lpMessage = OpenMessage(iItem, mfcmapiREQUEST_MODIFY);
				if (lpMessage)
				{
					EC_H_S(mapi::mapiui::OpenMessageNonModal(
						m_hWnd,
						lpMDB,
						lpMAPISession,
						m_lpFolder,
						m_lpContentsTableListCtrl,
						iItem,
						lpMessage,
						MyData.GetDecimal(0),
						nullptr));

					lpMessage->Release();
				}
			}
		}
	}

	_Check_return_ HRESULT
	CFolderDlg::OnResendSelectedItem(int /*iItem*/, _In_ controls::sortlistdata::SortListData* lpData)
	{
		auto hRes = S_OK;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!lpData || !lpData->Contents() || !m_lpFolder) return MAPI_E_INVALID_PARAMETER;

		if (lpData->Contents()->m_lpEntryID)
		{
			hRes = EC_H(mapi::ResendSingleMessage(m_lpFolder, lpData->Contents()->m_lpEntryID, m_hWnd));
		}

		return hRes;
	}

	_Check_return_ HRESULT
	CFolderDlg::OnRecipientProperties(int iItem, _In_ controls::sortlistdata::SortListData* /*lpData*/)
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;

		auto hRes = S_OK;
		auto lpMessage = OpenMessage(iItem, mfcmapiREQUEST_MODIFY);
		if (lpMessage)
		{
			hRes = EC_H(OpenRecipientsFromMessage(lpMessage));

			lpMessage->Release();
		}

		return hRes;
	}

	void CFolderDlg::OnRemoveOneOff()
	{
		auto hRes = S_OK;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpContentsTableListCtrl) return;

		editor::CEditor MyData(
			this, IDS_REMOVEONEOFF, IDS_REMOVEONEOFFPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, viewpane::CheckPane::Create(IDS_REMOVEPROPDEFSTREAM, true, false));

		if (!MyData.DisplayDialog()) return;

		auto iItem = m_lpContentsTableListCtrl->GetNextItem(-1, LVNI_SELECTED);
		while (iItem != -1)
		{
			auto lpMessage = OpenMessage(iItem, mfcmapiREQUEST_MODIFY);
			if (lpMessage)
			{
				output::DebugPrint(
					DBGGeneric,
					L"Calling RemoveOneOff on %p, %wsremoving property definition stream\n",
					lpMessage,
					MyData.GetCheck(0) ? L"" : L"not ");
				hRes = EC_H(mapi::RemoveOneOff(lpMessage, MyData.GetCheck(0)));

				lpMessage->Release();
			}

			iItem = m_lpContentsTableListCtrl->GetNextItem(iItem, LVNI_SELECTED);
			if (hRes != S_OK && -1 != iItem)
			{
				if (bShouldCancel(this, hRes)) break;
			}
		}
	}

#define RTF_SYNC_HTML_CHANGED ((ULONG) 0x00000004)

	void CFolderDlg::OnRTFSync()
	{
		auto hRes = S_OK;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpContentsTableListCtrl) return;

		editor::CEditor MyData(this, IDS_CALLRTFSYNC, IDS_CALLRTFSYNCPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_FLAGS, false));
		MyData.SetHex(0, RTF_SYNC_RTF_CHANGED);

		if (MyData.DisplayDialog())
		{
			BOOL bMessageUpdated = false;

			auto iItem = m_lpContentsTableListCtrl->GetNextItem(-1, LVNI_SELECTED);
			while (iItem != -1)
			{
				auto lpMessage = OpenMessage(iItem, mfcmapiREQUEST_MODIFY);
				if (lpMessage)
				{
					output::DebugPrint(
						DBGGeneric, L"Calling RTFSync on %p with flags 0x%X\n", lpMessage, MyData.GetHex(0));
					hRes = EC_MAPI(RTFSync(lpMessage, MyData.GetHex(0), &bMessageUpdated));
					output::DebugPrint(DBGGeneric, L"RTFSync returned %d\n", bMessageUpdated);

					if (SUCCEEDED(hRes))
					{
						hRes = EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
					}

					lpMessage->Release();
					lpMessage = nullptr;
				}

				iItem = m_lpContentsTableListCtrl->GetNextItem(iItem, LVNI_SELECTED);
				if (S_OK != hRes && -1 != iItem)
				{
					if (bShouldCancel(this, hRes)) break;
				}
			}
		}
	}

	_Check_return_ HRESULT
	CFolderDlg::OnSaveAttachments(int iItem, _In_ controls::sortlistdata::SortListData* /*lpData*/)
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;

		auto hRes = S_OK;
		auto lpMessage = OpenMessage(iItem, mfcmapiREQUEST_MODIFY);
		if (lpMessage)
		{
			hRes = EC_H(file::WriteAttachmentsToFile(lpMessage, m_hWnd));

			lpMessage->Release();
		}

		return hRes;
	}

	void CFolderDlg::OnSaveFolderContentsAsTextFiles()
	{
		if (!m_lpMapiObjects || !m_lpFolder) return;

		const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
		if (!lpMDB) return;

		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		file::SaveFolderContentsToTXT(
			lpMDB, m_lpFolder, (m_ulDisplayFlags & dfAssoc) == 0, (m_ulDisplayFlags & dfAssoc) != 0, false, m_hWnd);
	}

	void CFolderDlg::OnExportMessages()
	{
		const auto lpFolder = m_lpFolder;

		if (lpFolder)
		{
			file::ExportMessages(lpFolder, m_hWnd);
		}
	}

	void CFolderDlg::OnSaveMessageToFile()
	{
		auto hRes = S_OK;

		output::DebugPrintEx(DBGGeneric, CLASS, L"OnSaveMessageToFile", L"\n");

		editor::CEditor MyData(
			this, IDS_SAVEMESSAGETOFILE, IDS_SAVEMESSAGETOFILEPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		UINT uidDropDown[] = {IDS_DDTEXTFILE,
							  IDS_DDMSGFILEANSI,
							  IDS_DDMSGFILEUNICODE,
							  IDS_DDEMLFILE,
							  IDS_DDEMLFILEUSINGICONVERTERSESSION,
							  IDS_DDTNEFFILE};
		MyData.InitPane(
			0, viewpane::DropDownPane::Create(IDS_FORMATTOSAVEMESSAGE, _countof(uidDropDown), uidDropDown, true));
		const auto numSelected = m_lpContentsTableListCtrl->GetSelectedCount();
		if (numSelected > 1)
		{
			MyData.InitPane(1, viewpane::CheckPane::Create(IDS_EXPORTPROMPTLOCATION, false, false));
		}

		if (!MyData.DisplayDialog()) return;

		LPCWSTR szExt = nullptr;
		LPCWSTR szDotExt = nullptr;
		std::wstring szFilter;
		LPADRBOOK lpAddrBook = nullptr;
		switch (MyData.GetDropDown(0))
		{
		case 0:
			szExt = L"xml"; // STRING_OK
			szDotExt = L".xml"; // STRING_OK
			szFilter = strings::loadstring(IDS_XMLFILES);
			break;
		case 1:
		case 2:
			szExt = L"msg"; // STRING_OK
			szDotExt = L".msg"; // STRING_OK
			szFilter = strings::loadstring(IDS_MSGFILES);
			break;
		case 3:
		case 4:
			szExt = L"eml"; // STRING_OK
			szDotExt = L".eml"; // STRING_OK
			szFilter = strings::loadstring(IDS_EMLFILES);
			break;
		case 5:
			szExt = L"tnef"; // STRING_OK
			szDotExt = L".tnef"; // STRING_OK
			szFilter = strings::loadstring(IDS_TNEFFILES);

			lpAddrBook = m_lpMapiObjects->GetAddrBook(true); // do not release
			break;
		default:
			break;
		}

		std::wstring dir;
		const auto bPrompt = numSelected == 1 || MyData.GetCheck(1);
		if (!bPrompt)
		{
			// If we weren't asked to prompt for each item, we still need to ask for a directory
			dir = file::GetDirectoryPath(m_hWnd);
		}

		auto iItem = m_lpContentsTableListCtrl->GetNextItem(-1, LVNI_SELECTED);
		while (-1 != iItem)
		{
			auto lpMessage = OpenMessage(iItem, mfcmapiREQUEST_MODIFY);
			if (lpMessage)
			{
				auto filename = file::BuildFileName(szDotExt, dir, lpMessage);
				output::DebugPrint(DBGGeneric, L"BuildFileName built file name \"%ws\"\n", filename.c_str());

				if (bPrompt)
				{
					filename = file::CFileDialogExW::SaveAs(
						szExt, filename, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, szFilter, this);
				}

				if (!filename.empty())
				{
					switch (MyData.GetDropDown(0))
					{
					case 0:
						// Idea is to capture anything that may be important about this message to disk so it can be analyzed.
						{
							mapiprocessor::CDumpStore MyDumpStore;
							MyDumpStore.InitMessagePath(filename);
							// Just assume this message might have attachments
							MyDumpStore.ProcessMessage(lpMessage, true, nullptr);
						}

						break;
					case 1:
						hRes = EC_H(file::SaveToMSG(lpMessage, filename, false, m_hWnd, true));
						break;
					case 2:
						hRes = EC_H(file::SaveToMSG(lpMessage, filename, true, m_hWnd, true));
						break;
					case 3:
						hRes = EC_H(file::SaveToEML(lpMessage, filename));
						break;
					case 4:
					{
						ULONG ulConvertFlags = CCSF_SMTP;
						auto et = IET_UNKNOWN;
						auto mst = USE_DEFAULT_SAVETYPE;
						ULONG ulWrapLines = USE_DEFAULT_WRAPPING;
						auto bDoAdrBook = false;

						hRes = EC_H(mapi::mapimime::GetConversionToEMLOptions(
							this, &ulConvertFlags, &et, &mst, &ulWrapLines, &bDoAdrBook));
						if (hRes == S_OK)
						{
							LPADRBOOK lpAdrBook = nullptr;
							if (bDoAdrBook) lpAdrBook = m_lpMapiObjects->GetAddrBook(true); // do not release

							hRes = EC_H(mapi::mapimime::ExportIMessageToEML(
								lpMessage, filename.c_str(), ulConvertFlags, et, mst, ulWrapLines, lpAdrBook));
						}
					}

					break;
					case 5:
						hRes = EC_H(file::SaveToTNEF(lpMessage, lpAddrBook, filename));
						break;
					default:
						break;
					}
				}
				else
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

			iItem = m_lpContentsTableListCtrl->GetNextItem(iItem, LVNI_SELECTED);
			if (hRes != S_OK && -1 != iItem)
			{
				if (bShouldCancel(this, hRes)) break;
			}
		}
	}

	// Use CFileDialogExW to locate a .DAT or .TNEF file to load,
	// Pass the file name and a message to load in to LoadFromTNEF to do the work.
	void CFolderDlg::OnLoadFromTNEF()
	{
		LPMESSAGE lpNewMessage = nullptr;

		if (!m_lpFolder) return;

		const auto lpAddrBook = m_lpMapiObjects->GetAddrBook(true); // do not release
		if (lpAddrBook)
		{
			auto files = file::CFileDialogExW::OpenFiles(
				L"tnef", // STRING_OK
				strings::emptystring,
				OFN_FILEMUSTEXIST,
				strings::loadstring(IDS_TNEFFILES),
				this);

			if (!files.empty())
			{
				for (auto& lpszPath : files)
				{
					EC_MAPI_S(m_lpFolder->CreateMessage(
						nullptr, m_ulDisplayFlags & dfAssoc ? MAPI_ASSOCIATED : 0, &lpNewMessage));

					if (lpNewMessage)
					{
						EC_H_S(file::LoadFromTNEF(lpszPath, lpAddrBook, lpNewMessage));

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
		if (!m_lpFolder) return;

		LPMESSAGE lpNewMessage = nullptr;

		ULONG ulConvertFlags = CCSF_SMTP;
		auto bDoAdrBook = false;
		auto bDoApply = false;
		HCHARSET hCharSet = nullptr;
		auto cSetApplyType = CSET_APPLY_UNTAGGED;
		auto hRes = WC_H(mapi::mapimime::GetConversionFromEMLOptions(
			this, &ulConvertFlags, &bDoAdrBook, &bDoApply, &hCharSet, &cSetApplyType, nullptr));
		if (hRes == S_OK)
		{
			LPADRBOOK lpAdrBook = nullptr;
			if (bDoAdrBook) lpAdrBook = m_lpMapiObjects->GetAddrBook(true); // do not release

			auto files = file::CFileDialogExW::OpenFiles(
				L"eml", // STRING_OK
				strings::emptystring,
				OFN_FILEMUSTEXIST,
				strings::loadstring(IDS_EMLFILES),
				this);

			if (!files.empty())
			{
				for (auto& lpszPath : files)
				{
					EC_MAPI_S(m_lpFolder->CreateMessage(
						nullptr, m_ulDisplayFlags & dfAssoc ? MAPI_ASSOCIATED : 0, &lpNewMessage));

					if (lpNewMessage)
					{
						EC_H_S(mapi::mapimime::ImportEMLToIMessage(
							lpszPath.c_str(),
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

		editor::CEditor MyData(
			this, IDS_SENDBULKMAIL, IDS_SENDBULKMAILPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_NUMMESSAGES, false));
		MyData.InitPane(1, viewpane::TextPane::CreateSingleLinePane(IDS_RECIPNAME, false));
		MyData.InitPane(2, viewpane::TextPane::CreateSingleLinePane(IDS_SUBJECT, false));
		MyData.InitPane(
			3, viewpane::TextPane::CreateSingleLinePane(IDS_CLASS, std::wstring(L"IPM.Note"), false)); // STRING_OK
		MyData.InitPane(4, viewpane::TextPane::CreateMultiLinePane(IDS_BODY, false));

		if (!m_lpFolder) return;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		if (!MyData.DisplayDialog()) return;

		const auto ulNumMessages = MyData.GetDecimal(0);
		auto szSubject = MyData.GetStringW(2);

		for (ULONG i = 0; i < ulNumMessages; i++)
		{
			const auto szTestSubject = strings::formatmessage(IDS_TESTSUBJECT, szSubject.c_str(), i);

			hRes = EC_H(mapi::SendTestMessage(
				lpMAPISession,
				m_lpFolder,
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

	void CFolderDlg::OnSetReadFlag()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpContentsTableListCtrl) return;

		output::DebugPrintEx(DBGGeneric, CLASS, L"OnSetReadFlag", L"\n");

		editor::CEditor MyFlags(
			this, IDS_SETREADFLAG, IDS_SETREADFLAGPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyFlags.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_FLAGSINHEX, false));
		MyFlags.SetHex(0, CLEAR_READ_FLAG);

		if (!MyFlags.DisplayDialog()) return;

		const int iNumSelected = m_lpContentsTableListCtrl->GetSelectedCount();

		if (iNumSelected == 1)
		{
			auto lpMAPIProp = m_lpContentsTableListCtrl->OpenNextSelectedItemProp(nullptr, mfcmapiREQUEST_MODIFY);
			auto lpMessage = mapi::safe_cast<LPMESSAGE>(lpMAPIProp);

			if (lpMessage)
			{
				EC_MAPI_S(lpMessage->SetReadFlag(MyFlags.GetHex(0)));
				lpMessage->Release();
			}

			if (lpMAPIProp) lpMAPIProp->Release();
		}
		else if (iNumSelected > 1)
		{
			const auto lpEIDs = m_lpContentsTableListCtrl->GetSelectedItemEIDs();

			LPMAPIPROGRESS lpProgress =
				mapi::mapiui::GetMAPIProgress(L"IMAPIFolder::SetReadFlags", m_hWnd); // STRING_OK

			auto ulFlags = MyFlags.GetHex(0);

			if (lpProgress) ulFlags |= MESSAGE_DIALOG;

			EC_MAPI_S(m_lpFolder->SetReadFlags(
				lpEIDs, lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL, lpProgress, ulFlags));

			if (lpProgress) lpProgress->Release();

			MAPIFreeBuffer(lpEIDs);
		}
	}

	void CFolderDlg::OnGetMessageOptions()
	{
		auto hRes = S_OK;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpContentsTableListCtrl) return;

		auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		output::DebugPrintEx(DBGGeneric, CLASS, L"OnGetMessageOptions", L"\n");

		editor::CEditor MyAddress(
			this, IDS_MESSAGEOPTIONS, IDS_ADDRESSTYPEPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyAddress.InitPane(
			0, viewpane::TextPane::CreateSingleLinePane(IDS_ADDRESSTYPE, std::wstring(L"EX"), false)); // STRING_OK
		if (!MyAddress.DisplayDialog()) return;

		auto iItem = m_lpContentsTableListCtrl->GetNextItem(-1, LVNI_SELECTED);
		while (iItem != -1)
		{
			auto lpMessage = OpenMessage(iItem, mfcmapiREQUEST_MODIFY);
			if (lpMessage)
			{
				hRes = EC_MAPI(lpMAPISession->MessageOptions(
					reinterpret_cast<ULONG_PTR>(m_hWnd),
					NULL, // API doesn't like Unicode
					LPTSTR(strings::wstringTostring(MyAddress.GetStringW(0)).c_str()),
					lpMessage));

				lpMessage->Release();
			}

			iItem = m_lpContentsTableListCtrl->GetNextItem(iItem, LVNI_SELECTED);
			if (S_OK != hRes && -1 != iItem)
			{
				if (bShouldCancel(this, hRes)) break;
			}
		}
	}

	// Read properties of the current message and save them for a restriction
	// Designed to work with messages, but could be made to work with anything
	void CFolderDlg::OnCreateMessageRestriction()
	{
		if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;

		// These are the properties we're going to copy off of the current message and store
		// in some object level variables
		enum
		{
			frPR_SUBJECT,
			frPR_CLIENT_SUBMIT_TIME,
			frPR_MESSAGE_DELIVERY_TIME,
			frNUMCOLS
		};
		static const SizedSPropTagArray(frNUMCOLS, sptFRCols) = {
			frNUMCOLS, {PR_SUBJECT, PR_CLIENT_SUBMIT_TIME, PR_MESSAGE_DELIVERY_TIME}};

		auto lpMAPIProp = m_lpContentsTableListCtrl->OpenNextSelectedItemProp(nullptr, mfcmapiREQUEST_MODIFY);

		if (lpMAPIProp)
		{
			ULONG cVals = 0;
			LPSPropValue lpProps = nullptr;
			EC_H_GETPROPS_S(lpMAPIProp->GetProps(LPSPropTagArray(&sptFRCols), fMapiUnicode, &cVals, &lpProps));
			if (lpProps)
			{
				// Allocate and create our SRestriction
				// Allocate base memory:
				auto lpRes = mapi::allocate<LPSRestriction>(sizeof(SRestriction));

				// Check that all our allocations were good before going on
				// Root Node
				lpRes->rt = RES_AND; // We're doing an AND...
				lpRes->res.resAnd.cRes = 2; // ...of two criteria...
				auto lpResLevel1 = mapi::allocate<LPSRestriction>(sizeof(SRestriction) * 2, lpRes);
				lpRes->res.resAnd.lpRes = lpResLevel1; // ...described here

				if (lpResLevel1)
				{
					lpResLevel1[0].rt = RES_PROPERTY;
					lpResLevel1[0].res.resProperty.relop = RELOP_EQ;
					lpResLevel1[0].res.resProperty.ulPropTag = PR_SUBJECT;
					auto lpspvSubject = mapi::allocate<LPSPropValue>(sizeof(SPropValue), lpRes);
					lpResLevel1[0].res.resProperty.lpProp = lpspvSubject;
					if (lpspvSubject)
					{
						// Allocate and fill out properties:
						lpspvSubject->ulPropTag = PR_SUBJECT;

						if (mapi::CheckStringProp(&lpProps[frPR_SUBJECT], PT_TSTRING))
						{
							lpspvSubject->Value.LPSZ = mapi::CopyString(lpProps[frPR_SUBJECT].Value.LPSZ, lpRes);
						}
						else
						{
							lpspvSubject->Value.LPSZ = nullptr;
						}
					}

					lpResLevel1[1].rt = RES_OR;
					lpResLevel1[1].res.resOr.cRes = 2;
					auto lpResLevel2 = mapi::allocate<LPSRestriction>(sizeof(SRestriction) * 2, lpRes);
					lpResLevel1[1].res.resOr.lpRes = lpResLevel2;

					if (lpResLevel2)
					{
						lpResLevel2[0].rt = RES_PROPERTY;
						lpResLevel2[0].res.resProperty.relop = RELOP_EQ;
						lpResLevel2[0].res.resProperty.ulPropTag = PR_CLIENT_SUBMIT_TIME;
						auto lpspvSubmitTime = mapi::allocate<LPSPropValue>(sizeof(SPropValue), lpRes);
						lpResLevel2[0].res.resProperty.lpProp = lpspvSubmitTime;
						if (lpspvSubmitTime)
						{
							lpspvSubmitTime->ulPropTag = PR_CLIENT_SUBMIT_TIME;
							if (PR_CLIENT_SUBMIT_TIME == lpProps[frPR_CLIENT_SUBMIT_TIME].ulPropTag)
							{
								lpspvSubmitTime->Value.ft.dwLowDateTime =
									lpProps[frPR_CLIENT_SUBMIT_TIME].Value.ft.dwLowDateTime;
								lpspvSubmitTime->Value.ft.dwHighDateTime =
									lpProps[frPR_CLIENT_SUBMIT_TIME].Value.ft.dwHighDateTime;
							}
							else
							{
								lpspvSubmitTime->Value.ft.dwLowDateTime = 0x0;
								lpspvSubmitTime->Value.ft.dwHighDateTime = 0x0;
							}
						}

						lpResLevel2[1].rt = RES_PROPERTY;
						lpResLevel2[1].res.resProperty.relop = RELOP_EQ;
						lpResLevel2[1].res.resProperty.ulPropTag = PR_MESSAGE_DELIVERY_TIME;
						auto lpspvDeliveryTime = mapi::allocate<LPSPropValue>(sizeof(SPropValue), lpRes);
						lpResLevel2[1].res.resProperty.lpProp = lpspvDeliveryTime;
						if (lpspvDeliveryTime)
						{
							lpspvDeliveryTime->ulPropTag = PR_MESSAGE_DELIVERY_TIME;
							if (PR_MESSAGE_DELIVERY_TIME == lpProps[frPR_MESSAGE_DELIVERY_TIME].ulPropTag)
							{
								lpspvDeliveryTime->Value.ft.dwLowDateTime =
									lpProps[frPR_MESSAGE_DELIVERY_TIME].Value.ft.dwLowDateTime;
								lpspvDeliveryTime->Value.ft.dwHighDateTime =
									lpProps[frPR_MESSAGE_DELIVERY_TIME].Value.ft.dwHighDateTime;
							}
							else
							{
								lpspvDeliveryTime->Value.ft.dwLowDateTime = 0x0;
								lpspvDeliveryTime->Value.ft.dwHighDateTime = 0x0;
							}
						}
					}
				}

				output::DebugPrintEx(DBGGeneric, CLASS, L"OnCreateMessageRestriction", L"built restriction:\n");
				output::DebugPrintRestriction(DBGGeneric, lpRes, lpMAPIProp);

				m_lpContentsTableListCtrl->SetRestriction(lpRes);

				SetRestrictionType(mfcmapiNORMAL_RESTRICTION);
				MAPIFreeBuffer(lpProps);
			}

			lpMAPIProp->Release();
		}
	}

	_Check_return_ HRESULT
	CFolderDlg::OnGetMessageStatus(int /*iItem*/, _In_ controls::sortlistdata::SortListData* lpData)
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!lpData || !lpData->Contents() || !m_lpFolder) return MAPI_E_INVALID_PARAMETER;

		output::DebugPrintEx(DBGGeneric, CLASS, L"OnGetMessageStatus", L"\n");

		ULONG ulMessageStatus = NULL;

		const auto lpMessageEID = lpData->Contents()->m_lpEntryID;

		if (lpMessageEID)
		{
			auto hRes = EC_MAPI(m_lpFolder->GetMessageStatus(
				lpMessageEID->cb, reinterpret_cast<LPENTRYID>(lpMessageEID->lpb), NULL, &ulMessageStatus));

			if (SUCCEEDED(hRes))
			{
				editor::CEditor MyStatus(this, IDS_MESSAGESTATUS, NULL, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
				MyStatus.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_MESSAGESTATUS, true));
				MyStatus.SetHex(0, ulMessageStatus);

				(void) MyStatus.DisplayDialog();
			}
		}

		return S_OK;
	}

	void CFolderDlg::OnSetMessageStatus()
	{
		auto hRes = S_OK;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpContentsTableListCtrl || !m_lpFolder) return;

		editor::CEditor MyData(
			this, IDS_SETMSGSTATUS, IDS_SETMSGSTATUSPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_STATUSINHEX, false));
		MyData.InitPane(1, viewpane::TextPane::CreateSingleLinePane(IDS_MASKINHEX, false));

		output::DebugPrintEx(DBGGeneric, CLASS, L"OnSetMessageStatus", L"\n");

		if (MyData.DisplayDialog())
		{
			auto iItem = m_lpContentsTableListCtrl->GetNextItem(-1, LVNI_SELECTED);
			while (iItem != -1)
			{
				hRes = S_OK;
				const auto lpListData = m_lpContentsTableListCtrl->GetSortListData(iItem);
				if (lpListData && lpListData->Contents())
				{
					const auto lpMessageEID = lpListData->Contents()->m_lpEntryID;

					if (lpMessageEID)
					{
						ULONG ulOldStatus = NULL;

						hRes = EC_MAPI(m_lpFolder->SetMessageStatus(
							lpMessageEID->cb,
							reinterpret_cast<LPENTRYID>(lpMessageEID->lpb),
							MyData.GetHex(0),
							MyData.GetHex(1),
							&ulOldStatus));
					}
				}

				iItem = m_lpContentsTableListCtrl->GetNextItem(iItem, LVNI_SELECTED);
				if (S_OK != hRes && -1 != iItem)
				{
					if (bShouldCancel(this, hRes)) break;
				}
			}
		}
	}

	_Check_return_ HRESULT CFolderDlg::OnSubmitMessage(int iItem, _In_ controls::sortlistdata::SortListData* /*lpData*/)
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		output::DebugPrintEx(DBGGeneric, CLASS, L"OnSubmitMesssage", L"\n");

		if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;

		auto hRes = S_OK;
		auto lpMessage = OpenMessage(iItem, mfcmapiREQUEST_MODIFY);
		if (lpMessage)
		{
			// Get subject line of message to copy.
			// This will be used as the new file name.
			hRes = EC_MAPI(lpMessage->SubmitMessage(NULL));

			lpMessage->Release();
		}

		return hRes;
	}

	_Check_return_ HRESULT CFolderDlg::OnAbortSubmit(int iItem, _In_ controls::sortlistdata::SortListData* lpData)
	{
		auto hRes = S_OK;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		output::DebugPrintEx(DBGGeneric, CLASS, L"OnSubmitMesssage", L"\n");

		if (-1 == iItem) return MAPI_E_INVALID_PARAMETER;
		if (!m_lpMapiObjects || !lpData || !lpData->Contents()) return MAPI_E_INVALID_PARAMETER;

		auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release

		const auto lpMessageEID = lpData->Contents()->m_lpEntryID;

		if (lpMDB && lpMessageEID)
		{
			hRes = EC_MAPI(lpMDB->AbortSubmit(lpMessageEID->cb, reinterpret_cast<LPENTRYID>(lpMessageEID->lpb), NULL));
		}

		return hRes;
	}

	void CFolderDlg::OnDisplayFolder(WORD wMenuSelect)
	{
		auto otType = otDefault;
		switch (wMenuSelect)
		{
		case ID_HIERARCHY:
			otType = otHierarchy;
			break;
		case ID_CONTENTS:
			otType = otContents;
			break;
		case ID_HIDDENCONTENTS:
			otType = otAssocContents;
			break;
		}

		EC_H_S(DisplayObject(m_lpFolder, NULL, otType, this));
	}

	void CFolderDlg::HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ LPMAPICONTAINER /*lpContainer*/)
	{
		if (lpParams)
		{
			lpParams->lpFolder = m_lpFolder;
			lpParams->lpMessage = mapi::safe_cast<LPMESSAGE>(lpMAPIProp);
			// Add appropriate flag to context
			if (m_ulDisplayFlags & dfAssoc) lpParams->ulCurrentFlags |= MENU_FLAGS_FOLDER_ASSOC;
			if (m_ulDisplayFlags & dfDeleted) lpParams->ulCurrentFlags |= MENU_FLAGS_DELETED;
		}

		addin::InvokeAddInMenu(lpParams);

		if (lpParams && lpParams->lpMessage)
		{
			lpParams->lpMessage->Release();
			lpParams->lpMessage = nullptr;
		}
	}

	LPMESSAGE CFolderDlg::OpenMessage(int iSelectedItem, __mfcmapiModifyEnum bModify)
	{
		auto lpMapiProp = OpenItemProp(iSelectedItem, bModify);
		auto lpMessage = mapi::safe_cast<LPMESSAGE>(lpMapiProp);
		if (lpMapiProp) lpMapiProp->Release();
		return lpMessage;
	}
}