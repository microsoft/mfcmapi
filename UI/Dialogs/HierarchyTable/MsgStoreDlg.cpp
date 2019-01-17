// Displays the hierarchy tree of folders in a message store
#include <StdAfx.h>
#include <UI/Dialogs/HierarchyTable/MsgStoreDlg.h>
#include <UI/Controls/HierarchyTableTreeCtrl.h>
#include <MAPI/Cache/MapiObjects.h>
#include <MAPI/MAPIFunctions.h>
#include <UI/Controls/SingleMAPIPropListCtrl.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <Interpret/InterpretProp.h>
#include <IO/File.h>
#include <MAPI/MAPIProgress.h>
#include <UI/Controls/SortList/NodeData.h>
#include <MAPI/Cache/GlobalCache.h>
#include <UI/Dialogs/ContentsTable/FormContainerDlg.h>
#include <UI/Dialogs/ContentsTable/FolderDlg.h>
#include <UI/mapiui.h>
#include <UI/addinui.h>
#include <IO/MFCOutput.h>
#include <IO/Registry.h>
#include <core/utility/strings.h>

namespace dialog
{
	static std::wstring CLASS = L"CMsgStoreDlg";

	CMsgStoreDlg::CMsgStoreDlg(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ cache::CMapiObjects* lpMapiObjects,
		_In_opt_ LPMAPIPROP lpMDB,
		_In_opt_ LPMAPIPROP lpRootFolder,
		ULONG ulDisplayFlags)
		: CHierarchyTableDlg(
			  pParentWnd,
			  lpMapiObjects,
			  IDS_FOLDERTREE,
			  lpRootFolder,
			  IDR_MENU_MESSAGESTORE_POPUP,
			  MENU_CONTEXT_FOLDER_TREE)
	{
		TRACE_CONSTRUCTOR(CLASS);

		m_lpMDB = mapi::safe_cast<LPMDB>(lpMDB);
		m_ulDisplayFlags = ulDisplayFlags;

		if (m_lpMapiObjects)
		{
			if (m_lpMDB)
			{
				m_szTitle = mapi::GetTitle(m_lpMDB);

				if (!GetRootContainer())
				{
					// Open root container.
					auto container = mapi::CallOpenEntry<LPUNKNOWN>(
						m_lpMDB,
						nullptr,
						nullptr,
						nullptr,
						nullptr, // open root container
						nullptr,
						MAPI_BEST_ACCESS,
						nullptr);

					SetRootContainer(container);
					container->Release();
				}
			}
		}

		CreateDialogAndMenu(IDR_MENU_MESSAGESTORE);
	}

	CMsgStoreDlg::~CMsgStoreDlg()
	{
		TRACE_DESTRUCTOR(CLASS);
		if (m_lpMDB) m_lpMDB->Release();
	}

	BEGIN_MESSAGE_MAP(CMsgStoreDlg, CHierarchyTableDlg)
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
	ON_COMMAND(ID_RESOLVEMESSAGECLASS, OnResolveMessageClass)
	ON_COMMAND(ID_SELECTFORM, OnSelectForm)
	ON_COMMAND(ID_RESENDALLMESSAGES, OnResendAllMessages)
	ON_COMMAND(ID_RESETPERMISSIONSONITEMS, OnResetPermissionsOnItems)
	ON_COMMAND(ID_RESTOREDELETEDFOLDER, OnRestoreDeletedFolder)
	ON_COMMAND(ID_SAVEFOLDERCONTENTSASMSG, OnSaveFolderContentsAsMSG)
	ON_COMMAND(ID_SAVEFOLDERCONTENTSASTEXTFILES, OnSaveFolderContentsAsTextFiles)
	ON_COMMAND(ID_SETRECEIVEFOLDER, OnSetReceiveFolder)
	ON_COMMAND(ID_EXPORTMESSAGES, OnExportMessages)

	ON_COMMAND(ID_VALIDATEIPMSUBTREE, OnValidateIPMSubtree)
	END_MESSAGE_MAP()

	void CMsgStoreDlg::OnInitMenu(_In_ CMenu* pMenu)
	{
		if (!pMenu) return;

		const auto bItemSelected = m_lpHierarchyTableTreeCtrl && m_lpHierarchyTableTreeCtrl.IsItemSelected();

		const auto ulStatus = cache::CGlobalCache::getInstance().GetBufferStatus();
		pMenu->EnableMenuItem(ID_PASTE, DIM((ulStatus != BUFFER_EMPTY) && bItemSelected));
		pMenu->EnableMenuItem(ID_PASTE_RULES, DIM((ulStatus & BUFFER_FOLDER) && bItemSelected));

		pMenu->EnableMenuItem(ID_DISPLAYASSOCIATEDCONTENTS, DIM(bItemSelected));
		pMenu->EnableMenuItem(ID_DISPLAYDELETEDCONTENTS, DIM(bItemSelected));
		pMenu->EnableMenuItem(ID_DISPLAYDELETEDSUBFOLDERS, DIM(bItemSelected));

		pMenu->EnableMenuItem(ID_CREATESUBFOLDER, DIM(bItemSelected));
		pMenu->EnableMenuItem(ID_DISPLAYACLTABLE, DIM(bItemSelected));
		pMenu->EnableMenuItem(ID_DISPLAYMAILBOXTABLE, DIM(m_lpMDB));
		pMenu->EnableMenuItem(ID_DISPLAYOUTGOINGQUEUE, DIM(m_lpMDB));
		pMenu->EnableMenuItem(ID_DISPLAYRECEIVEFOLDERTABLE, DIM(m_lpMDB));
		pMenu->EnableMenuItem(ID_DISPLAYRULESTABLE, DIM(bItemSelected));
		pMenu->EnableMenuItem(ID_EMPTYFOLDER, DIM(bItemSelected));
		pMenu->EnableMenuItem(ID_OPENFORMCONTAINER, DIM(bItemSelected));
		pMenu->EnableMenuItem(ID_RESOLVEMESSAGECLASS, DIM(bItemSelected));
		pMenu->EnableMenuItem(ID_SELECTFORM, DIM(bItemSelected));
		pMenu->EnableMenuItem(ID_SAVEFOLDERCONTENTSASMSG, DIM(bItemSelected));
		pMenu->EnableMenuItem(ID_SAVEFOLDERCONTENTSASTEXTFILES, DIM(bItemSelected));
		pMenu->EnableMenuItem(ID_SETRECEIVEFOLDER, DIM(bItemSelected));

		pMenu->EnableMenuItem(ID_RESENDALLMESSAGES, DIM(bItemSelected));
		pMenu->EnableMenuItem(ID_RESETPERMISSIONSONITEMS, DIM(bItemSelected));
		pMenu->EnableMenuItem(ID_RESTOREDELETEDFOLDER, DIM(bItemSelected && m_ulDisplayFlags & dfDeleted));

		pMenu->EnableMenuItem(ID_COPY, DIM(bItemSelected));
		pMenu->EnableMenuItem(ID_DELETESELECTEDITEM, DIM(bItemSelected));

		pMenu->EnableMenuItem(ID_DISPLAYINBOX, DIM(m_lpMDB));
		pMenu->EnableMenuItem(ID_DISPLAYCALENDAR, DIM(m_lpMDB));
		pMenu->EnableMenuItem(ID_DISPLAYCONTACTS, DIM(m_lpMDB));
		pMenu->EnableMenuItem(ID_DISPLAYTASKS, DIM(m_lpMDB));

		CHierarchyTableDlg::OnInitMenu(pMenu);
	}

	// Gets the selected folder with a single addref
	LPMAPIFOLDER CMsgStoreDlg::GetSelectedFolder(__mfcmapiModifyEnum bModify) const
	{
		if (!m_lpHierarchyTableTreeCtrl) return nullptr;
		auto container = m_lpHierarchyTableTreeCtrl.GetSelectedContainer(bModify);
		if (!container) return nullptr;

		const auto ret = mapi::safe_cast<LPMAPIFOLDER>(container);
		container->Release();
		return ret;
	}

	void CMsgStoreDlg::OnDisplaySpecialFolder(ULONG ulFolder)
	{
		if (!m_lpMDB) return;

		auto lpFolder = mapi::OpenDefaultFolder(ulFolder, m_lpMDB);
		if (lpFolder)
		{
			EC_H_S(DisplayObject(lpFolder, NULL, otHierarchy, this));

			lpFolder->Release();
		}
	}

	void CMsgStoreDlg::OnDisplayInbox() { OnDisplaySpecialFolder(mapi::DEFAULT_INBOX); }

	// See Q171670 INFO: Entry IDs of Outlook Special Folders for more info on these tags
	void CMsgStoreDlg::OnDisplayCalendarFolder() { OnDisplaySpecialFolder(mapi::DEFAULT_CALENDAR); }

	void CMsgStoreDlg::OnDisplayContactsFolder() { OnDisplaySpecialFolder(mapi::DEFAULT_CONTACTS); }

	void CMsgStoreDlg::OnDisplayTasksFolder() { OnDisplaySpecialFolder(mapi::DEFAULT_TASKS); }

	void CMsgStoreDlg::OnDisplayReceiveFolderTable()
	{
		LPMAPITABLE lpMAPITable = nullptr;

		if (!m_lpMDB) return;

		EC_MAPI_S(m_lpMDB->GetReceiveFolderTable(fMapiUnicode, &lpMAPITable));
		if (lpMAPITable)
		{
			EC_H_S(DisplayTable(lpMAPITable, otReceive, this));
			lpMAPITable->Release();
		}
	}

	void CMsgStoreDlg::OnDisplayOutgoingQueueTable()
	{
		LPMAPITABLE lpMAPITable = nullptr;

		if (!m_lpMDB) return;

		EC_MAPI_S(m_lpMDB->GetOutgoingQueue(NULL, &lpMAPITable));
		if (lpMAPITable)
		{
			EC_H_S(DisplayTable(lpMAPITable, otDefault, this));
			lpMAPITable->Release();
		}
	}

	void CMsgStoreDlg::OnDisplayRulesTable()
	{
		if (!m_lpHierarchyTableTreeCtrl) return;

		auto lpMAPIFolder = m_lpHierarchyTableTreeCtrl.GetSelectedContainer(mfcmapiREQUEST_MODIFY);

		if (lpMAPIFolder)
		{
			EC_H_S(DisplayExchangeTable(lpMAPIFolder, PR_RULES_TABLE, otRules, this));
			lpMAPIFolder->Release();
		}
	}

	void CMsgStoreDlg::OnResolveMessageClass()
	{
		if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl || !m_lpPropDisplay) return;

		auto lpMAPIFolder = GetSelectedFolder(mfcmapiREQUEST_MODIFY);

		if (lpMAPIFolder)
		{
			LPMAPIFORMINFO lpMAPIFormInfo = nullptr;
			ResolveMessageClass(m_lpMapiObjects, lpMAPIFolder, &lpMAPIFormInfo);
			if (lpMAPIFormInfo)
			{
				EC_H_S(m_lpPropDisplay->SetDataSource(lpMAPIFormInfo, NULL, false));
				lpMAPIFormInfo->Release();
			}

			lpMAPIFolder->Release();
		}
	}

	void CMsgStoreDlg::OnSelectForm()
	{
		LPMAPIFORMINFO lpMAPIFormInfo = nullptr;

		if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl || !m_lpPropDisplay) return;

		auto lpMAPIFolder = GetSelectedFolder(mfcmapiREQUEST_MODIFY);
		if (lpMAPIFolder)
		{
			SelectForm(m_hWnd, m_lpMapiObjects, lpMAPIFolder, &lpMAPIFormInfo);
			if (lpMAPIFormInfo)
			{
				EC_H_S(m_lpPropDisplay->SetDataSource(lpMAPIFormInfo, NULL, false));
				lpMAPIFormInfo->Release();
			}

			lpMAPIFolder->Release();
		}
	}

	void CMsgStoreDlg::OnOpenFormContainer()
	{
		LPMAPIFORMMGR lpMAPIFormMgr = nullptr;
		LPMAPIFORMCONTAINER lpMAPIFormContainer = nullptr;

		if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl || !m_lpPropDisplay) return;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		auto lpMAPIFolder = m_lpHierarchyTableTreeCtrl.GetSelectedContainer(mfcmapiREQUEST_MODIFY);

		if (lpMAPIFolder)
		{
			EC_MAPI_S(MAPIOpenFormMgr(lpMAPISession, &lpMAPIFormMgr));

			if (lpMAPIFormMgr)
			{
				EC_MAPI_S(lpMAPIFormMgr->OpenFormContainer(HFRMREG_FOLDER, lpMAPIFolder, &lpMAPIFormContainer));

				if (lpMAPIFormContainer)
				{
					new CFormContainerDlg(m_lpParent, m_lpMapiObjects, lpMAPIFormContainer);

					lpMAPIFormContainer->Release();
				}

				lpMAPIFormMgr->Release();
			}

			lpMAPIFolder->Release();
		}
	}

	// newstyle copy folder
	void CMsgStoreDlg::HandleCopy()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		output::DebugPrintEx(DBGGeneric, CLASS, L"OnCopyItems", L"\n");
		if (!m_lpHierarchyTableTreeCtrl) return;

		auto lpMAPISourceFolder = GetSelectedFolder(mfcmapiREQUEST_MODIFY);

		auto lpSrcParentFolder = mapi::GetParentFolder(lpMAPISourceFolder, m_lpMDB);

		cache::CGlobalCache::getInstance().SetFolderToCopy(lpMAPISourceFolder, lpSrcParentFolder);

		if (lpSrcParentFolder) lpSrcParentFolder->Release();
		if (lpMAPISourceFolder) lpMAPISourceFolder->Release();
	}

	_Check_return_ bool CMsgStoreDlg::HandlePaste()
	{
		if (CBaseDialog::HandlePaste()) return true;

		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		output::DebugPrintEx(DBGGeneric, CLASS, L"HandlePaste", L"\n");
		if (!m_lpHierarchyTableTreeCtrl) return false;

		const auto ulStatus = cache::CGlobalCache::getInstance().GetBufferStatus();

		// Get the destination Folder
		auto lpMAPIDestFolder = m_lpHierarchyTableTreeCtrl.GetSelectedContainer(mfcmapiREQUEST_MODIFY);

		if (lpMAPIDestFolder && ulStatus & BUFFER_MESSAGES && ulStatus & BUFFER_PARENTFOLDER)
		{
			OnPasteMessages();
		}
		else if (lpMAPIDestFolder && ulStatus & BUFFER_FOLDER)
		{
			editor::CEditor MyData(
				this, IDS_PASTEFOLDER, IDS_PASTEFOLDERPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

			MyData.AddPane(viewpane::CheckPane::Create(0, IDS_PASTEFOLDERCONTENTS, false, false));
			if (MyData.DisplayDialog())
			{
				const auto bPasteContents = MyData.GetCheck(0);
				if (bPasteContents)
					OnPasteFolderContents();
				else
					OnPasteFolder();
			}
		}

		if (lpMAPIDestFolder) lpMAPIDestFolder->Release();
		return true;
	}

	void CMsgStoreDlg::OnPasteMessages()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		output::DebugPrintEx(DBGGeneric, CLASS, L"OnPasteMessages", L"\n");
		if (!m_lpHierarchyTableTreeCtrl) return;

		// Get the source Messages
		const auto lpEIDs = cache::CGlobalCache::getInstance().GetMessagesToCopy();
		auto lpMAPISourceFolder = cache::CGlobalCache::getInstance().GetSourceParentFolder();
		// Get the destination Folder
		auto lpMAPIDestFolder = m_lpHierarchyTableTreeCtrl.GetSelectedContainer(mfcmapiREQUEST_MODIFY);

		if (lpMAPIDestFolder && lpMAPISourceFolder && lpEIDs)
		{
			editor::CEditor MyData(
				this, IDS_COPYMESSAGE, IDS_COPYMESSAGEPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

			MyData.AddPane(viewpane::CheckPane::Create(0, IDS_MESSAGEMOVE, false, false));
			if (MyData.DisplayDialog())
			{
				auto ulMoveMessage = MyData.GetCheck(0) ? MESSAGE_MOVE : 0;

				auto lpProgress =
					mapi::mapiui::GetMAPIProgress(L"IMAPIFolder::CopyMessages", m_hWnd); // STRING_OK

				if (lpProgress) ulMoveMessage |= MESSAGE_DIALOG;

				EC_MAPI_S(lpMAPISourceFolder->CopyMessages(
					lpEIDs,
					&IID_IMAPIFolder,
					lpMAPIDestFolder,
					lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL,
					lpProgress,
					ulMoveMessage));

				if (lpProgress) lpProgress->Release();
			}
		}

		if (lpMAPIDestFolder) lpMAPIDestFolder->Release();
		if (lpMAPISourceFolder) lpMAPISourceFolder->Release();
	}

	void CMsgStoreDlg::OnPasteFolder()
	{
		auto hRes = S_OK;
		ULONG cProps;
		LPSPropValue lpProps = nullptr;

		enum
		{
			NAME,
			EID,
			NUM_COLS
		};
		static const SizedSPropTagArray(NUM_COLS, sptaSrcFolder) = {NUM_COLS, {PR_DISPLAY_NAME_W, PR_ENTRYID}};

		output::DebugPrintEx(DBGGeneric, CLASS, L"OnPasteFolder", L"\n");

		// Get the source folder
		auto lpMAPISourceFolder = cache::CGlobalCache::getInstance().GetFolderToCopy();
		auto lpSrcParentFolder = cache::CGlobalCache::getInstance().GetSourceParentFolder();
		// Get the Destination Folder
		auto lpMAPIDestFolder = GetSelectedFolder(mfcmapiREQUEST_MODIFY);

		if (lpMAPISourceFolder && lpMAPIDestFolder)
		{
			output::DebugPrint(DBGGeneric, L"Folder Source Object = %p\n", lpMAPISourceFolder);
			output::DebugPrint(DBGGeneric, L"Folder Source Object Parent = %p\n", lpSrcParentFolder);
			output::DebugPrint(DBGGeneric, L"Folder Destination Object = %p\n", lpMAPIDestFolder);

			editor::CEditor MyData(
				this, IDS_PASTEFOLDER, IDS_PASTEFOLDERNEWNAMEPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_FOLDERNAME, false));
			MyData.AddPane(viewpane::CheckPane::Create(1, IDS_COPYSUBFOLDERS, false, false));
			MyData.AddPane(viewpane::CheckPane::Create(2, IDS_FOLDERMOVE, false, false));

			// Get required properties from the source folder
			EC_H_GETPROPS_S(
				lpMAPISourceFolder->GetProps(LPSPropTagArray(&sptaSrcFolder), fMapiUnicode, &cProps, &lpProps));

			if (lpProps)
			{
				if (mapi::CheckStringProp(&lpProps[NAME], PT_UNICODE))
				{
					output::DebugPrint(DBGGeneric, L"Folder Source Name = \"%ws\"\n", lpProps[NAME].Value.lpszW);
					MyData.SetStringW(0, lpProps[NAME].Value.lpszW);
				}
			}

			auto lpCopyRoot = lpSrcParentFolder;
			if (!lpSrcParentFolder) lpCopyRoot = lpMAPIDestFolder;

			if (MyData.DisplayDialog())
			{
				CWaitCursor Wait; // Change the mouse to an hourglass while we work.

				auto lpProgress =
					mapi::mapiui::GetMAPIProgress(L"IMAPIFolder::CopyFolder", m_hWnd); // STRING_OK

				auto ulCopyFlags = MAPI_UNICODE;
				if (MyData.GetCheck(1)) ulCopyFlags |= COPY_SUBFOLDERS;
				if (MyData.GetCheck(2)) ulCopyFlags |= FOLDER_MOVE;
				if (lpProgress) ulCopyFlags |= FOLDER_DIALOG;

				hRes = WC_MAPI(lpCopyRoot->CopyFolder(
					lpProps[EID].Value.bin.cb,
					reinterpret_cast<LPENTRYID>(lpProps[EID].Value.bin.lpb),
					&IID_IMAPIFolder,
					lpMAPIDestFolder,
					LPTSTR(MyData.GetStringW(0).c_str()),
					lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL,
					lpProgress, // Progress
					ulCopyFlags));
				if (hRes == MAPI_E_COLLISION)
				{
					error::ErrDialog(__FILE__, __LINE__, IDS_EDDUPEFOLDER);
				}
				else
					CHECKHRESMSG(hRes, IDS_COPYFOLDERFAILED);

				if (lpProgress) lpProgress->Release();
			}

			MAPIFreeBuffer(lpProps);
		}

		if (lpMAPIDestFolder) lpMAPIDestFolder->Release();
		if (lpMAPISourceFolder) lpMAPISourceFolder->Release();
		if (lpSrcParentFolder) lpSrcParentFolder->Release();
	}

	void CMsgStoreDlg::OnPasteFolderContents()
	{
		output::DebugPrintEx(DBGGeneric, CLASS, L"OnPasteFolderContents", L"\n");

		if (!m_lpHierarchyTableTreeCtrl) return;

		// Get the Source Folder
		auto lpMAPISourceFolder = cache::CGlobalCache::getInstance().GetFolderToCopy();
		// Get the Destination Folder
		auto lpMAPIDestFolder = GetSelectedFolder(mfcmapiREQUEST_MODIFY);

		if (lpMAPISourceFolder && lpMAPIDestFolder)
		{
			output::DebugPrint(DBGGeneric, L"Folder Source Object = %p\n", lpMAPISourceFolder);
			output::DebugPrint(DBGGeneric, L"Folder Destination Object = %p\n", lpMAPIDestFolder);

			editor::CEditor MyData(
				this, IDS_COPYFOLDERCONTENTS, IDS_PICKOPTIONSPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyData.AddPane(viewpane::CheckPane::Create(0, IDS_COPYASSOCIATEDITEMS, false, false));
			MyData.AddPane(viewpane::CheckPane::Create(1, IDS_MOVEMESSAGES, false, false));
			MyData.AddPane(viewpane::CheckPane::Create(2, IDS_SINGLECALLCOPY, true, false));
			if (MyData.DisplayDialog())
			{
				CWaitCursor Wait; // Change the mouse to an hourglass while we work.

				mapi::CopyFolderContents(
					lpMAPISourceFolder,
					lpMAPIDestFolder,
					MyData.GetCheck(0), // associated contents
					MyData.GetCheck(1), // move
					MyData.GetCheck(2), // Single CopyMessages call
					m_hWnd);
			}
		}

		if (lpMAPIDestFolder) lpMAPIDestFolder->Release();
		if (lpMAPISourceFolder) lpMAPISourceFolder->Release();
	}

	void CMsgStoreDlg::OnPasteRules()
	{
		output::DebugPrintEx(DBGGeneric, CLASS, L"OnPasteRules", L"\n");

		if (!m_lpHierarchyTableTreeCtrl) return;

		// Get the Source Folder
		auto lpMAPISourceFolder = cache::CGlobalCache::getInstance().GetFolderToCopy();
		// Get the Destination Folder
		auto lpMAPIDestFolder = GetSelectedFolder(mfcmapiREQUEST_MODIFY);

		if (lpMAPISourceFolder && lpMAPIDestFolder)
		{
			output::DebugPrint(DBGGeneric, L"Folder Source Object = %p\n", lpMAPISourceFolder);
			output::DebugPrint(DBGGeneric, L"Folder Destination Object = %p\n", lpMAPIDestFolder);

			editor::CEditor MyData(
				this, IDS_COPYFOLDERRULES, IDS_COPYFOLDERRULESPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyData.AddPane(viewpane::CheckPane::Create(0, IDS_REPLACERULES, false, false));
			if (MyData.DisplayDialog())
			{
				CWaitCursor Wait; // Change the mouse to an hourglass while we work.

				EC_H_S(mapi::CopyFolderRules(lpMAPISourceFolder, lpMAPIDestFolder,
											 MyData.GetCheck(0))); // move
			}
		}

		if (lpMAPIDestFolder) lpMAPIDestFolder->Release();
		if (lpMAPISourceFolder) lpMAPISourceFolder->Release();
	}

	void CMsgStoreDlg::OnCreateSubFolder()
	{
		LPMAPIFOLDER lpMAPISubFolder = nullptr;

		editor::CEditor MyData(
			this, IDS_ADDSUBFOLDER, IDS_ADDSUBFOLDERPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.SetPromptPostFix(interpretprop::AllFlagsToString(PROP_ID(PR_FOLDER_TYPE), true));
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePaneID(0, IDS_FOLDERNAME, IDS_FOLDERNAMEVALUE, false));
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_FOLDERTYPE, false));
		MyData.SetHex(1, FOLDER_GENERIC);
		auto szProduct = strings::loadstring(ID_PRODUCTNAME);
		const auto szFolderComment = strings::formatmessage(IDS_FOLDERCOMMENTVALUE, szProduct.c_str());
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_FOLDERCOMMENT, szFolderComment, false));
		MyData.AddPane(viewpane::CheckPane::Create(3, IDS_PASSOPENIFEXISTS, false, false));

		auto lpMAPIFolder = GetSelectedFolder(mfcmapiREQUEST_MODIFY);

		if (lpMAPIFolder)
		{
			if (MyData.DisplayDialog())
			{
				EC_MAPI_S(lpMAPIFolder->CreateFolder(
					MyData.GetHex(1),
					LPTSTR(MyData.GetStringW(0).c_str()),
					LPTSTR(MyData.GetStringW(2).c_str()),
					NULL, // interface
					MAPI_UNICODE | (MyData.GetCheck(3) ? OPEN_IF_EXISTS : 0),
					&lpMAPISubFolder));

				if (lpMAPISubFolder) lpMAPISubFolder->Release();
			}

			lpMAPIFolder->Release();
		}
	}

	void CMsgStoreDlg::OnDisplayACLTable()
	{
		if (!m_lpHierarchyTableTreeCtrl) return;

		auto lpMAPIFolder = m_lpHierarchyTableTreeCtrl.GetSelectedContainer(mfcmapiREQUEST_MODIFY);

		if (lpMAPIFolder)
		{
			EC_H_S(DisplayExchangeTable(lpMAPIFolder, PR_ACL_TABLE, otACL, this));
			lpMAPIFolder->Release();
		}
	}

	void CMsgStoreDlg::OnDisplayAssociatedContents()
	{
		if (!m_lpHierarchyTableTreeCtrl) return;

		// Find the highlighted item
		auto lpMAPIFolder = m_lpHierarchyTableTreeCtrl.GetSelectedContainer(mfcmapiREQUEST_MODIFY);

		if (lpMAPIFolder)
		{
			(void) DisplayObject(lpMAPIFolder, NULL, otAssocContents, this);

			lpMAPIFolder->Release();
		}
	}

	void CMsgStoreDlg::OnDisplayDeletedContents()
	{
		if (!m_lpHierarchyTableTreeCtrl || !m_lpMapiObjects) return;

		// Find the highlighted item
		const auto lpItemEID = m_lpHierarchyTableTreeCtrl.GetSelectedItemEID();

		if (lpItemEID)
		{
			if (m_lpMDB)
			{
				auto lpMAPIFolder = mapi::CallOpenEntry<LPMAPIFOLDER>(
					m_lpMDB,
					nullptr,
					nullptr,
					nullptr,
					lpItemEID->cb,
					reinterpret_cast<LPENTRYID>(lpItemEID->lpb),
					nullptr,
					MAPI_BEST_ACCESS | SHOW_SOFT_DELETES | MAPI_NO_CACHE,
					nullptr);
				if (lpMAPIFolder)
				{
					// call the dialog
					new CFolderDlg(m_lpParent, m_lpMapiObjects, lpMAPIFolder, dfDeleted);

					lpMAPIFolder->Release();
				}
			}
		}
	}

	void CMsgStoreDlg::OnDisplayDeletedSubFolders()
	{
		if (!m_lpHierarchyTableTreeCtrl) return;

		// Must open the folder with MODIFY permissions if I'm going to restore the folder!
		auto lpFolder = m_lpHierarchyTableTreeCtrl.GetSelectedContainer(mfcmapiREQUEST_MODIFY);

		if (lpFolder)
		{
			new CMsgStoreDlg(m_lpParent, m_lpMapiObjects, m_lpMDB, lpFolder, dfDeleted);
			lpFolder->Release();
		}
	}

	void CMsgStoreDlg::OnDisplayMailboxTable()
	{
		if (!m_lpParent || !m_lpMapiObjects) return;

		DisplayMailboxTable(m_lpParent, m_lpMapiObjects);
	}

	void CMsgStoreDlg::OnEmptyFolder()
	{
		if (!m_lpHierarchyTableTreeCtrl) return;

		// Find the highlighted item
		auto lpMAPIFolderToEmpty = GetSelectedFolder(mfcmapiREQUEST_MODIFY);

		if (lpMAPIFolderToEmpty)
		{
			editor::CEditor MyData(
				this, IDS_DELETEITEMSANDSUB, IDS_DELETEITEMSANDSUBPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyData.AddPane(viewpane::CheckPane::Create(0, IDS_DELASSOCIATED, false, false));
			MyData.AddPane(viewpane::CheckPane::Create(1, IDS_HARDDELETION, false, false));
			MyData.AddPane(viewpane::CheckPane::Create(2, IDS_MANUALLYEMPTYFOLDER, false, false));

			if (MyData.DisplayDialog())
			{
				if (MyData.GetCheck(2))
				{
					EC_H_S(mapi::ManuallyEmptyFolder(lpMAPIFolderToEmpty, MyData.GetCheck(0), MyData.GetCheck(1)));
				}
				else
				{
					auto ulFlags = MyData.GetCheck(0) ? DEL_ASSOCIATED : 0;
					ulFlags |= MyData.GetCheck(1) ? DELETE_HARD_DELETE : 0;
					auto lpProgress =
						mapi::mapiui::GetMAPIProgress(L"IMAPIFolder::EmptyFolder", m_hWnd); // STRING_OK

					if (lpProgress) ulFlags |= FOLDER_DIALOG;

					output::DebugPrintEx(
						DBGGeneric,
						CLASS,
						L"OnEmptyFolder",
						L"Calling EmptyFolder on %p, ulFlags = 0x%08X.\n",
						lpMAPIFolderToEmpty,
						ulFlags);

					EC_MAPI_S(lpMAPIFolderToEmpty->EmptyFolder(
						lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL, lpProgress, ulFlags));

					if (lpProgress) lpProgress->Release();
				}
			}

			lpMAPIFolderToEmpty->Release();
		}
	}

	void CMsgStoreDlg::OnDeleteSelectedItem()
	{
		auto hRes = S_OK;
		LPSBinary lpItemEID = nullptr;

		if (!m_lpHierarchyTableTreeCtrl) return;

		const auto hItem = m_lpHierarchyTableTreeCtrl.GetSelectedItem();
		if (hItem)
		{
			const auto lpData = m_lpHierarchyTableTreeCtrl.GetSortListData(hItem);
			if (lpData && lpData->Node()) lpItemEID = lpData->Node()->m_lpEntryID;
		}

		if (!lpItemEID) return;
		if (!m_lpMDB) return;

		auto lpFolderToDelete = GetSelectedFolder(mfcmapiDO_NOT_REQUEST_MODIFY);
		if (lpFolderToDelete)
		{
			auto lpParentFolder = mapi::GetParentFolder(lpFolderToDelete, m_lpMDB);
			if (lpParentFolder)
			{
				const ULONG bShiftPressed = GetKeyState(VK_SHIFT) < 0;
				auto bDelete = true;
				auto bHardDelete = bShiftPressed;
				if (!bShiftPressed)
				{
					editor::CEditor MyData(
						this, IDS_DELETEFOLDER, IDS_DELETEFOLDERPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
					MyData.AddPane(viewpane::CheckPane::Create(0, IDS_HARDDELETION, false, false));

					if (MyData.DisplayDialog())
					{
						bHardDelete = MyData.GetCheck(0);
					}
					else
					{
						bDelete = false;
					}
				}

				if (bDelete)
				{
					auto ulFlags = DEL_FOLDERS | DEL_MESSAGES | (bHardDelete ? DELETE_HARD_DELETE : 0);

					output::DebugPrintEx(
						DBGDeleteSelectedItem,
						CLASS,
						L"OnDeleteSelectedItem",
						L"Calling DeleteFolder on folder. ulFlags = 0x%08X.\n",
						ulFlags);
					output::DebugPrintBinary(DBGGeneric, *lpItemEID);

					auto lpProgress =
						mapi::mapiui::GetMAPIProgress(L"IMAPIFolder::DeleteFolder", m_hWnd); // STRING_OK

					if (lpProgress) ulFlags |= FOLDER_DIALOG;

					hRes = EC_MAPI(lpParentFolder->DeleteFolder(
						lpItemEID->cb,
						reinterpret_cast<LPENTRYID>(lpItemEID->lpb),
						lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL,
						lpProgress,
						ulFlags));

					if (SUCCEEDED(hRes))
					{
						// Delete the item from the UI since we cannot rely on notifications to handle this for us
						WC_B_S(m_lpHierarchyTableTreeCtrl.DeleteItem(hItem));
					}

					if (lpProgress) lpProgress->Release();
				}

				lpParentFolder->Release();
			}

			lpFolderToDelete->Release();
		}
	}

	void CMsgStoreDlg::OnSaveFolderContentsAsMSG()
	{
		if (!m_lpHierarchyTableTreeCtrl) return;

		output::DebugPrintEx(DBGGeneric, CLASS, L"OnSaveFolderContentsAsMSG", L"\n");

		// Find the highlighted item
		auto lpMAPIFolder = GetSelectedFolder(mfcmapiDO_NOT_REQUEST_MODIFY);
		if (!lpMAPIFolder) return;

		editor::CEditor MyData(
			this, IDS_SAVEFOLDERASMSG, IDS_SAVEFOLDERASMSGPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.AddPane(viewpane::CheckPane::Create(0, IDS_SAVEASSOCIATEDCONTENTS, false, false));
		MyData.AddPane(viewpane::CheckPane::Create(1, IDS_SAVEUNICODE, false, false));
		if (MyData.DisplayDialog())
		{
			auto szDir = file::GetDirectoryPath(m_hWnd);
			if (!szDir.empty())
			{
				CWaitCursor Wait; // Change the mouse to an hourglass while we work.

				EC_H_S(
					file::SaveFolderContentsToMSG(lpMAPIFolder, szDir, MyData.GetCheck(0), MyData.GetCheck(1), m_hWnd));
			}
		}

		lpMAPIFolder->Release();
	}

	void CMsgStoreDlg::OnSaveFolderContentsAsTextFiles()
	{
		if (!m_lpMDB || !m_lpHierarchyTableTreeCtrl) return;

		auto lpFolder = GetSelectedFolder(mfcmapiDO_NOT_REQUEST_MODIFY);

		if (lpFolder)
		{
			editor::CEditor MyData(
				this, IDS_SAVEFOLDERASPROPFILES, IDS_PICKOPTIONSPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyData.AddPane(viewpane::CheckPane::Create(0, IDS_RECURSESUBFOLDERS, false, false));
			MyData.AddPane(viewpane::CheckPane::Create(1, IDS_SAVEREGULARCONTENTS, true, false));
			MyData.AddPane(viewpane::CheckPane::Create(2, IDS_SAVEASSOCIATEDCONTENTS, true, false));

			if (MyData.DisplayDialog())
			{
				file::SaveFolderContentsToTXT(
					m_lpMDB, lpFolder, MyData.GetCheck(1), MyData.GetCheck(2), MyData.GetCheck(0), m_hWnd);
			}

			lpFolder->Release();
		}
	}

	void CMsgStoreDlg::OnExportMessages()
	{
		if (!m_lpHierarchyTableTreeCtrl) return;

		auto lpFolder = GetSelectedFolder(mfcmapiDO_NOT_REQUEST_MODIFY);

		if (lpFolder)
		{
			ui::mapiui::ExportMessages(lpFolder, m_hWnd);

			lpFolder->Release();
		}
	}

	void CMsgStoreDlg::OnSetReceiveFolder()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpMDB || !m_lpHierarchyTableTreeCtrl) return;

		editor::CEditor MyData(
			this, IDS_SETRECFOLDER, IDS_SETRECFOLDERPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_CLASS, false));
		MyData.AddPane(viewpane::CheckPane::Create(1, IDS_DELETEASSOCIATION, false, false));

		// Find the highlighted item
		const auto lpEID = m_lpHierarchyTableTreeCtrl.GetSelectedItemEID();

		if (MyData.DisplayDialog())
		{
			if (MyData.GetCheck(1))
			{
				EC_MAPI_S(m_lpMDB->SetReceiveFolder(LPTSTR(MyData.GetStringW(0).c_str()), MAPI_UNICODE, NULL, NULL));
			}
			else if (lpEID)
			{
				EC_MAPI_S(m_lpMDB->SetReceiveFolder(
					LPTSTR(MyData.GetStringW(0).c_str()),
					MAPI_UNICODE,
					lpEID->cb,
					reinterpret_cast<LPENTRYID>(lpEID->lpb)));
			}
		}
	}

	void CMsgStoreDlg::OnResendAllMessages()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpHierarchyTableTreeCtrl) return;

		// Find the highlighted item
		auto lpMAPIFolder = GetSelectedFolder(mfcmapiREQUEST_MODIFY);

		if (lpMAPIFolder)
		{
			EC_H_S(mapi::ResendMessages(lpMAPIFolder, m_hWnd));

			lpMAPIFolder->Release();
		}
	}

	// Iterate through items in the selected folder and attempt to delete PR_NT_SECURITY_DESCRIPTOR
	void CMsgStoreDlg::OnResetPermissionsOnItems()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpMDB || !m_lpHierarchyTableTreeCtrl) return;

		// Find the highlighted item
		auto lpMAPIFolder = GetSelectedFolder(mfcmapiREQUEST_MODIFY);

		if (lpMAPIFolder)
		{
			EC_H_S(mapi::ResetPermissionsOnItems(m_lpMDB, lpMAPIFolder));
			lpMAPIFolder->Release();
		}
	}

	// Copy selected folder back to the land of the living
	void CMsgStoreDlg::OnRestoreDeletedFolder()
	{
		ULONG cProps;
		LPSPropValue lpProps = nullptr;

		if (!m_lpMDB || !m_lpHierarchyTableTreeCtrl) return;

		enum
		{
			NAME,
			EID,
			NUM_COLS
		};
		static const SizedSPropTagArray(NUM_COLS, sptaSrcFolder) = {NUM_COLS, {PR_DISPLAY_NAME_W, PR_ENTRYID}};

		auto lpSrcFolder = GetSelectedFolder(mfcmapiREQUEST_MODIFY);

		if (lpSrcFolder)
		{
			auto lpSrcParentFolder = mapi::GetParentFolder(lpSrcFolder, m_lpMDB);

			// Get required properties from the source folder
			if (lpSrcParentFolder)
			{
				EC_H_GETPROPS_S(
					lpSrcFolder->GetProps(LPSPropTagArray(&sptaSrcFolder), fMapiUnicode, &cProps, &lpProps));
			}

			editor::CEditor MyData(
				this, IDS_RESTOREDELFOLD, IDS_RESTOREDELFOLDPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_FOLDERNAME, false));
			MyData.AddPane(viewpane::CheckPane::Create(1, IDS_COPYSUBFOLDERS, false, false));

			if (lpProps)
			{
				if (mapi::CheckStringProp(&lpProps[NAME], PT_UNICODE))
				{
					output::DebugPrint(DBGGeneric, L"Folder Source Name = \"%ws\"\n", lpProps[NAME].Value.lpszW);
					MyData.SetStringW(0, lpProps[NAME].Value.lpszW);
				}
			}

			if (MyData.DisplayDialog())
			{
				// Restore the folder up under m_lpContainer
				CWaitCursor Wait; // Change the mouse to an hourglass while we work.

				output::DebugPrint(DBGGeneric, L"Restoring %p to %p as \n", lpSrcFolder, GetRootContainer());

				if (!lpSrcParentFolder) lpSrcParentFolder = mapi::safe_cast<LPMAPIFOLDER>(GetRootContainer());

				auto lpProgress =
					mapi::mapiui::GetMAPIProgress(L"IMAPIFolder::CopyFolder", m_hWnd); // STRING_OK

				auto ulCopyFlags = MAPI_UNICODE | (MyData.GetCheck(1) ? COPY_SUBFOLDERS : 0);

				if (lpProgress) ulCopyFlags |= FOLDER_DIALOG;

				const auto hRes = WC_MAPI(lpSrcParentFolder->CopyFolder(
					lpProps[EID].Value.bin.cb,
					reinterpret_cast<LPENTRYID>(lpProps[EID].Value.bin.lpb),
					&IID_IMAPIFolder,
					GetRootContainer(),
					LPTSTR(MyData.GetStringW(0).c_str()),
					lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL,
					lpProgress,
					ulCopyFlags));
				if (hRes == MAPI_E_COLLISION)
				{
					error::ErrDialog(__FILE__, __LINE__, IDS_EDDUPEFOLDER);
				}
				else if (hRes == MAPI_W_PARTIAL_COMPLETION)
				{
					error::ErrDialog(__FILE__, __LINE__, IDS_EDRESTOREFAILED);
				}
				else
					CHECKHRESMSG(hRes, IDS_COPYFOLDERFAILED);

				if (lpProgress) lpProgress->Release();
			}

			MAPIFreeBuffer(lpProps);
			if (lpSrcParentFolder) lpSrcParentFolder->Release();
			lpSrcFolder->Release();
		}
	}

	void CMsgStoreDlg::OnValidateIPMSubtree()
	{
		if (!m_lpMDB) return;

		ULONG ulValues = 0;
		LPSPropValue lpProps = nullptr;
		LPMAPIERROR lpErr = nullptr;

		editor::CEditor MyData(
			this, IDS_VALIDATEIPMSUB, IDS_PICKOPTIONSPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.AddPane(viewpane::CheckPane::Create(0, IDS_MAPIFORCECREATE, false, false));
		MyData.AddPane(viewpane::CheckPane::Create(1, IDS_MAPIFULLIPMTREE, false, false));

		if (MyData.DisplayDialog())
		{
			const auto ulFlags =
				(MyData.GetCheck(0) ? MAPI_FORCE_CREATE : 0) | (MyData.GetCheck(1) ? MAPI_FULL_IPM_TREE : 0);

			output::DebugPrintEx(DBGGeneric, CLASS, L"OnValidateIPMSubtree", L"ulFlags = 0x%08X\n", ulFlags);

			EC_MAPI_S(HrValidateIPMSubtree(m_lpMDB, ulFlags, &ulValues, &lpProps, &lpErr));
			EC_MAPIERR(fMapiUnicode, lpErr);
			MAPIFreeBuffer(lpErr);

			if (ulValues > 0 && lpProps)
			{
				output::DebugPrintEx(
					DBGGeneric,
					CLASS,
					L"OnValidateIPMSubtree",
					L"HrValidateIPMSubtree returned 0x%08X properties:\n",
					ulValues);
				output::DebugPrintProperties(DBGGeneric, ulValues, lpProps, m_lpMDB);
			}

			MAPIFreeBuffer(lpProps);
		}
	}

	void CMsgStoreDlg::HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_ LPMAPIPROP /*lpMAPIProp*/,
		_In_ LPMAPICONTAINER lpContainer)
	{
		if (lpParams)
		{
			lpParams->lpFolder = mapi::safe_cast<LPMAPIFOLDER>(lpContainer);
		}

		ui::addinui::InvokeAddInMenu(lpParams);

		if (lpParams && lpParams->lpAbCont)
		{
			lpParams->lpAbCont->Release();
			lpParams->lpAbCont = nullptr;
		}
	}
} // namespace dialog