// Displays the hierarchy tree of folders in a message store
#include "StdAfx.h"
#include <UI/Dialogs/HierarchyTable/MsgStoreDlg.h>
#include <UI/Controls/HierarchyTableTreeCtrl.h>
#include <MAPI/MapiObjects.h>
#include <MAPI/MAPIFunctions.h>
#include <UI/Controls/SingleMAPIPropListCtrl.h>
#include <UI/MFCUtilityFunctions.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <Interpret/InterpretProp.h>
#include <IO/File.h>
#include <MAPI/MAPIProgress.h>
#include <UI/Controls/SortList/NodeData.h>
#include <MAPI/GlobalCache.h>
#include <UI/Dialogs/ContentsTable/FormContainerDlg.h>
#include <UI/Dialogs/ContentsTable/FolderDlg.h>

static std::wstring CLASS = L"CMsgStoreDlg";

CMsgStoreDlg::CMsgStoreDlg(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects,
	_In_opt_ LPMAPIFOLDER lpRootFolder,
	ULONG ulDisplayFlags
) :
	CHierarchyTableDlg(
		pParentWnd,
		lpMapiObjects,
		IDS_FOLDERTREE,
		lpRootFolder,
		IDR_MENU_MESSAGESTORE_POPUP,
		MENU_CONTEXT_FOLDER_TREE)
{
	TRACE_CONSTRUCTOR(CLASS);
	auto hRes = S_OK;

	m_ulDisplayFlags = ulDisplayFlags;

	if (m_lpMapiObjects)
	{
		const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
		if (lpMDB)
		{
			m_szTitle = GetTitle(lpMDB);

			if (!m_lpContainer)
			{
				// Open root container.
				EC_H(CallOpenEntry(
					lpMDB,
					NULL,
					NULL,
					NULL,
					NULL, // open root container
					NULL,
					MAPI_BEST_ACCESS,
					NULL,
					reinterpret_cast<LPUNKNOWN*>(&m_lpContainer)));
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

	LPMDB lpMDB = nullptr;
	const auto bItemSelected = m_lpHierarchyTableTreeCtrl && m_lpHierarchyTableTreeCtrl->IsItemSelected();

	if (m_lpMapiObjects)
	{
		lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	}

	const auto ulStatus = CGlobalCache::getInstance().GetBufferStatus();
	pMenu->EnableMenuItem(ID_PASTE, DIM((ulStatus != BUFFER_EMPTY) && bItemSelected));
	pMenu->EnableMenuItem(ID_PASTE_RULES, DIM((ulStatus & BUFFER_FOLDER) && bItemSelected));

	pMenu->EnableMenuItem(ID_DISPLAYASSOCIATEDCONTENTS, DIM(bItemSelected));
	pMenu->EnableMenuItem(ID_DISPLAYDELETEDCONTENTS, DIM(bItemSelected));
	pMenu->EnableMenuItem(ID_DISPLAYDELETEDSUBFOLDERS, DIM(bItemSelected));

	pMenu->EnableMenuItem(ID_CREATESUBFOLDER, DIM(bItemSelected));
	pMenu->EnableMenuItem(ID_DISPLAYACLTABLE, DIM(bItemSelected));
	pMenu->EnableMenuItem(ID_DISPLAYMAILBOXTABLE, DIM(lpMDB));
	pMenu->EnableMenuItem(ID_DISPLAYOUTGOINGQUEUE, DIM(lpMDB));
	pMenu->EnableMenuItem(ID_DISPLAYRECEIVEFOLDERTABLE, DIM(lpMDB));
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

	pMenu->EnableMenuItem(ID_DISPLAYINBOX, DIM(lpMDB));
	pMenu->EnableMenuItem(ID_DISPLAYCALENDAR, DIM(lpMDB));
	pMenu->EnableMenuItem(ID_DISPLAYCONTACTS, DIM(lpMDB));
	pMenu->EnableMenuItem(ID_DISPLAYTASKS, DIM(lpMDB));

	CHierarchyTableDlg::OnInitMenu(pMenu);
}

void CMsgStoreDlg::OnDisplaySpecialFolder(ULONG ulFolder)
{
	auto hRes = S_OK;
	LPMAPIFOLDER lpFolder = nullptr;

	if (!m_lpMapiObjects) return;

	const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return;

	EC_H(OpenDefaultFolder(ulFolder, lpMDB, &lpFolder));

	if (lpFolder)
	{
		EC_H(DisplayObject(
			lpFolder,
			NULL,
			otHierarchy,
			this));

		lpFolder->Release();
	}
}

void CMsgStoreDlg::OnDisplayInbox()
{
	OnDisplaySpecialFolder(DEFAULT_INBOX);
}

// See Q171670 INFO: Entry IDs of Outlook Special Folders for more info on these tags
void CMsgStoreDlg::OnDisplayCalendarFolder()
{
	OnDisplaySpecialFolder(DEFAULT_CALENDAR);
}

void CMsgStoreDlg::OnDisplayContactsFolder()
{
	OnDisplaySpecialFolder(DEFAULT_CONTACTS);
}

void CMsgStoreDlg::OnDisplayTasksFolder()
{
	OnDisplaySpecialFolder(DEFAULT_TASKS);
}

void CMsgStoreDlg::OnDisplayReceiveFolderTable()
{
	auto hRes = S_OK;
	LPMAPITABLE lpMAPITable = nullptr;

	if (!m_lpMapiObjects) return;

	auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return;

	EC_MAPI(lpMDB->GetReceiveFolderTable(
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
}

void CMsgStoreDlg::OnDisplayOutgoingQueueTable()
{
	auto hRes = S_OK;
	LPMAPITABLE lpMAPITable = nullptr;

	if (!m_lpMapiObjects) return;

	auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return;

	EC_MAPI(lpMDB->GetOutgoingQueue(
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
}

void CMsgStoreDlg::OnDisplayRulesTable()
{
	auto hRes = S_OK;

	if (!m_lpHierarchyTableTreeCtrl) return;

	auto lpMAPIFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY));

	if (lpMAPIFolder)
	{
		EC_H(DisplayExchangeTable(
			lpMAPIFolder,
			PR_RULES_TABLE,
			otRules,
			this));
		lpMAPIFolder->Release();
	}
}

void CMsgStoreDlg::OnResolveMessageClass()
{
	auto hRes = S_OK;
	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl || !m_lpPropDisplay) return;

	auto lpMAPIFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY));

	if (lpMAPIFolder)
	{
		LPMAPIFORMINFO lpMAPIFormInfo = nullptr;
		ResolveMessageClass(m_lpMapiObjects, lpMAPIFolder, &lpMAPIFormInfo);
		if (lpMAPIFormInfo)
		{
			EC_H(m_lpPropDisplay->SetDataSource(lpMAPIFormInfo, NULL, false));
			lpMAPIFormInfo->Release();
		}
		lpMAPIFolder->Release();
	}
}

void CMsgStoreDlg::OnSelectForm()
{
	auto hRes = S_OK;
	LPMAPIFORMINFO lpMAPIFormInfo = nullptr;

	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl || !m_lpPropDisplay) return;

	auto lpMAPIFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY));
	if (lpMAPIFolder)
	{
		SelectForm(m_hWnd, m_lpMapiObjects, lpMAPIFolder, &lpMAPIFormInfo);
		if (lpMAPIFormInfo)
		{
			EC_H(m_lpPropDisplay->SetDataSource(lpMAPIFormInfo, NULL, false));
			lpMAPIFormInfo->Release();
		}
		lpMAPIFolder->Release();
	}
}

void CMsgStoreDlg::OnOpenFormContainer()
{
	auto hRes = S_OK;
	LPMAPIFORMMGR lpMAPIFormMgr = nullptr;
	LPMAPIFORMCONTAINER lpMAPIFormContainer = nullptr;

	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl || !m_lpPropDisplay) return;

	const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	auto lpMAPIFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY));

	if (lpMAPIFolder)
	{
		EC_MAPI(MAPIOpenFormMgr(lpMAPISession, &lpMAPIFormMgr));

		if (lpMAPIFormMgr)
		{
			EC_MAPI(lpMAPIFormMgr->OpenFormContainer(
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
}

// newstyle copy folder
void CMsgStoreDlg::HandleCopy()
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric, CLASS, L"OnCopyItems", L"\n");
	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return;

	auto lpMAPISourceFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY));

	const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	LPMAPIFOLDER lpSrcParentFolder = nullptr;
	WC_H(GetParentFolder(lpMAPISourceFolder, lpMDB, &lpSrcParentFolder));

	CGlobalCache::getInstance().SetFolderToCopy(lpMAPISourceFolder, lpSrcParentFolder);

	if (lpSrcParentFolder) lpSrcParentFolder->Release();
	if (lpMAPISourceFolder) lpMAPISourceFolder->Release();
}

_Check_return_ bool CMsgStoreDlg::HandlePaste()
{
	if (CBaseDialog::HandlePaste()) return true;

	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric, CLASS, L"HandlePaste", L"\n");
	if (!m_lpHierarchyTableTreeCtrl) return false;

	const auto ulStatus = CGlobalCache::getInstance().GetBufferStatus();

	// Get the destination Folder
	auto lpMAPIDestFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY));

	if (lpMAPIDestFolder && ulStatus & BUFFER_MESSAGES && ulStatus & BUFFER_PARENTFOLDER)
	{
		OnPasteMessages();
	}
	else if (lpMAPIDestFolder && ulStatus & BUFFER_FOLDER)
	{
		editor::CEditor MyData(
			this,
			IDS_PASTEFOLDER,
			IDS_PASTEFOLDERPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.InitPane(0, CheckPane::Create(IDS_PASTEFOLDERCONTENTS, false, false));
		WC_H(MyData.DisplayDialog());
		if (S_OK == hRes)
		{
			const auto bPasteContents = MyData.GetCheck(0);
			if (bPasteContents) OnPasteFolderContents();
			else OnPasteFolder();
		}
	}
	if (lpMAPIDestFolder) lpMAPIDestFolder->Release();
	return true;
}

void CMsgStoreDlg::OnPasteMessages()
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric, CLASS, L"OnPasteMessages", L"\n");
	if (!m_lpHierarchyTableTreeCtrl) return;

	// Get the source Messages
	const auto lpEIDs = CGlobalCache::getInstance().GetMessagesToCopy();
	auto lpMAPISourceFolder = CGlobalCache::getInstance().GetSourceParentFolder();
	// Get the destination Folder
	auto lpMAPIDestFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY));

	if (lpMAPIDestFolder && lpMAPISourceFolder && lpEIDs)
	{
		editor::CEditor MyData(
			this,
			IDS_COPYMESSAGE,
			IDS_COPYMESSAGEPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.InitPane(0, CheckPane::Create(IDS_MESSAGEMOVE, false, false));
		WC_H(MyData.DisplayDialog());
		if (S_OK == hRes)
		{
			auto ulMoveMessage = MyData.GetCheck(0) ? MESSAGE_MOVE : 0;

			LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IMAPIFolder::CopyMessages", m_hWnd); // STRING_OK

			if (lpProgress)
				ulMoveMessage |= MESSAGE_DIALOG;

			EC_MAPI(lpMAPISourceFolder->CopyMessages(
				lpEIDs,
				&IID_IMAPIFolder,
				lpMAPIDestFolder,
				lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL,
				lpProgress,
				ulMoveMessage));

			if (lpProgress)
				lpProgress->Release();
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
	static const SizedSPropTagArray(NUM_COLS, sptaSrcFolder) =
	{
		NUM_COLS,
		PR_DISPLAY_NAME_W,
		PR_ENTRYID
	};

	DebugPrintEx(DBGGeneric, CLASS, L"OnPasteFolder", L"\n");

	// Get the source folder
	auto lpMAPISourceFolder = CGlobalCache::getInstance().GetFolderToCopy();
	auto lpSrcParentFolder = CGlobalCache::getInstance().GetSourceParentFolder();
	// Get the Destination Folder
	auto lpMAPIDestFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY));

	if (lpMAPISourceFolder && lpMAPIDestFolder)
	{
		DebugPrint(DBGGeneric, L"Folder Source Object = %p\n", lpMAPISourceFolder);
		DebugPrint(DBGGeneric, L"Folder Source Object Parent = %p\n", lpSrcParentFolder);
		DebugPrint(DBGGeneric, L"Folder Destination Object = %p\n", lpMAPIDestFolder);

		editor::CEditor MyData(
			this,
			IDS_PASTEFOLDER,
			IDS_PASTEFOLDERNEWNAMEPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, TextPane::CreateSingleLinePane(IDS_FOLDERNAME, false));
		MyData.InitPane(1, CheckPane::Create(IDS_COPYSUBFOLDERS, false, false));
		MyData.InitPane(2, CheckPane::Create(IDS_FOLDERMOVE, false, false));

		// Get required properties from the source folder
		EC_H_GETPROPS(lpMAPISourceFolder->GetProps(
			LPSPropTagArray(&sptaSrcFolder),
			fMapiUnicode,
			&cProps,
			&lpProps));

		if (lpProps)
		{
			if (CheckStringProp(&lpProps[NAME], PT_UNICODE))
			{
				DebugPrint(DBGGeneric, L"Folder Source Name = \"%ws\"\n", lpProps[NAME].Value.lpszW);
				MyData.SetStringW(0, lpProps[NAME].Value.lpszW);
			}
		}

		WC_H(MyData.DisplayDialog());

		auto lpCopyRoot = lpSrcParentFolder;
		if (!lpSrcParentFolder) lpCopyRoot = lpMAPIDestFolder;

		if (S_OK == hRes)
		{
			CWaitCursor Wait; // Change the mouse to an hourglass while we work.

			LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IMAPIFolder::CopyFolder", m_hWnd); // STRING_OK

			auto ulCopyFlags = MAPI_UNICODE;
			if (MyData.GetCheck(1))
				ulCopyFlags |= COPY_SUBFOLDERS;
			if (MyData.GetCheck(2))
				ulCopyFlags |= FOLDER_MOVE;
			if (lpProgress)
				ulCopyFlags |= FOLDER_DIALOG;

			WC_MAPI(lpCopyRoot->CopyFolder(
				lpProps[EID].Value.bin.cb,
				reinterpret_cast<LPENTRYID>(lpProps[EID].Value.bin.lpb),
				&IID_IMAPIFolder,
				lpMAPIDestFolder,
				LPTSTR(MyData.GetStringW(0).c_str()),
				lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL,
				lpProgress, // Progress
				ulCopyFlags));
			if (MAPI_E_COLLISION == hRes)
			{
				error::ErrDialog(__FILE__, __LINE__, IDS_EDDUPEFOLDER);
			}
			else CHECKHRESMSG(hRes, IDS_COPYFOLDERFAILED);

			if (lpProgress)
				lpProgress->Release();
		}

		MAPIFreeBuffer(lpProps);
	}

	if (lpMAPIDestFolder) lpMAPIDestFolder->Release();
	if (lpMAPISourceFolder) lpMAPISourceFolder->Release();
	if (lpSrcParentFolder) lpSrcParentFolder->Release();
}

void CMsgStoreDlg::OnPasteFolderContents()
{
	auto hRes = S_OK;

	DebugPrintEx(DBGGeneric, CLASS, L"OnPasteFolderContents", L"\n");

	if (!m_lpHierarchyTableTreeCtrl) return;

	// Get the Source Folder
	auto lpMAPISourceFolder = CGlobalCache::getInstance().GetFolderToCopy();
	// Get the Destination Folder
	auto lpMAPIDestFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY));

	if (lpMAPISourceFolder && lpMAPIDestFolder)
	{
		DebugPrint(DBGGeneric, L"Folder Source Object = %p\n", lpMAPISourceFolder);
		DebugPrint(DBGGeneric, L"Folder Destination Object = %p\n", lpMAPIDestFolder);

		editor::CEditor MyData(
			this,
			IDS_COPYFOLDERCONTENTS,
			IDS_PICKOPTIONSPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, CheckPane::Create(IDS_COPYASSOCIATEDITEMS, false, false));
		MyData.InitPane(1, CheckPane::Create(IDS_MOVEMESSAGES, false, false));
		MyData.InitPane(2, CheckPane::Create(IDS_SINGLECALLCOPY, false, false));
		WC_H(MyData.DisplayDialog());

		if (S_OK == hRes)
		{
			CWaitCursor Wait; // Change the mouse to an hourglass while we work.

			EC_H(CopyFolderContents(
				lpMAPISourceFolder,
				lpMAPIDestFolder,
				MyData.GetCheck(0), // associated contents
				MyData.GetCheck(1), // move
				MyData.GetCheck(2), // Single CopyMessages call
				m_hWnd));
		}
	}
	if (lpMAPIDestFolder) lpMAPIDestFolder->Release();
	if (lpMAPISourceFolder) lpMAPISourceFolder->Release();
}

void CMsgStoreDlg::OnPasteRules()
{
	auto hRes = S_OK;

	DebugPrintEx(DBGGeneric, CLASS, L"OnPasteRules", L"\n");

	if (!m_lpHierarchyTableTreeCtrl) return;

	// Get the Source Folder
	auto lpMAPISourceFolder = CGlobalCache::getInstance().GetFolderToCopy();
	// Get the Destination Folder
	auto lpMAPIDestFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY));

	if (lpMAPISourceFolder && lpMAPIDestFolder)
	{
		DebugPrint(DBGGeneric, L"Folder Source Object = %p\n", lpMAPISourceFolder);
		DebugPrint(DBGGeneric, L"Folder Destination Object = %p\n", lpMAPIDestFolder);

		editor::CEditor MyData(
			this,
			IDS_COPYFOLDERRULES,
			IDS_COPYFOLDERRULESPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, CheckPane::Create(IDS_REPLACERULES, false, false));
		WC_H(MyData.DisplayDialog());

		if (S_OK == hRes)
		{
			CWaitCursor Wait; // Change the mouse to an hourglass while we work.

			EC_H(CopyFolderRules(
				lpMAPISourceFolder,
				lpMAPIDestFolder,
				MyData.GetCheck(0))); // move
		}
	}
	if (lpMAPIDestFolder) lpMAPIDestFolder->Release();
	if (lpMAPISourceFolder) lpMAPISourceFolder->Release();
}

void CMsgStoreDlg::OnCreateSubFolder()
{
	auto hRes = S_OK;
	LPMAPIFOLDER lpMAPISubFolder = nullptr;

	editor::CEditor MyData(
		this,
		IDS_ADDSUBFOLDER,
		IDS_ADDSUBFOLDERPROMPT,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.SetPromptPostFix(interpretprop::AllFlagsToString(PROP_ID(PR_FOLDER_TYPE), true));
	MyData.InitPane(0, TextPane::CreateSingleLinePaneID(IDS_FOLDERNAME, IDS_FOLDERNAMEVALUE, false));
	MyData.InitPane(1, TextPane::CreateSingleLinePane(IDS_FOLDERTYPE, false));
	MyData.SetHex(1, FOLDER_GENERIC);
	auto szProduct = strings::loadstring(ID_PRODUCTNAME);
	const auto szFolderComment = strings::formatmessage(IDS_FOLDERCOMMENTVALUE, szProduct.c_str());
	MyData.InitPane(2, TextPane::CreateSingleLinePane(IDS_FOLDERCOMMENT, szFolderComment, false));
	MyData.InitPane(3, CheckPane::Create(IDS_PASSOPENIFEXISTS, false, false));

	auto lpMAPIFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(
		mfcmapiREQUEST_MODIFY));

	if (lpMAPIFolder)
	{
		WC_H(MyData.DisplayDialog());

		EC_MAPI(lpMAPIFolder->CreateFolder(
			MyData.GetHex(1),
			LPTSTR(MyData.GetStringW(0).c_str()),
			LPTSTR(MyData.GetStringW(2).c_str()),
			NULL, // interface
			MAPI_UNICODE | (MyData.GetCheck(3) ? OPEN_IF_EXISTS : 0),
			&lpMAPISubFolder));

		if (lpMAPISubFolder) lpMAPISubFolder->Release();
		lpMAPIFolder->Release();
	}
}

void CMsgStoreDlg::OnDisplayACLTable()
{
	auto hRes = S_OK;

	if (!m_lpHierarchyTableTreeCtrl) return;

	auto lpMAPIFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY));

	if (lpMAPIFolder)
	{
		EC_H(DisplayExchangeTable(
			lpMAPIFolder,
			PR_ACL_TABLE,
			otACL,
			this));
		lpMAPIFolder->Release();
	}
}

void CMsgStoreDlg::OnDisplayAssociatedContents()
{
	if (!m_lpHierarchyTableTreeCtrl) return;

	// Find the highlighted item
	auto lpMAPIFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY));

	if (lpMAPIFolder)
	{
		(void)DisplayObject(
			lpMAPIFolder,
			NULL,
			otAssocContents,
			this);

		lpMAPIFolder->Release();
	}
}

void CMsgStoreDlg::OnDisplayDeletedContents()
{
	if (!m_lpHierarchyTableTreeCtrl || !m_lpMapiObjects) return;

	// Find the highlighted item
	auto hRes = S_OK;
	const auto lpItemEID = m_lpHierarchyTableTreeCtrl->GetSelectedItemEID();

	if (lpItemEID)
	{
		const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
		if (lpMDB)
		{
			LPMAPIFOLDER lpMAPIFolder = nullptr;
			WC_H(CallOpenEntry(
				lpMDB,
				NULL,
				NULL,
				NULL,
				lpItemEID->cb,
				reinterpret_cast<LPENTRYID>(lpItemEID->lpb),
				NULL,
				MAPI_BEST_ACCESS | SHOW_SOFT_DELETES | MAPI_NO_CACHE,
				NULL,
				reinterpret_cast<LPUNKNOWN*>(&lpMAPIFolder)));
			if (lpMAPIFolder)
			{
				// call the dialog
				new CFolderDlg(
					m_lpParent,
					m_lpMapiObjects,
					lpMAPIFolder,
					dfDeleted
				);
				lpMAPIFolder->Release();
			}
		}
	}
}

void CMsgStoreDlg::OnDisplayDeletedSubFolders()
{
	if (!m_lpHierarchyTableTreeCtrl) return;

	// Must open the folder with MODIFY permissions if I'm going to restore the folder!
	auto lpFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY));

	if (lpFolder)
	{
		new CMsgStoreDlg(
			m_lpParent,
			m_lpMapiObjects,
			lpFolder,
			dfDeleted);
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
	auto hRes = S_OK;

	if (!m_lpHierarchyTableTreeCtrl) return;

	// Find the highlighted item
	auto lpMAPIFolderToEmpty = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY));

	if (lpMAPIFolderToEmpty)
	{
		editor::CEditor MyData(
			this,
			IDS_DELETEITEMSANDSUB,
			IDS_DELETEITEMSANDSUBPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, CheckPane::Create(IDS_DELASSOCIATED, false, false));
		MyData.InitPane(1, CheckPane::Create(IDS_HARDDELETION, false, false));
		MyData.InitPane(2, CheckPane::Create(IDS_MANUALLYEMPTYFOLDER, false, false));

		WC_H(MyData.DisplayDialog());
		if (S_OK == hRes)
		{
			if (MyData.GetCheck(2))
			{
				EC_H(ManuallyEmptyFolder(lpMAPIFolderToEmpty, MyData.GetCheck(0), MyData.GetCheck(1)));
			}
			else
			{
				auto ulFlags = MyData.GetCheck(0) ? DEL_ASSOCIATED : 0;
				ulFlags |= MyData.GetCheck(1) ? DELETE_HARD_DELETE : 0;
				LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IMAPIFolder::EmptyFolder", m_hWnd); // STRING_OK

				if (lpProgress)
					ulFlags |= FOLDER_DIALOG;

				DebugPrintEx(DBGGeneric, CLASS, L"OnEmptyFolder", L"Calling EmptyFolder on %p, ulFlags = 0x%08X.\n", lpMAPIFolderToEmpty, ulFlags);

				EC_MAPI(lpMAPIFolderToEmpty->EmptyFolder(
					lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL,
					lpProgress,
					ulFlags));

				if (lpProgress)
					lpProgress->Release();
			}
		}
		lpMAPIFolderToEmpty->Release();
	}
}

void CMsgStoreDlg::OnDeleteSelectedItem()
{
	auto hRes = S_OK;
	LPSBinary lpItemEID = nullptr;

	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return;

	const ULONG bShiftPressed = GetKeyState(VK_SHIFT) < 0;

	const auto hItem = m_lpHierarchyTableTreeCtrl->GetSelectedItem();
	if (hItem)
	{
		const auto lpData = m_lpHierarchyTableTreeCtrl->GetSortListData(hItem);
		if (lpData && lpData->Node()) lpItemEID = lpData->Node()->m_lpEntryID;
	}

	if (!lpItemEID) return;

	const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return;

	auto lpFolderToDelete = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(
		mfcmapiDO_NOT_REQUEST_MODIFY));

	if (lpFolderToDelete)
	{
		LPMAPIFOLDER lpParentFolder = nullptr;
		EC_H(GetParentFolder(lpFolderToDelete, lpMDB, &lpParentFolder));
		if (lpParentFolder)
		{
			editor::CEditor MyData(
				this,
				IDS_DELETEFOLDER,
				IDS_DELETEFOLDERPROMPT,
				CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyData.InitPane(0, CheckPane::Create(IDS_HARDDELETION, false, false));
			if (!bShiftPressed)
				WC_H(MyData.DisplayDialog());
			if (S_OK == hRes)
			{
				auto ulFlags = DEL_FOLDERS | DEL_MESSAGES;
				ulFlags |= bShiftPressed || MyData.GetCheck(0) ? DELETE_HARD_DELETE : 0;

				DebugPrintEx(DBGDeleteSelectedItem, CLASS, L"OnDeleteSelectedItem", L"Calling DeleteFolder on folder. ulFlags = 0x%08X.\n", ulFlags);
				DebugPrintBinary(DBGGeneric, *lpItemEID);

				LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IMAPIFolder::DeleteFolder", m_hWnd); // STRING_OK

				if (lpProgress)
					ulFlags |= FOLDER_DIALOG;

				EC_MAPI(lpParentFolder->DeleteFolder(
					lpItemEID->cb,
					reinterpret_cast<LPENTRYID>(lpItemEID->lpb),
					lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL,
					lpProgress,
					ulFlags));

				// Delete the item from the UI since we cannot rely on notifications to handle this for us
				WC_B(m_lpHierarchyTableTreeCtrl->DeleteItem(hItem));

				if (lpProgress)
					lpProgress->Release();
			}

			lpParentFolder->Release();
		}

		lpFolderToDelete->Release();
	}
}

void CMsgStoreDlg::OnSaveFolderContentsAsMSG()
{
	auto hRes = S_OK;
	if (!m_lpHierarchyTableTreeCtrl) return;

	DebugPrintEx(DBGGeneric, CLASS, L"OnSaveFolderContentsAsMSG", L"\n");

	// Find the highlighted item
	auto lpMAPIFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiDO_NOT_REQUEST_MODIFY));
	if (!lpMAPIFolder) return;

	editor::CEditor MyData(
		this,
		IDS_SAVEFOLDERASMSG,
		IDS_SAVEFOLDERASMSGPROMPT,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CheckPane::Create(IDS_SAVEASSOCIATEDCONTENTS, false, false));
	MyData.InitPane(1, CheckPane::Create(IDS_SAVEUNICODE, false, false));
	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		auto szDir = file::GetDirectoryPath(m_hWnd);
		if (!szDir.empty())
		{
			CWaitCursor Wait; // Change the mouse to an hourglass while we work.

			EC_H(file::SaveFolderContentsToMSG(
				lpMAPIFolder,
				szDir,
				MyData.GetCheck(0) ? true : false,
				MyData.GetCheck(1) ? true : false,
				m_hWnd));
		}
	}

	lpMAPIFolder->Release();
}

void CMsgStoreDlg::OnSaveFolderContentsAsTextFiles()
{
	auto hRes = S_OK;

	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return;

	const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return;

	auto lpFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiDO_NOT_REQUEST_MODIFY));

	if (lpFolder)
	{
		editor::CEditor MyData(
			this,
			IDS_SAVEFOLDERASPROPFILES,
			IDS_PICKOPTIONSPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, CheckPane::Create(IDS_RECURSESUBFOLDERS, false, false));
		MyData.InitPane(1, CheckPane::Create(IDS_SAVEREGULARCONTENTS, true, false));
		MyData.InitPane(2, CheckPane::Create(IDS_SAVEASSOCIATEDCONTENTS, true, false));

		WC_H(MyData.DisplayDialog());

		if (S_OK == hRes)
		{
			file::SaveFolderContentsToTXT(
				lpMDB,
				lpFolder,
				MyData.GetCheck(1),
				MyData.GetCheck(2),
				MyData.GetCheck(0),
				m_hWnd);
		}
		lpFolder->Release();
	}
}

void CMsgStoreDlg::OnExportMessages()
{
	if (!m_lpHierarchyTableTreeCtrl) return;

	auto lpFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiDO_NOT_REQUEST_MODIFY));

	if (lpFolder)
	{
		file::ExportMessages(lpFolder, m_hWnd);

		lpFolder->Release();
	}
}

void CMsgStoreDlg::OnSetReceiveFolder()
{
	auto hRes = S_OK;

	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return;

	auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return;

	editor::CEditor MyData(
		this,
		IDS_SETRECFOLDER,
		IDS_SETRECFOLDERPROMPT,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, TextPane::CreateSingleLinePane(IDS_CLASS, false));
	MyData.InitPane(1, CheckPane::Create(IDS_DELETEASSOCIATION, false, false));

	// Find the highlighted item
	const auto lpEID = m_lpHierarchyTableTreeCtrl->GetSelectedItemEID();

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		if (MyData.GetCheck(1))
		{
			EC_MAPI(lpMDB->SetReceiveFolder(
				LPTSTR(MyData.GetStringW(0).c_str()),
				MAPI_UNICODE,
				NULL,
				NULL));
		}
		else if (lpEID)
		{
			EC_MAPI(lpMDB->SetReceiveFolder(
				LPTSTR(MyData.GetStringW(0).c_str()),
				MAPI_UNICODE,
				lpEID->cb,
				reinterpret_cast<LPENTRYID>(lpEID->lpb)));
		}
	}
}

void CMsgStoreDlg::OnResendAllMessages()
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpHierarchyTableTreeCtrl) return;

	// Find the highlighted item
	auto lpMAPIFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY));

	if (lpMAPIFolder)
	{
		EC_H(ResendMessages(lpMAPIFolder, m_hWnd));

		lpMAPIFolder->Release();
	}
}

// Iterate through items in the selected folder and attempt to delete PR_NT_SECURITY_DESCRIPTOR
void CMsgStoreDlg::OnResetPermissionsOnItems()
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return;

	const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release

	// Find the highlighted item
	auto lpMAPIFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY));

	if (lpMAPIFolder)
	{
		EC_H(ResetPermissionsOnItems(lpMDB, lpMAPIFolder));
		lpMAPIFolder->Release();
	}
}

// Copy selected folder back to the land of the living
void CMsgStoreDlg::OnRestoreDeletedFolder()
{
	auto hRes = S_OK;
	ULONG cProps;
	LPSPropValue lpProps = nullptr;

	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return;

	enum
	{
		NAME,
		EID,
		NUM_COLS
	};
	static const SizedSPropTagArray(NUM_COLS, sptaSrcFolder) =
	{
		NUM_COLS,
		PR_DISPLAY_NAME_W,
		PR_ENTRYID
	};

	const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return;

	auto lpSrcFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY));

	if (lpSrcFolder)
	{
		LPMAPIFOLDER lpSrcParentFolder = nullptr;
		WC_H(GetParentFolder(lpSrcFolder, lpMDB, &lpSrcParentFolder));
		hRes = S_OK;

		// Get required properties from the source folder
		EC_H_GETPROPS(lpSrcFolder->GetProps(
			LPSPropTagArray(&sptaSrcFolder),
			fMapiUnicode,
			&cProps,
			&lpProps));

		editor::CEditor MyData(
			this,
			IDS_RESTOREDELFOLD,
			IDS_RESTOREDELFOLDPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, TextPane::CreateSingleLinePane(IDS_FOLDERNAME, false));
		MyData.InitPane(1, CheckPane::Create(IDS_COPYSUBFOLDERS, false, false));

		if (lpProps)
		{
			if (CheckStringProp(&lpProps[NAME], PT_UNICODE))
			{
				DebugPrint(DBGGeneric, L"Folder Source Name = \"%ws\"\n", lpProps[NAME].Value.lpszW);
				MyData.SetStringW(0, lpProps[NAME].Value.lpszW);
			}
		}

		WC_H(MyData.DisplayDialog());

		if (S_OK == hRes)
		{
			// Restore the folder up under m_lpContainer
			CWaitCursor Wait; // Change the mouse to an hourglass while we work.

			DebugPrint(DBGGeneric, L"Restoring %p to %p as \n", lpSrcFolder, m_lpContainer);

			auto lpCopyRoot = lpSrcParentFolder;
			if (!lpSrcParentFolder) lpCopyRoot = static_cast<LPMAPIFOLDER>(m_lpContainer);

			LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IMAPIFolder::CopyFolder", m_hWnd); // STRING_OK

			auto ulCopyFlags = MAPI_UNICODE | (MyData.GetCheck(1) ? COPY_SUBFOLDERS : 0);

			if (lpProgress)
				ulCopyFlags |= FOLDER_DIALOG;

			WC_MAPI(lpCopyRoot->CopyFolder(
				lpProps[EID].Value.bin.cb,
				reinterpret_cast<LPENTRYID>(lpProps[EID].Value.bin.lpb),
				&IID_IMAPIFolder,
				static_cast<LPMAPIFOLDER>(m_lpContainer),
				LPTSTR(MyData.GetStringW(0).c_str()),
				lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL,
				lpProgress,
				ulCopyFlags));
			if (MAPI_E_COLLISION == hRes)
			{
				error::ErrDialog(__FILE__, __LINE__, IDS_EDDUPEFOLDER);
			}
			else if (MAPI_W_PARTIAL_COMPLETION == hRes)
			{
				error::ErrDialog(__FILE__, __LINE__, IDS_EDRESTOREFAILED);
			}
			else CHECKHRESMSG(hRes, IDS_COPYFOLDERFAILED);

			if (lpProgress)
				lpProgress->Release();
		}

		MAPIFreeBuffer(lpProps);
		if (lpSrcParentFolder) lpSrcParentFolder->Release();
		lpSrcFolder->Release();
	}
}

void CMsgStoreDlg::OnValidateIPMSubtree()
{
	auto hRes = S_OK;
	ULONG ulValues = 0;
	LPSPropValue lpProps = nullptr;
	LPMAPIERROR lpErr = nullptr;

	editor::CEditor MyData(
		this,
		IDS_VALIDATEIPMSUB,
		IDS_PICKOPTIONSPROMPT,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CheckPane::Create(IDS_MAPIFORCECREATE, false, false));
	MyData.InitPane(1, CheckPane::Create(IDS_MAPIFULLIPMTREE, false, false));

	if (!m_lpMapiObjects) return;

	const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB) return;

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		const auto ulFlags = (MyData.GetCheck(0) ? MAPI_FORCE_CREATE : 0) |
			(MyData.GetCheck(1) ? MAPI_FULL_IPM_TREE : 0);

		DebugPrintEx(DBGGeneric, CLASS, L"OnValidateIPMSubtree", L"ulFlags = 0x%08X\n", ulFlags);

		EC_MAPI(HrValidateIPMSubtree(
			lpMDB,
			ulFlags,
			&ulValues,
			&lpProps,
			&lpErr));
		EC_MAPIERR(fMapiUnicode, lpErr);
		MAPIFreeBuffer(lpErr);

		if (ulValues > 0 && lpProps)
		{
			DebugPrintEx(DBGGeneric, CLASS, L"OnValidateIPMSubtree", L"HrValidateIPMSubtree returned 0x%08X properties:\n", ulValues);
			DebugPrintProperties(DBGGeneric, ulValues, lpProps, lpMDB);
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
		lpParams->lpFolder = static_cast<LPMAPIFOLDER>(lpContainer); // GetSelectedContainer returns LPMAPIFOLDER
	}

	InvokeAddInMenu(lpParams);
}
