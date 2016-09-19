// MFCUtilityFunctions.h : Common functions for MFC MAPI

#include "stdafx.h"
#include "MFCUtilityFunctions.h"
#include "MAPIFunctions.h"
#include "MAPIStoreFunctions.h"
#include "ColumnTags.h"
#include "BaseDialog.h"
#include "MapiObjects.h"
#include "MsgStoreDlg.h"
#include "FolderDlg.h"
#include "AbDlg.h"
#include "Editor.h"
#include "SingleMAPIPropListCtrl.h"
#include "SingleRecipientDialog.h"
#include "AclDlg.h"
#include "RulesDlg.h"
#include "MailboxTableDlg.h"
#include "PublicFolderTableDlg.h"
#include "SmartView/SmartView.h"
#include "SingleMessageDialog.h"
#include "AttachmentsDlg.h"
#include "RecipientsDlg.h"
#include "InterpretProp2.h"

_Check_return_ HRESULT DisplayObject(
	_In_ LPMAPIPROP lpUnk,
	ULONG ulObjType,
	ObjectType tType,
	_In_ CBaseDialog* lpHostDlg)
{
	auto hRes = S_OK;

	if (!lpHostDlg || !lpUnk) return MAPI_E_INVALID_PARAMETER;

	auto lpMapiObjects = lpHostDlg->GetMapiObjects(); // do not release
	if (!lpMapiObjects) return MAPI_E_INVALID_PARAMETER;

	auto lpParentWnd = lpHostDlg->GetParentWnd(); // do not release
	if (!lpParentWnd) return MAPI_E_INVALID_PARAMETER;

	// If we weren't passed an object type, go get one - careful! Some objects lie!
	if (!ulObjType)
	{
		ulObjType = GetMAPIObjectType(static_cast<LPMAPIPROP>(lpUnk));
	}

	auto szFlags = InterpretNumberAsStringProp(ulObjType, PR_OBJECT_TYPE);
	DebugPrint(DBGGeneric, L"DisplayObject asked to display %p, with ObjectType of 0x%08X and MAPI type of 0x%08X = %ws\n",
		lpUnk,
		tType,
		ulObjType,
		szFlags.c_str());

	LPMDB lpMDB = nullptr;
	// call the dialog
	switch (ulObjType)
	{
		// #define MAPI_STORE ((ULONG) 0x00000001) /* Message Store */
	case MAPI_STORE:
		lpHostDlg->OnUpdateSingleMAPIPropListCtrl(lpUnk, nullptr);

		lpMDB = lpMapiObjects->GetMDB(); // do not release
		if (lpMDB) lpMDB->AddRef(); // hold on to this so that...
		lpMapiObjects->SetMDB(static_cast<LPMDB>(lpUnk));

		new CMsgStoreDlg(
			lpParentWnd,
			lpMapiObjects,
			nullptr,
			otStoreDeletedItems == tType ? dfDeleted : dfNormal);

		// restore the old MDB
		lpMapiObjects->SetMDB(lpMDB); // ...we can put it back
		if (lpMDB) lpMDB->Release();
		break;
		// #define MAPI_FOLDER ((ULONG) 0x00000003) /* Folder */
	case MAPI_FOLDER:
		// There are two ways to display a folder...either the contents table or the hierarchy table.
		if (otHierarchy == tType)
		{
			lpMDB = lpMapiObjects->GetMDB(); // do not release
			if (lpMDB)
			{
				new CMsgStoreDlg(
					lpParentWnd,
					lpMapiObjects,
					static_cast<LPMAPIFOLDER>(lpUnk),
					dfNormal);
			}
			else
			{
				// Since lpMDB was NULL, let's get a good MDB
				auto lpMAPISession = lpMapiObjects->GetSession(); // do not release
				if (lpMAPISession)
				{
					LPMDB lpNewMDB = nullptr;
					EC_H(OpenStoreFromMAPIProp(lpMAPISession, static_cast<LPMAPIPROP>(lpUnk), &lpNewMDB));
					if (lpNewMDB)
					{
						lpMapiObjects->SetMDB(lpNewMDB);

						new CMsgStoreDlg(
							lpParentWnd,
							lpMapiObjects,
							static_cast<LPMAPIFOLDER>(lpUnk),
							dfNormal);

						// restore the old MDB
						lpMapiObjects->SetMDB(nullptr);
						lpNewMDB->Release();
					}
				}
			}
		}
		else if (otContents == tType || otAssocContents == tType)
		{
			new CFolderDlg(
				lpParentWnd,
				lpMapiObjects,
				static_cast<LPMAPIFOLDER>(lpUnk),
				otAssocContents == tType ? dfAssoc : dfNormal);
		}
		break;
		// #define MAPI_ABCONT ((ULONG) 0x00000004) /* Address Book Container */
	case MAPI_ABCONT:
		new CAbDlg(
			lpParentWnd,
			lpMapiObjects,
			static_cast<LPABCONT>(lpUnk));
		break;
		// #define MAPI_MESSAGE ((ULONG) 0x00000005) /* Message */
	case MAPI_MESSAGE:
		new SingleMessageDialog(
			lpParentWnd,
			lpMapiObjects,
			static_cast<LPMESSAGE>(lpUnk));
		break;
		// #define MAPI_MAILUSER ((ULONG) 0x00000006) /* Individual Recipient */
	case MAPI_MAILUSER:
		new SingleRecipientDialog(
			lpParentWnd,
			lpMapiObjects,
			static_cast<LPMAILUSER>(lpUnk));
		break;
		// #define MAPI_DISTLIST ((ULONG) 0x00000008) /* Distribution List Recipient */
	case MAPI_DISTLIST: // A DistList is really an Address book anyways
		new SingleRecipientDialog(
			lpParentWnd,
			lpMapiObjects,
			static_cast<LPMAILUSER>(lpUnk));
		new CAbDlg(
			lpParentWnd,
			lpMapiObjects,
			static_cast<LPABCONT>(lpUnk));
		break;
		// The following types don't have special viewers - just dump their props in the property pane
		// #define MAPI_ADDRBOOK ((ULONG) 0x00000002) /* Address Book */
		// #define MAPI_ATTACH ((ULONG) 0x00000007) /* Attachment */
		// #define MAPI_PROFSECT ((ULONG) 0x00000009) /* Profile Section */
		// #define MAPI_STATUS ((ULONG) 0x0000000A) /* Status Object */
		// #define MAPI_SESSION ((ULONG) 0x0000000B) /* Session */
		// #define MAPI_FORMINFO ((ULONG) 0x0000000C) /* Form Information */
	default:
		lpHostDlg->OnUpdateSingleMAPIPropListCtrl(lpUnk, nullptr);

		szFlags = InterpretNumberAsStringProp(ulObjType, PR_OBJECT_TYPE);
		DebugPrint(DBGGeneric,
			L"DisplayObject: Object type: 0x%08X = %ws not implemented\r\n" // STRING_OK
			L"This is not an error. It just means no specialized viewer has been implemented for this object type.", // STRING_OK
			ulObjType,
			szFlags.c_str());
		break;
	}

	return hRes;
}

_Check_return_ HRESULT DisplayTable(
	_In_ LPMAPITABLE lpTable,
	ObjectType tType,
	_In_ CBaseDialog* lpHostDlg)
{
	if (!lpHostDlg) return MAPI_E_INVALID_PARAMETER;

	auto lpMapiObjects = lpHostDlg->GetMapiObjects(); // do not release
	if (!lpMapiObjects) return MAPI_E_INVALID_PARAMETER;

	auto lpParentWnd = lpHostDlg->GetParentWnd(); // do not release
	if (!lpParentWnd) return MAPI_E_INVALID_PARAMETER;

	DebugPrint(DBGGeneric, L"DisplayTable asked to display %p\n", lpTable);

	switch (tType)
	{
	case otStatus:
	{
		if (!lpTable) return MAPI_E_INVALID_PARAMETER;
		new CContentsTableDlg(
			lpParentWnd,
			lpMapiObjects,
			IDS_STATUSTABLE,
			mfcmapiCALL_CREATE_DIALOG,
			lpTable,
			LPSPropTagArray(&sptSTATUSCols),
			NUMSTATUSCOLUMNS,
			STATUSColumns,
			NULL,
			MENU_CONTEXT_STATUS_TABLE);
		break;
	}
	case otReceive:
	{
		if (!lpTable) return MAPI_E_INVALID_PARAMETER;
		new CContentsTableDlg(
			lpParentWnd,
			lpMapiObjects,
			IDS_RECEIVEFOLDERTABLE,
			mfcmapiCALL_CREATE_DIALOG,
			lpTable,
			LPSPropTagArray(&sptRECEIVECols),
			NUMRECEIVECOLUMNS,
			RECEIVEColumns,
			NULL,
			MENU_CONTEXT_RECIEVE_FOLDER_TABLE);
		break;
	}
	case otHierarchy:
	{
		if (!lpTable) return MAPI_E_INVALID_PARAMETER;
		new CContentsTableDlg(
			lpParentWnd,
			lpMapiObjects,
			IDS_HIERARCHYTABLE,
			mfcmapiCALL_CREATE_DIALOG,
			lpTable,
			LPSPropTagArray(&sptHIERARCHYCols),
			NUMHIERARCHYCOLUMNS,
			HIERARCHYColumns,
			NULL,
			MENU_CONTEXT_HIER_TABLE);
		break;
	}
	default:
	case otDefault:
	{
		if (!lpTable) return MAPI_E_INVALID_PARAMETER;
		if (otDefault != tType) ErrDialog(__FILE__, __LINE__, IDS_EDDISPLAYTABLE, tType);
		new CContentsTableDlg(
			lpParentWnd,
			lpMapiObjects,
			IDS_CONTENTSTABLE,
			mfcmapiCALL_CREATE_DIALOG,
			lpTable,
			LPSPropTagArray(&sptDEFCols),
			NUMDEFCOLUMNS,
			DEFColumns,
			NULL,
			MENU_CONTEXT_DEFAULT_TABLE);
		break;
	}
	}

	return S_OK;
}

_Check_return_ HRESULT DisplayTable(
	_In_ LPMAPIPROP lpMAPIProp,
	ULONG ulPropTag,
	ObjectType tType,
	_In_ CBaseDialog* lpHostDlg)
{
	auto hRes = S_OK;
	LPMAPITABLE lpTable = nullptr;

	if (!lpHostDlg || !lpMAPIProp) return MAPI_E_INVALID_PARAMETER;
	if (PT_OBJECT != PROP_TYPE(ulPropTag)) return MAPI_E_INVALID_TYPE;

	WC_MAPI(lpMAPIProp->OpenProperty(
		ulPropTag,
		&IID_IMAPITable,
		fMapiUnicode,
		0,
		reinterpret_cast<LPUNKNOWN *>(&lpTable)));
	if (MAPI_E_INTERFACE_NOT_SUPPORTED == hRes)
	{
		hRes = S_OK;
		switch (PROP_ID(ulPropTag))
		{
		case PROP_ID(PR_MESSAGE_ATTACHMENTS):
			static_cast<LPMESSAGE>(lpMAPIProp)->GetAttachmentTable(
				NULL,
				&lpTable);
			break;
		case PROP_ID(PR_MESSAGE_RECIPIENTS):
			static_cast<LPMESSAGE>(lpMAPIProp)->GetRecipientTable(
				NULL,
				&lpTable);
			break;
		}
	}

	if (lpTable)
	{
		switch (PROP_ID(ulPropTag))
		{
		case PROP_ID(PR_MESSAGE_ATTACHMENTS):
			new CAttachmentsDlg(
				lpHostDlg->GetParentWnd(),
				lpHostDlg->GetMapiObjects(),
				lpTable,
				static_cast<LPMESSAGE>(lpMAPIProp));
			break;
		case PROP_ID(PR_MESSAGE_RECIPIENTS):
			new CRecipientsDlg(
				lpHostDlg->GetParentWnd(),
				lpHostDlg->GetMapiObjects(),
				lpTable,
				static_cast<LPMESSAGE>(lpMAPIProp));
			break;
		default:
			EC_H(DisplayTable(
				lpTable,
				tType,
				lpHostDlg));
			break;
		}

		lpTable->Release();
	}

	return hRes;
}

_Check_return_ HRESULT DisplayExchangeTable(
	_In_ LPMAPIPROP lpMAPIProp,
	ULONG ulPropTag,
	ObjectType tType,
	_In_ CBaseDialog* lpHostDlg)
{
	auto hRes = S_OK;
	LPEXCHANGEMODIFYTABLE lpExchTbl = nullptr;
	LPMAPITABLE lpMAPITable = nullptr;

	if (!lpMAPIProp || !lpHostDlg) return MAPI_E_INVALID_PARAMETER;

	auto lpMapiObjects = lpHostDlg->GetMapiObjects(); // do not release
	if (!lpMapiObjects) return MAPI_E_INVALID_PARAMETER;

	auto lpParentWnd = lpHostDlg->GetParentWnd(); // do not release
	if (!lpParentWnd) return MAPI_E_INVALID_PARAMETER;

	// Open the table in an IExchangeModifyTable interface
	EC_MAPI(lpMAPIProp->OpenProperty(
		ulPropTag,
		const_cast<LPGUID>(&IID_IExchangeModifyTable),
		0,
		MAPI_DEFERRED_ERRORS,
		reinterpret_cast<LPUNKNOWN*>(&lpExchTbl)));

	if (lpExchTbl)
	{
		switch (tType)
		{
		case otRules:
			new CRulesDlg(
				lpParentWnd,
				lpMapiObjects,
				lpExchTbl);
			break;
		case otACL:
		{
			CEditor MyData(
				lpHostDlg,
				IDS_ACLTABLE,
				IDS_ACLTABLEPROMPT,
				1,
				CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

			MyData.InitPane(0, CreateCheckPane(IDS_FBRIGHTSVISIBLE, false, false));

			WC_H(MyData.DisplayDialog());
			if (S_OK == hRes)
			{
				new CAclDlg(
					lpParentWnd,
					lpMapiObjects,
					lpExchTbl,
					MyData.GetCheck(0));
			}

			if (MAPI_E_USER_CANCEL == hRes)
				hRes = S_OK; // don't propogate the error past here
		}
		break;
		default:
			// Open a MAPI table on the Exchange table property. This table can be
			// read to determine what the Exchange table looks like.
			EC_MAPI(lpExchTbl->GetTable(0, &lpMAPITable));

			if (lpMAPITable)
			{
				EC_H(DisplayTable(
					lpMAPITable,
					tType,
					lpHostDlg));
				lpMAPITable->Release();
			}

			break;
		}
		lpExchTbl->Release();
	}
	return hRes;
}

_Check_return_ bool bShouldCancel(_In_opt_ CWnd* cWnd, HRESULT hResPrev)
{
	auto bGotError = false;

	if (S_OK != hResPrev)
	{
		if (MAPI_E_USER_CANCEL != hResPrev && MAPI_E_CANCEL != hResPrev)
		{
			bGotError = true;
		}

		auto hRes = S_OK;

		CEditor Cancel(
			cWnd,
			ID_PRODUCTNAME,
			IDS_CANCELPROMPT,
			bGotError ? 1 : 0,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		if (bGotError)
		{
			auto szPrevErr = formatmessage(IDS_PREVIOUSCALL, ErrorNameFromErrorCode(hResPrev).c_str(), hResPrev);
			Cancel.InitPane(0, CreateSingleLinePane(IDS_ERROR, szPrevErr, true));
		}
		WC_H(Cancel.DisplayDialog());
		if (S_OK != hRes)
		{
			DebugPrint(DBGGeneric, L"bShouldCancel: User asked to cancel\n");
			return true;
		}
	}
	return false;
}

void DisplayMailboxTable(_In_ CParentWnd* lpParent,
	_In_ CMapiObjects* lpMapiObjects)
{
	if (!lpParent || !lpMapiObjects) return;
	auto hRes = S_OK;
	LPMDB lpPrivateMDB = nullptr;
	auto lpMDB = lpMapiObjects->GetMDB(); // do not release
	auto lpMAPISession = lpMapiObjects->GetSession(); // do not release

	// try the 'current' MDB first
	if (!StoreSupportsManageStore(lpMDB))
	{
		// if that MDB doesn't support manage store, try to get one that does
		EC_H(OpenMessageStoreGUID(lpMAPISession, pbExchangeProviderPrimaryUserGuid, &lpPrivateMDB));
		lpMDB = lpPrivateMDB;
	}

	if (lpMDB && StoreSupportsManageStore(lpMDB))
	{
		LPMAPITABLE lpMailboxTable = nullptr;
		auto szServerName = stringTowstring(GetServerName(lpMAPISession));

		CEditor MyData(
			static_cast<CWnd*>(lpParent),
			IDS_DISPLAYMAILBOXTABLE,
			IDS_SERVERNAMEPROMPT,
			4,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, CreateSingleLinePane(IDS_SERVERNAME, szServerName, false));
		MyData.InitPane(1, CreateSingleLinePane(IDS_OFFSET, false));
		MyData.SetHex(1, 0);
		MyData.InitPane(2, CreateSingleLinePane(IDS_MAILBOXGUID, false));
		UINT uidDropDown[] = {
		IDS_GETMBXINTERFACE1,
		IDS_GETMBXINTERFACE3,
		IDS_GETMBXINTERFACE5
		};
		MyData.InitPane(3, CreateDropDownPane(IDS_GETMBXINTERFACE, _countof(uidDropDown), uidDropDown, true));
		WC_H(MyData.DisplayDialog());

		if (SUCCEEDED(hRes) && 0 != MyData.GetHex(1) && 0 == MyData.GetDropDown(3))
		{
			ErrDialog(__FILE__, __LINE__, IDS_EDOFFSETWITHWRONGINTERFACE);
		}

		else if (S_OK == hRes)
		{
			auto szServerDN = BuildServerDN(
				wstringTostring(MyData.GetStringW(0)),
				"");
			if (!szServerDN.empty())
			{
				LPMDB lpOldMDB = nullptr;

				// if we got a new MDB, set it in lpMapiObjects
				if (lpPrivateMDB)
				{
					lpOldMDB = lpMapiObjects->GetMDB(); // do not release
					if (lpOldMDB) lpOldMDB->AddRef(); // hold on to this so that...
					// If we don't do this, we crash when destroying the Mailbox Table Window
					lpMapiObjects->SetMDB(lpMDB);
				}
				switch (MyData.GetDropDown(3))
				{
				case 0:
					EC_H(GetMailboxTable1(
						lpMDB,
						szServerDN,
						fMapiUnicode,
						&lpMailboxTable));
					break;
				case 1:
					EC_H(GetMailboxTable3(
						lpMDB,
						szServerDN,
						MyData.GetHex(1),
						fMapiUnicode,
						&lpMailboxTable));
					break;
				case 2:
				{
					GUID MyGUID = { 0 };
					auto bHaveGUID = false;

					auto pszGUID = MyData.GetStringW(2);

					if (!pszGUID.empty())
					{
						bHaveGUID = true;

						MyGUID = StringToGUID(pszGUID);
						if (MyGUID == GUID_NULL)
						{
							ErrDialog(__FILE__, __LINE__, IDS_EDINVALIDGUID);
							break;
						}
					}

					EC_H(GetMailboxTable5(
						lpMDB,
						szServerDN,
						MyData.GetHex(1),
						fMapiUnicode,
						bHaveGUID ? &MyGUID : 0,
						&lpMailboxTable));
					break;
				}
				}

				if (SUCCEEDED(hRes) && lpMailboxTable)
				{
					new CMailboxTableDlg(
						lpParent,
						lpMapiObjects,
						MyData.GetStringW(0),
						lpMailboxTable);
				}
				else if (MAPI_E_NO_ACCESS == hRes || MAPI_E_NETWORK_ERROR == hRes)
				{
					ErrDialog(__FILE__, __LINE__,
						IDS_EDGETMAILBOXTABLEFAILED,
						_T("GetMailboxTable"), _T("GetMailboxTable")); // STRING_OK
				}
				if (lpMailboxTable) lpMailboxTable->Release();

				if (lpOldMDB)
				{
					lpMapiObjects->SetMDB(lpOldMDB); // ...we can put it back
					if (lpOldMDB) lpOldMDB->Release();
				}
			}
		}
	}

	if (lpPrivateMDB) lpPrivateMDB->Release();
}

void DisplayPublicFolderTable(_In_ CParentWnd* lpParent,
	_In_ CMapiObjects* lpMapiObjects)
{
	if (!lpParent || !lpMapiObjects) return;
	auto hRes = S_OK;
	LPMDB lpPrivateMDB = nullptr;
	auto lpMDB = lpMapiObjects->GetMDB(); // do not release
	auto lpMAPISession = lpMapiObjects->GetSession(); // do not release

	// try the 'current' MDB first
	if (!StoreSupportsManageStore(lpMDB))
	{
		// if that MDB doesn't support manage store, try to get one that does
		EC_H(OpenMessageStoreGUID(lpMAPISession, pbExchangeProviderPrimaryUserGuid, &lpPrivateMDB));
		lpMDB = lpPrivateMDB;
	}

	if (lpMDB && StoreSupportsManageStore(lpMDB))
	{
		LPMAPITABLE lpPFTable = nullptr;
		auto szServerName = stringTowstring(GetServerName(lpMAPISession));

		CEditor MyData(
			static_cast<CWnd*>(lpParent),
			IDS_DISPLAYPFTABLE,
			IDS_DISPLAYPFTABLEPROMPT,
			5,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, CreateSingleLinePane(IDS_SERVERNAME, szServerName, false));
		MyData.InitPane(1, CreateSingleLinePane(IDS_OFFSET, false));
		MyData.SetHex(1, 0);
		MyData.InitPane(2, CreateSingleLinePane(IDS_FLAGS, false));
		MyData.SetHex(2, MDB_IPM);
		MyData.InitPane(3, CreateSingleLinePane(IDS_PUBLICFOLDERGUID, false));
		UINT uidDropDown[] = {
		IDS_GETPFINTERFACE1,
		IDS_GETPFINTERFACE4,
		IDS_GETPFINTERFACE5
		};
		MyData.InitPane(4, CreateDropDownPane(IDS_GETMBXINTERFACE, _countof(uidDropDown), uidDropDown, true));
		WC_H(MyData.DisplayDialog());

		if (SUCCEEDED(hRes) && 0 != MyData.GetHex(1) && 0 == MyData.GetDropDown(4))
		{
			ErrDialog(__FILE__, __LINE__, IDS_EDOFFSETWITHWRONGINTERFACE);
		}

		else if (S_OK == hRes)
		{
			auto szServerDN = BuildServerDN(
				wstringTostring(MyData.GetStringW(0)),
				"");
			if (!szServerDN.empty())
			{
				LPMDB lpOldMDB = nullptr;

				// if we got a new MDB, set it in lpMapiObjects
				if (lpPrivateMDB)
				{
					lpOldMDB = lpMapiObjects->GetMDB(); // do not release
					if (lpOldMDB) lpOldMDB->AddRef(); // hold on to this so that...
					// If we don't do this, we crash when destroying the Mailbox Table Window
					lpMapiObjects->SetMDB(lpMDB);
				}

				switch (MyData.GetDropDown(4))
				{
				case 0:
					EC_H(GetPublicFolderTable1(
						lpMDB,
						szServerDN,
						MyData.GetHex(2) | fMapiUnicode,
						&lpPFTable));
					break;
				case 1:
					EC_H(GetPublicFolderTable4(
						lpMDB,
						szServerDN,
						MyData.GetHex(1),
						MyData.GetHex(2) | fMapiUnicode,
						&lpPFTable));
					break;
				case 2:
				{
					GUID MyGUID = { 0 };
					auto bHaveGUID = false;

					auto pszGUID = MyData.GetStringW(3);
					if (!pszGUID.empty())
					{
						bHaveGUID = true;

						MyGUID = StringToGUID(pszGUID);
						if (MyGUID == GUID_NULL)
						{
							ErrDialog(__FILE__, __LINE__, IDS_EDINVALIDGUID);
							break;
						}
					}

					EC_H(GetPublicFolderTable5(
						lpMDB,
						szServerDN,
						MyData.GetHex(1),
						MyData.GetHex(2) | fMapiUnicode,
						bHaveGUID ? &MyGUID : 0,
						&lpPFTable));
					break;
				}
				}

				if (SUCCEEDED(hRes) && lpPFTable)
				{
					new CPublicFolderTableDlg(
						lpParent,
						lpMapiObjects,
						MyData.GetStringW(0),
						lpPFTable);
				}
				else if (MAPI_E_NO_ACCESS == hRes || MAPI_E_NETWORK_ERROR == hRes)
				{
					ErrDialog(__FILE__, __LINE__,
						IDS_EDGETMAILBOXTABLEFAILED,
						_T("GetPublicFolderTable"), _T("GetPublicFolderTable")); // STRING_OK
				}
				if (lpPFTable) lpPFTable->Release();

				if (lpOldMDB)
				{
					lpMapiObjects->SetMDB(lpOldMDB); // ...we can put it back
					if (lpOldMDB) lpOldMDB->Release();
				}
			}
		}
	}

	if (lpPrivateMDB) lpPrivateMDB->Release();
}

void ResolveMessageClass(_In_ CMapiObjects* lpMapiObjects, _In_opt_ LPMAPIFOLDER lpMAPIFolder, _Out_ LPMAPIFORMINFO* lppMAPIFormInfo)
{
	auto hRes = S_OK;
	LPMAPIFORMMGR lpMAPIFormMgr = nullptr;
	if (!lpMapiObjects || !lppMAPIFormInfo) return;

	*lppMAPIFormInfo = nullptr;

	auto lpMAPISession = lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	EC_MAPI(MAPIOpenFormMgr(lpMAPISession, &lpMAPIFormMgr));
	if (lpMAPIFormMgr)
	{
		DebugPrint(DBGForms, L"OnResolveMessageClass: resolving message class\n");
		CEditor MyData(
			nullptr,
			IDS_RESOLVECLASS,
			IDS_RESOLVECLASSPROMPT,
			static_cast<ULONG>(2),
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, CreateSingleLinePane(IDS_CLASS, false));
		MyData.InitPane(1, CreateSingleLinePane(IDS_FLAGS, false));

		WC_H(MyData.DisplayDialog());
		if (S_OK == hRes)
		{
			auto szClass = MyData.GetStringW(0); // ResolveMessageClass requires an ANSI string
			auto ulFlags = MyData.GetHex(1);
			if (!szClass.empty())
			{
				LPMAPIFORMINFO lpMAPIFormInfo = nullptr;
				DebugPrint(DBGForms, L"OnResolveMessageClass: Calling ResolveMessageClass(\"%ws\",0x%08X)\n", szClass.c_str(), ulFlags); // STRING_OK
				EC_MAPI(lpMAPIFormMgr->ResolveMessageClass(wstringTostring(szClass).c_str(), ulFlags, lpMAPIFolder, &lpMAPIFormInfo));
				if (lpMAPIFormInfo)
				{
					DebugPrintFormInfo(DBGForms, lpMAPIFormInfo);
					*lppMAPIFormInfo = lpMAPIFormInfo;
				}
			}
		}

		lpMAPIFormMgr->Release();
	}
}

void SelectForm(_In_ HWND hWnd, _In_ CMapiObjects* lpMapiObjects, _In_opt_ LPMAPIFOLDER lpMAPIFolder, _Out_ LPMAPIFORMINFO* lppMAPIFormInfo)
{
	auto hRes = S_OK;
	LPMAPIFORMMGR lpMAPIFormMgr = nullptr;

	if (!lpMapiObjects || !lppMAPIFormInfo) return;

	*lppMAPIFormInfo = nullptr;

	auto lpMAPISession = lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	EC_MAPI(MAPIOpenFormMgr(lpMAPISession, &lpMAPIFormMgr));

	if (lpMAPIFormMgr)
	{
		LPMAPIFORMINFO lpMAPIFormInfo = nullptr;
		// Apparently, SelectForm doesn't support unicode
		// CString doesn't provide a way to extract just ANSI strings, so we do this manually
		auto szTitle = wstringTostring(loadstring(IDS_SELECTFORMPROPS));
		EC_H_CANCEL(lpMAPIFormMgr->SelectForm(
			reinterpret_cast<ULONG_PTR>(hWnd),
			0, // fMapiUnicode,
			LPCTSTR(szTitle.c_str()),
			lpMAPIFolder,
			&lpMAPIFormInfo));

		if (lpMAPIFormInfo)
		{
			DebugPrintFormInfo(DBGForms, lpMAPIFormInfo);
			*lppMAPIFormInfo = lpMAPIFormInfo;
		}

		lpMAPIFormMgr->Release();
	}
}