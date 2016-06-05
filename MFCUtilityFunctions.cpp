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
#include "SmartView\SmartView.h"
#include "SingleMessageDialog.h"
#include "AttachmentsDlg.h"
#include "RecipientsDlg.h"

_Check_return_ HRESULT DisplayObject(
	_In_ LPMAPIPROP lpUnk,
	ULONG ulObjType,
	ObjectType tType,
	_In_ CBaseDialog* lpHostDlg)
{
	HRESULT			hRes = S_OK;

	if (!lpHostDlg || !lpUnk) return MAPI_E_INVALID_PARAMETER;

	CMapiObjects* lpMapiObjects = lpHostDlg->GetMapiObjects(); // do not release
	if (!lpMapiObjects) return MAPI_E_INVALID_PARAMETER;

	CParentWnd* lpParentWnd = lpHostDlg->GetParentWnd(); // do not release
	if (!lpParentWnd) return MAPI_E_INVALID_PARAMETER;

	// If we weren't passed an object type, go get one - careful! Some objects lie!
	if (!ulObjType)
	{
		ulObjType = GetMAPIObjectType((LPMAPIPROP)lpUnk);
	}

	wstring szFlags = InterpretNumberAsStringProp(ulObjType, PR_OBJECT_TYPE);
	DebugPrint(DBGGeneric, L"DisplayObject asked to display %p, with ObjectType of 0x%08X and MAPI type of 0x%08X = %ws\n",
		lpUnk,
		tType,
		ulObjType,
		szFlags.c_str());

	LPMDB lpMDB = NULL;
	// call the dialog
	switch (ulObjType)
	{
		// #define MAPI_STORE		((ULONG) 0x00000001)	/* Message Store */
	case MAPI_STORE:
		lpHostDlg->OnUpdateSingleMAPIPropListCtrl(lpUnk, NULL);

		lpMDB = lpMapiObjects->GetMDB(); // do not release
		if (lpMDB) lpMDB->AddRef(); // hold on to this so that...
		lpMapiObjects->SetMDB((LPMDB)lpUnk);

		new CMsgStoreDlg(
			lpParentWnd,
			lpMapiObjects,
			NULL,
			(otStoreDeletedItems == tType) ? dfDeleted : dfNormal);

		// restore the old MDB
		lpMapiObjects->SetMDB(lpMDB); // ...we can put it back
		if (lpMDB) lpMDB->Release();
		break;
		// #define MAPI_FOLDER		((ULONG) 0x00000003)	/* Folder */
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
					(LPMAPIFOLDER)lpUnk,
					dfNormal);
			}
			else
			{
				// Since lpMDB was NULL, let's get a good MDB
				LPMAPISESSION lpMAPISession = lpMapiObjects->GetSession(); // do not release
				if (lpMAPISession)
				{
					LPMDB lpNewMDB = NULL;
					EC_H(OpenStoreFromMAPIProp(lpMAPISession, (LPMAPIPROP)lpUnk, &lpNewMDB));
					if (lpNewMDB)
					{
						lpMapiObjects->SetMDB(lpNewMDB);

						new CMsgStoreDlg(
							lpParentWnd,
							lpMapiObjects,
							(LPMAPIFOLDER)lpUnk,
							dfNormal);

						// restore the old MDB
						lpMapiObjects->SetMDB(NULL);
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
				(LPMAPIFOLDER)lpUnk,
				(otAssocContents == tType) ? dfAssoc : dfNormal);
		}
		break;
		// #define MAPI_ABCONT		((ULONG) 0x00000004)	/* Address Book Container */
	case MAPI_ABCONT:
		new CAbDlg(
			lpParentWnd,
			lpMapiObjects,
			(LPABCONT)lpUnk);
		break;
		// #define MAPI_MESSAGE	((ULONG) 0x00000005)	/* Message */
	case MAPI_MESSAGE:
		new SingleMessageDialog(
			lpParentWnd,
			lpMapiObjects,
			(LPMESSAGE)lpUnk);
		break;
		// #define MAPI_MAILUSER	((ULONG) 0x00000006)	/* Individual Recipient */
	case MAPI_MAILUSER:
		new SingleRecipientDialog(
			lpParentWnd,
			lpMapiObjects,
			(LPMAILUSER)lpUnk);
		break;
		// #define MAPI_DISTLIST	((ULONG) 0x00000008)	/* Distribution List Recipient */
	case MAPI_DISTLIST: // A DistList is really an Address book anyways
		new SingleRecipientDialog(
			lpParentWnd,
			lpMapiObjects,
			(LPMAILUSER)lpUnk);
		new CAbDlg(
			lpParentWnd,
			lpMapiObjects,
			(LPABCONT)lpUnk);
		break;
		// The following types don't have special viewers - just dump their props in the property pane
		// #define MAPI_ADDRBOOK	((ULONG) 0x00000002)	/* Address Book */
		// #define MAPI_ATTACH	((ULONG) 0x00000007)	/* Attachment */
		// #define MAPI_PROFSECT	((ULONG) 0x00000009)	/* Profile Section */
		// #define MAPI_STATUS	((ULONG) 0x0000000A)	/* Status Object */
		// #define MAPI_SESSION	((ULONG) 0x0000000B)	/* Session */
		// #define MAPI_FORMINFO	((ULONG) 0x0000000C)	/* Form Information */
	default:
		lpHostDlg->OnUpdateSingleMAPIPropListCtrl(lpUnk, NULL);

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

	CMapiObjects* lpMapiObjects = lpHostDlg->GetMapiObjects(); // do not release
	if (!lpMapiObjects) return MAPI_E_INVALID_PARAMETER;

	CParentWnd* lpParentWnd = lpHostDlg->GetParentWnd(); // do not release
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
			(LPSPropTagArray)&sptSTATUSCols,
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
			(LPSPropTagArray)&sptRECEIVECols,
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
			(LPSPropTagArray)&sptHIERARCHYCols,
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
			(LPSPropTagArray)&sptDEFCols,
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
	HRESULT		hRes = S_OK;
	LPMAPITABLE	lpTable = NULL;

	if (!lpHostDlg || !lpMAPIProp) return MAPI_E_INVALID_PARAMETER;
	if (PT_OBJECT != PROP_TYPE(ulPropTag)) return MAPI_E_INVALID_TYPE;

	WC_MAPI(lpMAPIProp->OpenProperty(
		ulPropTag,
		&IID_IMAPITable,
		fMapiUnicode,
		0,
		(LPUNKNOWN *)&lpTable));
	if (MAPI_E_INTERFACE_NOT_SUPPORTED == hRes)
	{
		hRes = S_OK;
		switch (PROP_ID(ulPropTag))
		{
		case PROP_ID(PR_MESSAGE_ATTACHMENTS):
			((LPMESSAGE)lpMAPIProp)->GetAttachmentTable(
				NULL,
				&lpTable);
			break;
		case PROP_ID(PR_MESSAGE_RECIPIENTS):
			((LPMESSAGE)lpMAPIProp)->GetRecipientTable(
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
				(LPMESSAGE)lpMAPIProp);
			break;
		case PROP_ID(PR_MESSAGE_RECIPIENTS):
			new CRecipientsDlg(
				lpHostDlg->GetParentWnd(),
				lpHostDlg->GetMapiObjects(),
				lpTable,
				(LPMESSAGE)lpMAPIProp);
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
	HRESULT			hRes = S_OK;
	LPEXCHANGEMODIFYTABLE	lpExchTbl = NULL;
	LPMAPITABLE		lpMAPITable = NULL;

	if (!lpMAPIProp || !lpHostDlg) return MAPI_E_INVALID_PARAMETER;

	CMapiObjects* lpMapiObjects = lpHostDlg->GetMapiObjects(); // do not release
	if (!lpMapiObjects) return MAPI_E_INVALID_PARAMETER;

	CParentWnd* lpParentWnd = lpHostDlg->GetParentWnd(); // do not release
	if (!lpParentWnd) return MAPI_E_INVALID_PARAMETER;

	// Open the table in an IExchangeModifyTable interface
	EC_MAPI(lpMAPIProp->OpenProperty(
		ulPropTag,
		(LPGUID)&IID_IExchangeModifyTable,
		0,
		MAPI_DEFERRED_ERRORS,
		(LPUNKNOWN*)&lpExchTbl));

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
} // DisplayExchangeTable

_Check_return_ bool bShouldCancel(_In_opt_ CWnd* cWnd, HRESULT hResPrev)
{
	bool bGotError = false;

	if (S_OK != hResPrev)
	{
		if (MAPI_E_USER_CANCEL != hResPrev && MAPI_E_CANCEL != hResPrev)
		{
			bGotError = true;
		}

		HRESULT hRes = S_OK;

		CEditor Cancel(
			cWnd,
			ID_PRODUCTNAME,
			IDS_CANCELPROMPT,
			bGotError ? 1 : 0,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		if (bGotError)
		{
			CString szPrevErr;
			szPrevErr.FormatMessage(IDS_PREVIOUSCALL, ErrorNameFromErrorCode(hResPrev), hResPrev);
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

void DisplayMailboxTable(_In_ CParentWnd*	lpParent,
	_In_ CMapiObjects* lpMapiObjects)
{
	if (!lpParent || !lpMapiObjects) return;
	HRESULT			hRes = S_OK;
	LPMDB			lpPrivateMDB = NULL;
	LPMDB			lpMDB = lpMapiObjects->GetMDB(); // do not release
	LPMAPISESSION	lpMAPISession = lpMapiObjects->GetSession(); // do not release

	// try the 'current' MDB first
	if (!StoreSupportsManageStore(lpMDB))
	{
		// if that MDB doesn't support manage store, try to get one that does
		EC_H(OpenMessageStoreGUID(lpMAPISession, pbExchangeProviderPrimaryUserGuid, &lpPrivateMDB));
		lpMDB = lpPrivateMDB;
	}

	if (lpMDB && StoreSupportsManageStore(lpMDB))
	{
		LPMAPITABLE	lpMailboxTable = NULL;
		LPTSTR		szServerName = NULL;
		EC_H(GetServerName(lpMAPISession, &szServerName));

		CEditor MyData(
			(CWnd*)lpParent,
			IDS_DISPLAYMAILBOXTABLE,
			IDS_SERVERNAMEPROMPT,
			4,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, CreateSingleLinePane(IDS_SERVERNAME, szServerName, false));
		MyData.InitPane(1, CreateSingleLinePane(IDS_OFFSET, NULL, false));
		MyData.SetHex(1, 0);
		MyData.InitPane(2, CreateSingleLinePane(IDS_MAILBOXGUID, NULL, false));
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
			LPTSTR	szServerDN = NULL;

			EC_H(BuildServerDN(
				MyData.GetString(0),
				_T(""),
				&szServerDN));
			if (szServerDN)
			{
				LPMDB lpOldMDB = NULL;

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
					bool bHaveGUID = false;

					LPTSTR pszGUID = NULL;
					pszGUID = MyData.GetString(2);

					if (_T('\0') != pszGUID[0])
					{
						bHaveGUID = true;

						WC_H(StringToGUID(MyData.GetStringW(2), &MyGUID));
						if (FAILED(hRes))
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
						MyData.GetString(0),
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
			MAPIFreeBuffer(szServerDN);
		}
		MAPIFreeBuffer(szServerName);
	}
	if (lpPrivateMDB) lpPrivateMDB->Release();
}

void DisplayPublicFolderTable(_In_ CParentWnd* lpParent,
	_In_ CMapiObjects* lpMapiObjects)
{
	if (!lpParent || !lpMapiObjects) return;
	HRESULT			hRes = S_OK;
	LPMDB			lpPrivateMDB = NULL;
	LPMDB			lpMDB = lpMapiObjects->GetMDB(); // do not release
	LPMAPISESSION	lpMAPISession = lpMapiObjects->GetSession(); // do not release

	// try the 'current' MDB first
	if (!StoreSupportsManageStore(lpMDB))
	{
		// if that MDB doesn't support manage store, try to get one that does
		EC_H(OpenMessageStoreGUID(lpMAPISession, pbExchangeProviderPrimaryUserGuid, &lpPrivateMDB));
		lpMDB = lpPrivateMDB;
	}

	if (lpMDB && StoreSupportsManageStore(lpMDB))
	{
		LPMAPITABLE	lpPFTable = NULL;
		LPTSTR		szServerName = NULL;
		EC_H(GetServerName(lpMAPISession, &szServerName));

		CEditor MyData(
			(CWnd*)lpParent,
			IDS_DISPLAYPFTABLE,
			IDS_DISPLAYPFTABLEPROMPT,
			5,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, CreateSingleLinePane(IDS_SERVERNAME, szServerName, false));
		MyData.InitPane(1, CreateSingleLinePane(IDS_OFFSET, NULL, false));
		MyData.SetHex(1, 0);
		MyData.InitPane(2, CreateSingleLinePane(IDS_FLAGS, NULL, false));
		MyData.SetHex(2, MDB_IPM);
		MyData.InitPane(3, CreateSingleLinePane(IDS_PUBLICFOLDERGUID, NULL, false));
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
			LPTSTR	szServerDN = NULL;

			EC_H(BuildServerDN(
				MyData.GetString(0),
				_T(""),
				&szServerDN));
			if (szServerDN)
			{
				LPMDB lpOldMDB = NULL;

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
					bool bHaveGUID = false;

					LPTSTR pszGUID = NULL;
					pszGUID = MyData.GetString(3);

					if (_T('\0') != pszGUID[0])
					{
						bHaveGUID = true;

						WC_H(StringToGUID(MyData.GetStringW(3), &MyGUID));
						if (FAILED(hRes))
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
						MyData.GetString(0),
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
			MAPIFreeBuffer(szServerDN);
		}
		MAPIFreeBuffer(szServerName);
	}
	if (lpPrivateMDB) lpPrivateMDB->Release();
}

void ResolveMessageClass(_In_ CMapiObjects* lpMapiObjects, _In_opt_ LPMAPIFOLDER lpMAPIFolder, _Out_ LPMAPIFORMINFO* lppMAPIFormInfo)
{
	HRESULT	hRes = S_OK;
	LPMAPIFORMMGR lpMAPIFormMgr = NULL;
	if (!lpMapiObjects || !lppMAPIFormInfo) return;

	*lppMAPIFormInfo = NULL;

	LPMAPISESSION lpMAPISession = lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	EC_MAPI(MAPIOpenFormMgr(lpMAPISession, &lpMAPIFormMgr));
	if (lpMAPIFormMgr)
	{
		DebugPrint(DBGForms, L"OnResolveMessageClass: resolving message class\n");
		CEditor MyData(
			NULL,
			IDS_RESOLVECLASS,
			IDS_RESOLVECLASSPROMPT,
			(ULONG)2,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, CreateSingleLinePane(IDS_CLASS, NULL, false));
		MyData.InitPane(1, CreateSingleLinePane(IDS_FLAGS, NULL, false));

		WC_H(MyData.DisplayDialog());
		if (S_OK == hRes)
		{
			LPSTR szClass = MyData.GetStringA(0); // ResolveMessageClass requires an ANSI string
			ULONG ulFlags = MyData.GetHex(1);
			if (szClass)
			{
				LPMAPIFORMINFO lpMAPIFormInfo = NULL;
				DebugPrint(DBGForms, L"OnResolveMessageClass: Calling ResolveMessageClass(\"%hs\",0x%08X)\n", szClass, ulFlags); // STRING_OK
				EC_MAPI(lpMAPIFormMgr->ResolveMessageClass(szClass, ulFlags, lpMAPIFolder, &lpMAPIFormInfo));
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
	HRESULT hRes = S_OK;
	LPMAPIFORMMGR lpMAPIFormMgr = NULL;

	if (!lpMapiObjects || !lppMAPIFormInfo) return;

	*lppMAPIFormInfo = NULL;

	LPMAPISESSION lpMAPISession = lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	EC_MAPI(MAPIOpenFormMgr(lpMAPISession, &lpMAPIFormMgr));

	if (lpMAPIFormMgr)
	{
		LPMAPIFORMINFO lpMAPIFormInfo = NULL;
		// Apparently, SelectForm doesn't support unicode
		// CString doesn't provide a way to extract just ANSI strings, so we do this manually
		CHAR szTitle[256];
		int iRet = NULL;
		EC_D(iRet, LoadStringA(GetModuleHandle(NULL),
			IDS_SELECTFORMPROPS,
			szTitle,
			_countof(szTitle)));
#pragma warning(push)
#pragma warning(disable:4616)
#pragma warning(disable:6276)
		EC_H_CANCEL(lpMAPIFormMgr->SelectForm(
			(ULONG_PTR)hWnd,
			0, // fMapiUnicode,
			(LPCTSTR)szTitle,
			lpMAPIFolder,
			&lpMAPIFormInfo));
#pragma warning(pop)

		if (lpMAPIFormInfo)
		{
			DebugPrintFormInfo(DBGForms, lpMAPIFormInfo);
			*lppMAPIFormInfo = lpMAPIFormInfo;
		}
		lpMAPIFormMgr->Release();
	}
}