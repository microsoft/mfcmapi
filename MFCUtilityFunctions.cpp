// MFCUtilityFunctions.h : Common functions for MFC MAPI

#include "stdafx.h"
#include "Error.h"

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
#include "InterpretProp2.h"
#include "AclDlg.h"
#include "RulesDlg.h"
#include "MailboxTableDlg.h"
#include "PublicFolderTableDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

HRESULT DisplayObject(
					  LPMAPIPROP lpUnk,
					  ULONG ulObjType,
					  ObjectType tType,
					  CBaseDialog* lpHostDlg)
{
	HRESULT			hRes = S_OK;
	__mfcmapiAssociatedContentsEnum	bAssocContents;

	if (!lpHostDlg || !lpHostDlg->m_lpMapiObjects || !lpUnk) return MAPI_E_INVALID_PARAMETER;

	//If we weren't passed an object type, go get one - careful! Some objects lie!
	if (!ulObjType)
	{
		ulObjType = GetMAPIObjectType((LPMAPIPROP)lpUnk);
	}

	bAssocContents = (otAssocContents == tType)?mfcmapiSHOW_ASSOC_CONTENTS:mfcmapiSHOW_NORMAL_CONTENTS;

	lpHostDlg->OnUpdateSingleMAPIPropListCtrl(lpUnk, NULL);

	LPTSTR szFlags = NULL;
	EC_H(InterpretFlags(PROP_ID(PR_OBJECT_TYPE), ulObjType, &szFlags));
	DebugPrint(DBGGeneric,_T("DisplayObject asked to display 0x%08X, with ObjectType of 0x%08X and MAPI type of 0x%08X = %s\n"),
		lpUnk,
		tType,
		ulObjType,
		szFlags);
	MAPIFreeBuffer(szFlags);
	szFlags = NULL;

	//call the dialog
	switch(ulObjType)
	{
//#define MAPI_STORE		((ULONG) 0x00000001)	/* Message Store */
	case MAPI_STORE:
		{
			LPMDB lpMDB = lpHostDlg->m_lpMapiObjects->GetMDB();//do not release
			if (lpMDB) lpMDB->AddRef();//hold on to this so that...
			lpHostDlg->m_lpMapiObjects->SetMDB((LPMDB) lpUnk);

			new CMsgStoreDlg(
				lpHostDlg->m_lpParent,
				lpHostDlg->m_lpMapiObjects,
				NULL,
				(otStoreDeletedItems == tType)?mfcmapiSHOW_DELETED_ITEMS:mfcmapiDO_NOT_SHOW_DELETED_ITEMS);

			//restore the old MDB
			lpHostDlg->m_lpMapiObjects->SetMDB(lpMDB);//...we can put it back
			if (lpMDB) lpMDB->Release();
			break;
		}
	case MAPI_FOLDER:
		{
			//There are two ways to display a folder...either the contents table or the hierarchy table.
			if (otHierarchy == tType)
			{
				LPMDB lpMDB = lpHostDlg->m_lpMapiObjects->GetMDB();//do not release
				if (lpMDB)
				{
					new CMsgStoreDlg(
						lpHostDlg->m_lpParent,
						lpHostDlg->m_lpMapiObjects,
						(LPMAPIFOLDER) lpUnk,
						mfcmapiDO_NOT_SHOW_DELETED_ITEMS);
				}
				else
				{
					//Since lpMDB was NULL, let's get a good MDB
					LPMAPISESSION lpMAPISession = lpHostDlg->m_lpMapiObjects->GetSession();//do not release
					if (lpMAPISession)
					{
						LPMDB lpNewMDB = NULL;
						EC_H(OpenStoreFromMAPIProp(lpMAPISession, (LPMAPIPROP)lpUnk,&lpNewMDB));

						if (lpNewMDB)
						{
							lpHostDlg->m_lpMapiObjects->SetMDB(lpNewMDB);

							new CMsgStoreDlg(
								lpHostDlg->m_lpParent,
								lpHostDlg->m_lpMapiObjects,
								(LPMAPIFOLDER) lpUnk,
								mfcmapiDO_NOT_SHOW_DELETED_ITEMS);

							//restore the old MDB
							lpHostDlg->m_lpMapiObjects->SetMDB(NULL);
							lpNewMDB->Release();
						}
					}
				}
			}
			else if (otContents == tType || otAssocContents == tType)
			{
				new CFolderDlg(
					lpHostDlg->m_lpParent,
					lpHostDlg->m_lpMapiObjects,
					(LPMAPIFOLDER) lpUnk,
					bAssocContents,
					mfcmapiDO_NOT_SHOW_DELETED_ITEMS);
			}
			break;
		}
//#define MAPI_ABCONT		((ULONG) 0x00000004)	/* Address Book Container */
	case MAPI_ABCONT:
		{
			new CAbDlg(
				lpHostDlg->m_lpParent,
				lpHostDlg->m_lpMapiObjects,
				(LPABCONT) lpUnk);
			break;
		}
//#define MAPI_DISTLIST	((ULONG) 0x00000008)	/* Distribution List Recipient */
	case MAPI_DISTLIST://A DistList is really an Address book anyways
		{
			new CAbDlg(
				lpHostDlg->m_lpParent,
				lpHostDlg->m_lpMapiObjects,
				(LPABCONT) lpUnk);
			break;
		}
		//#define MAPI_ADDRBOOK	((ULONG) 0x00000002)	/* Address Book */
		//#define MAPI_FOLDER	((ULONG) 0x00000003)	/* Folder */
		//#define MAPI_MESSAGE	((ULONG) 0x00000005)	/* Message */
		//#define MAPI_MAILUSER	((ULONG) 0x00000006)	/* Individual Recipient */
		//#define MAPI_ATTACH	((ULONG) 0x00000007)	/* Attachment */
		//#define MAPI_PROFSECT	((ULONG) 0x00000009)	/* Profile Section */
		//#define MAPI_STATUS	((ULONG) 0x0000000A)	/* Status Object */
		//#define MAPI_SESSION	((ULONG) 0x0000000B)	/* Session */
		//#define MAPI_FORMINFO	((ULONG) 0x0000000C)	/* Form Information */
	default:
		EC_H(InterpretFlags(PROP_ID(PR_OBJECT_TYPE), ulObjType, &szFlags));
		ErrDialog(__FILE__,__LINE__,
			IDS_EDVIEWERNOTIMPLEMENTED,
			ulObjType,
			szFlags);
		MAPIFreeBuffer(szFlags);
		szFlags = NULL;
		//Perhaps we could create a 'Single Object Property Display' window for this case?
		break;
	}

	return hRes;
}

HRESULT DisplayTable(
					 LPMAPITABLE lpTable,
					 ObjectType tType,
					 CBaseDialog* lpHostDlg)
{
	if (!lpHostDlg) return MAPI_E_INVALID_PARAMETER;

	DebugPrint(DBGGeneric,_T("DisplayTable asked to display 0x%X\n"),lpTable);

	switch (tType)
	{
	case otStatus:
		{
			if (!lpTable) return MAPI_E_INVALID_PARAMETER;
			new CContentsTableDlg(
				lpHostDlg->m_lpParent,
				lpHostDlg->m_lpMapiObjects,
				IDS_STATUSTABLE,
				mfcmapiCALL_CREATE_DIALOG,
				lpTable,
				(LPSPropTagArray) &sptSTATUSCols,
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
				lpHostDlg->m_lpParent,
				lpHostDlg->m_lpMapiObjects,
				IDS_RECEIVEFOLDERTABLE,
				mfcmapiCALL_CREATE_DIALOG,
				lpTable,
				(LPSPropTagArray) &sptRECEIVECols,
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
				lpHostDlg->m_lpParent,
				lpHostDlg->m_lpMapiObjects,
				IDS_HIERARCHYTABLE,
				mfcmapiCALL_CREATE_DIALOG,
				lpTable,
				(LPSPropTagArray) &sptHIERARCHYCols,
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
			if (otDefault != tType) ErrDialog(__FILE__,__LINE__,IDS_EDDISPLAYTABLE,tType);
			new CContentsTableDlg(
				lpHostDlg->m_lpParent,
				lpHostDlg->m_lpMapiObjects,
				IDS_CONTENTSTABLE,
				mfcmapiCALL_CREATE_DIALOG,
				lpTable,
				(LPSPropTagArray) &sptDEFCols,
				NUMDEFCOLUMNS,
				DEFColumns,
				NULL,
				MENU_CONTEXT_DEFAULT_TABLE);
			break;
		}
	}

	return S_OK;
}

HRESULT DisplayTable(
					 LPMAPIPROP lpMAPIProp,
					 ULONG ulPropTag,
					 ObjectType tType,
					 CBaseDialog* lpHostDlg)
{
	HRESULT		hRes = S_OK;
	LPMAPITABLE	lpTable = NULL;

	if (!lpHostDlg || !lpMAPIProp) return MAPI_E_INVALID_PARAMETER;
	if (PT_OBJECT != PROP_TYPE(ulPropTag)) return MAPI_E_INVALID_TYPE;

	EC_H(lpMAPIProp->OpenProperty(
		ulPropTag,
		&IID_IMAPITable,
		0,
		0,
		(LPUNKNOWN *) &lpTable));

	if (lpTable)
	{
		EC_H(DisplayTable(
			lpTable,
			tType,
			lpHostDlg));

		lpTable->Release();
	}

	return hRes;
}

HRESULT DisplayExchangeTable(
							 LPMAPIPROP lpMAPIProp,
							 ULONG ulPropTag,
							 ObjectType tType,
							 CBaseDialog* lpHostDlg)
{
	HRESULT			hRes = S_OK;
	LPEXCHANGEMODIFYTABLE	lpExchTbl = NULL;
	LPMAPITABLE		lpMAPITable = NULL;

	if (!lpMAPIProp || !lpHostDlg) return MAPI_E_INVALID_PARAMETER;

	// Open the table in an IExchangeModifyTable interface
	EC_H(lpMAPIProp->OpenProperty(
		ulPropTag,
		(LPGUID)&IID_IExchangeModifyTable,
		0,
		MAPI_DEFERRED_ERRORS,
		(LPUNKNOWN FAR *)&lpExchTbl));

	if (lpExchTbl)
	{
		switch (tType)
		{
		case otRules:
			{
				new CRulesDlg(
					lpHostDlg->m_lpParent,
					lpHostDlg->m_lpMapiObjects,
					lpExchTbl);
				break;
			}
		case otACL:
			{
				CEditor MyData(
					lpHostDlg,
					IDS_ACLTABLE,
					IDS_ACLTABLEPROMPT,
					1,
					CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

				MyData.InitCheck(0,IDS_FBRIGHTSVISIBLE, false, false);

				WC_H(MyData.DisplayDialog());
				if (S_OK == hRes)
				{
					new CAclDlg(
						lpHostDlg->m_lpParent,
						lpHostDlg->m_lpMapiObjects,
						lpExchTbl,
						MyData.GetCheck(0));
				}

				if (MAPI_E_USER_CANCEL == hRes)
					hRes = S_OK;  // don't propogate the error past here

				break;
			}
		default:
			{
				// Open a MAPI table on the Exchange table property. This table can be
				// read to determine what the Exchange table looks like.
				EC_H(lpExchTbl->GetTable(0, &lpMAPITable));

				if (lpMAPITable)
				{
					EC_H(DisplayTable(
						lpMAPITable,
						tType,
						lpHostDlg));
				}
				lpMAPITable->Release();
				break;
			}
		}
		lpExchTbl->Release();
	}
	return hRes;
}

BOOL UpdateMenuString(CWnd *cWnd, UINT uiMenuTag, UINT uidNewString)
{
	HRESULT hRes = S_OK;

	CString szNewString;
	szNewString.LoadString(uidNewString);

	DebugPrint(DBGMenu,_T("UpdateMenuString: Changing menu item 0x%X on window 0x%X to \"%s\"\n"),uiMenuTag,cWnd,szNewString);
	HMENU hMenu = ::GetMenu(cWnd->m_hWnd);

	if (!hMenu) return false;

	MENUITEMINFO MenuInfo = {0};

	ZeroMemory(&MenuInfo,sizeof(MENUITEMINFO));

	MenuInfo.cbSize = sizeof(MENUITEMINFO);
	MenuInfo.fMask = MIIM_STRING;
	MenuInfo.dwTypeData = (LPTSTR)(LPCTSTR)szNewString;

	EC_B(SetMenuItemInfo(
		hMenu,
		uiMenuTag,
		FALSE,
		&MenuInfo));
	return HRES_TO_BOOL(hRes);
}

BOOL MergeMenu(CMenu * pMenuDestination, const CMenu * pMenuAdd)
{
	HRESULT hRes = S_OK;
	int iMenuDestItemCount = pMenuDestination->GetMenuItemCount();
	int iMenuAddItemCount = pMenuAdd->GetMenuItemCount();

	if(iMenuAddItemCount == 0) return true;

	if(iMenuDestItemCount > 0) pMenuDestination->AppendMenu(MF_SEPARATOR);

	for(int iLoop = 0; iLoop < iMenuAddItemCount; iLoop++ )
	{
		CString sMenuAddString;
		pMenuAdd->GetMenuString(iLoop, sMenuAddString, MF_BYPOSITION);

		CMenu* pSubMenu = pMenuAdd->GetSubMenu(iLoop);

		if (!pSubMenu)
		{
			UINT nState = pMenuAdd->GetMenuState( iLoop, MF_BYPOSITION );
			UINT nItemID = pMenuAdd->GetMenuItemID( iLoop );

			EC_B(pMenuDestination->AppendMenu(nState, nItemID, sMenuAddString));
			if (FAILED(hRes)) return false;
			iMenuDestItemCount++;
		}
		else
		{
			int iInsertPosDefault = pMenuDestination->GetMenuItemCount();

			CMenu NewPopupMenu;
			EC_B(NewPopupMenu.CreatePopupMenu());
			if (FAILED(hRes)) return false;

			EC_B(MergeMenu(&NewPopupMenu, pSubMenu));
			if (FAILED(hRes)) return false;

			HMENU hNewMenu = NewPopupMenu.GetSafeHmenu();

			EC_B(pMenuDestination->InsertMenu(
				iInsertPosDefault,
				MF_BYPOSITION | MF_POPUP | MF_ENABLED,
				(UINT_PTR)hNewMenu,
				sMenuAddString));
			if (FAILED(hRes)) return false;
			iMenuDestItemCount++;

			NewPopupMenu.Detach();
		}
	}

	return HRES_TO_BOOL(hRes);
}

BOOL DisplayContextMenu(UINT uiClassMenu, UINT uiControlMenu, CWnd* pParent, int x, int y)
{
	HRESULT hRes = S_OK;
	CMenu pPopup;
	CMenu pContext;
	EC_B(pPopup.CreateMenu());
	EC_B(pContext.LoadMenu(uiClassMenu));


	EC_B(InsertMenu(
		pPopup.m_hMenu,
		0,
		MF_BYPOSITION | MF_POPUP,
		(UINT_PTR) pContext.m_hMenu,
		_T("")));

	CMenu* pRealPopup = pPopup.GetSubMenu(0);

	if (pRealPopup)
	{
		if (uiControlMenu)
		{
			CMenu pAppended;
			EC_B(pAppended.LoadMenu(uiControlMenu));
			EC_B(MergeMenu(pRealPopup,&pAppended));
			EC_B(pAppended.DestroyMenu());
		}

		if (IDR_MENU_PROPERTY_POPUP == uiClassMenu)
		{
			ExtendAddInMenu(pRealPopup->m_hMenu, MENU_CONTEXT_PROPERTY);
		}

		EC_B(pRealPopup->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON, x, y, pParent));
	}

	EC_B(pContext.DestroyMenu());
	EC_B(pPopup.DestroyMenu());
	return HRES_TO_BOOL(hRes);
}

int GetEditHeight(HWND hwndEdit)
{
	HRESULT		hRes = S_OK;
	HFONT		hSysFont = 0;
	HFONT		hOldFont = 0;
	HDC			hdc = 0;
	TEXTMETRIC	tmFont = {0};
	TEXTMETRIC	tmSys = {0};
	int			iHeight = 0;

	// Get the DC for the edit control.
	EC_D(hdc, GetDC(hwndEdit));

	// Get the metrics for the current font.
	EC_B(::GetTextMetrics(hdc, &tmFont));

	// Get the metrics for the system font.
	WC_D(hSysFont, (HFONT) GetStockObject(SYSTEM_FONT));
	hOldFont = (HFONT) SelectObject(hdc, hSysFont);
	EC_B(::GetTextMetrics(hdc, &tmSys));

	// Select the original font back into the DC and release the DC.
	SelectObject(hdc, hOldFont);
	DeleteObject(hSysFont);
	ReleaseDC(hwndEdit, hdc);

	// Calculate the new height for the edit control.
	iHeight =
		tmFont.tmHeight
//		+ tmFont.tmExternalLeading
//		+ (min(tmFont.tmHeight, tmSys.tmHeight)/2)
		+ 2 * GetSystemMetrics(SM_CYFIXEDFRAME)//Adjust for the edit border
		+ 2 * GetSystemMetrics(SM_CXEDGE);//Adjust for the edit border
	return iHeight;
}

int GetTextHeight(HWND hwndEdit)
{
	HRESULT		hRes = S_OK;
	HDC			hdc = 0;
	TEXTMETRIC	tmFont = {0};
	int			iHeight = 0;

	// Get the DC for the edit control.
	EC_D(hdc, GetDC(hwndEdit));

	// Get the metrics for the current font.
	EC_B(::GetTextMetrics(hdc, &tmFont));

	ReleaseDC(hwndEdit, hdc);

	// Calculate the new height for the edit control.
	iHeight = tmFont.tmHeight - tmFont.tmInternalLeading;
	return iHeight;
}

BOOL bShouldCancel(CWnd* cWnd, HRESULT hResPrev)
{
	BOOL bGotError = false;

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
			bGotError?1:0,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
		if (bGotError)
		{
			CString szPrevErr;
			szPrevErr.FormatMessage(IDS_PREVIOUSCALL,ErrorNameFromErrorCode(hResPrev),hResPrev);
			Cancel.InitSingleLineSz(0,IDS_ERROR,szPrevErr,true);
		}
		WC_H(Cancel.DisplayDialog());
		if (S_OK != hRes)
		{
			DebugPrint(DBGGeneric,_T("bShouldCancel: User asked to cancel\n"));
			return true;
		}
	}
	return false;
}

void DisplayMailboxTable(CParentWnd*	lpParent,
						 CMapiObjects* lpMapiObjects)
{
	if (!lpParent || !lpMapiObjects) return;
	HRESULT			hRes = S_OK;
	LPMDB			lpPrivateMDB = NULL;
	LPMDB			lpMDB = lpMapiObjects->GetMDB();//do not release
	LPMAPISESSION	lpMAPISession = lpMapiObjects->GetSession();//do not release

	//try the 'current' MDB first
	if (!StoreSupportsManageStore(lpMDB))
	{
		//if that MDB doesn't support manage store, try to get one that does
		EC_H(OpenMessageStoreGUID(lpMAPISession,pbExchangeProviderPrimaryUserGuid,&lpPrivateMDB));
		lpMDB = lpPrivateMDB;
	}

	if (lpMDB && StoreSupportsManageStore(lpMDB))
	{
		LPMAPITABLE	lpMailboxTable = NULL;
		LPTSTR		szServerName = NULL;
		EC_H(GetServerName(lpMAPISession, &szServerName));

		CEditor MyData(
			(CWnd*) lpParent,
			IDS_DISPLAYMAILBOXTABLE,
			IDS_SERVERNAMEPROMPT,
			4,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
		MyData.InitSingleLineSz(0,IDS_SERVERNAME,szServerName,false);
		MyData.InitSingleLine(1,IDS_OFFSET,NULL,false);
		MyData.SetHex(1,0);
		MyData.InitSingleLine(2,IDS_MAILBOXGUID,NULL,false);
		UINT uidDropDown[3] = {
			IDS_GETMBXINTERFACE1,
			IDS_GETMBXINTERFACE3,
			IDS_GETMBXINTERFACE5
		};
		MyData.InitDropDown(3,IDS_GETMBXINTERFACE,3,uidDropDown,true);
		WC_H(MyData.DisplayDialog());

		if (SUCCEEDED(hRes) && 0 != MyData.GetHex(1) && 0 == MyData.GetDropDown(3))
		{
			ErrDialog(__FILE__,__LINE__,IDS_EDOFFSETWITHWRONGINTERFACE);
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

				//if we got a new MDB, set it in lpMapiObjects
				if (lpPrivateMDB)
				{
					lpOldMDB = lpMapiObjects->GetMDB();//do not release
					if (lpOldMDB) lpOldMDB->AddRef();//hold on to this so that...
					//If we don't do this, we crash when destroying the Mailbox Table Window
					lpMapiObjects->SetMDB(lpMDB);
				}
				switch(MyData.GetDropDown(3))
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
						GUID MyGUID = {0};
						BOOL bHaveGUID = false;

						LPTSTR pszGUID = NULL;
						pszGUID = MyData.GetString(2);

						if (_T('\0') != pszGUID[0])
						{
							bHaveGUID = true;

							WC_H(StringToGUID(MyData.GetString(2),&MyGUID));
							if (FAILED(hRes))
							{
								ErrDialog(__FILE__,__LINE__,IDS_EDINVALIDGUID);
								break;
							}
						}

						EC_H(GetMailboxTable5(
							lpMDB,
							szServerDN,
							MyData.GetHex(1),
							fMapiUnicode,
							bHaveGUID?&MyGUID:0,
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
					ErrDialog(__FILE__,__LINE__,
						IDS_EDGETMAILBOXTABLEFAILED,
						_T("GetMailboxTable"),_T("GetMailboxTable"));// STRING_OK
				}
				if (lpMailboxTable) lpMailboxTable->Release();

				if (lpOldMDB)
				{
					lpMapiObjects->SetMDB(lpOldMDB);//...we can put it back
					if (lpOldMDB) lpOldMDB->Release();
				}
			}
			MAPIFreeBuffer(szServerDN);
		}
		MAPIFreeBuffer(szServerName);
	}
	if (lpPrivateMDB) lpPrivateMDB->Release();
}

void DisplayPublicFolderTable(CParentWnd* lpParent,
							  CMapiObjects* lpMapiObjects)
{
	if (!lpParent || !lpMapiObjects) return;
	HRESULT			hRes = S_OK;
	LPMDB			lpPrivateMDB = NULL;
	LPMDB			lpMDB = lpMapiObjects->GetMDB();//do not release
	LPMAPISESSION	lpMAPISession = lpMapiObjects->GetSession();//do not release

	//try the 'current' MDB first
	if (!StoreSupportsManageStore(lpMDB))
	{
		//if that MDB doesn't support manage store, try to get one that does
		EC_H(OpenMessageStoreGUID(lpMAPISession,pbExchangeProviderPrimaryUserGuid,&lpPrivateMDB));
		lpMDB = lpPrivateMDB;
	}

	if (lpMDB && StoreSupportsManageStore(lpMDB))
	{
		LPMAPITABLE	lpPFTable = NULL;
		LPTSTR		szServerName = NULL;
		EC_H(GetServerName(lpMAPISession, &szServerName));

		CEditor MyData(
			(CWnd*) lpParent,
			IDS_DISPLAYPFTABLE,
			IDS_DISPLAYPFTABLEPROMPT,
			5,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
		MyData.InitSingleLineSz(0,IDS_SERVERNAME,szServerName,false);
		MyData.InitSingleLine(1,IDS_OFFSET,NULL,false);
		MyData.SetHex(1,0);
		MyData.InitSingleLine(2,IDS_FLAGS,NULL,false);
		MyData.SetHex(2,MDB_IPM);
		MyData.InitSingleLine(3,IDS_PUBLICFOLDERGUID,NULL,false);
		UINT uidDropDown[3] = {
			IDS_GETPFINTERFACE1,
			IDS_GETPFINTERFACE4,
			IDS_GETPFINTERFACE5
		};
		MyData.InitDropDown(4,IDS_GETMBXINTERFACE,3,uidDropDown,true);
		WC_H(MyData.DisplayDialog());

		if (SUCCEEDED(hRes) && 0 != MyData.GetHex(1) && 0 == MyData.GetDropDown(4))
		{
			ErrDialog(__FILE__,__LINE__,IDS_EDOFFSETWITHWRONGINTERFACE);
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

				//if we got a new MDB, set it in lpMapiObjects
				if (lpPrivateMDB)
				{
					lpOldMDB = lpMapiObjects->GetMDB();//do not release
					if (lpOldMDB) lpOldMDB->AddRef();//hold on to this so that...
					//If we don't do this, we crash when destroying the Mailbox Table Window
					lpMapiObjects->SetMDB(lpMDB);
				}

				switch(MyData.GetDropDown(4))
				{
				case 0:
					EC_H(GetPublicFolderTable1(
						lpMDB,
						szServerDN,
						MyData.GetHex(2)| fMapiUnicode,
						&lpPFTable));
					break;
				case 1:
					EC_H(GetPublicFolderTable4(
						lpMDB,
						szServerDN,
						MyData.GetHex(1),
						MyData.GetHex(2)| fMapiUnicode,
						&lpPFTable));
					break;
				case 2:
					{
						GUID MyGUID = {0};
						BOOL bHaveGUID = false;

						LPTSTR pszGUID = NULL;
						pszGUID = MyData.GetString(2);

						if (_T('\0') != pszGUID[0])
						{
							bHaveGUID = true;

							WC_H(StringToGUID(MyData.GetString(2),&MyGUID));
							if (FAILED(hRes))
							{
								ErrDialog(__FILE__,__LINE__,IDS_EDINVALIDGUID);
								break;
							}
						}

						EC_H(GetPublicFolderTable5(
							lpMDB,
							szServerDN,
							MyData.GetHex(1),
							MyData.GetHex(2)| fMapiUnicode,
							bHaveGUID?&MyGUID:0,
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
					ErrDialog(__FILE__,__LINE__,
						IDS_EDGETMAILBOXTABLEFAILED,
						_T("GetPublicFolderTable"),_T("GetPublicFolderTable"));// STRING_OK
				}
				if (lpPFTable) lpPFTable->Release();

				if (lpOldMDB)
				{
					lpMapiObjects->SetMDB(lpOldMDB);//...we can put it back
					if (lpOldMDB) lpOldMDB->Release();
				}
			}
			MAPIFreeBuffer(szServerDN);
		}
		MAPIFreeBuffer(szServerName);
	}
	if (lpPrivateMDB) lpPrivateMDB->Release();
}
