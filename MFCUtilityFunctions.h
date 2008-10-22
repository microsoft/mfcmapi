#pragma once
// MFCUtilityFunctions.h : Common functions for MFC MAPI

class CBaseDialog;
class CParentWnd;
class CMapiObjects;

enum ObjectType
{
	otDefault,
	otAssocContents,
	otContents,
	otHierarchy,
	otACL,
	otStatus,
	otReceive,
	otRules,
	otStore,
	otStoreDeletedItems,
};

HRESULT DisplayObject(
					  LPMAPIPROP lpUnk,
					  ULONG ulObjType,
					  ObjectType Table,
					  CBaseDialog* lpHostDlg);

HRESULT DisplayExchangeTable(
							 LPMAPIPROP lpMAPIProp,
							 ULONG ulPropTag,
							 ObjectType tType,
							 CBaseDialog* lpHostDlg);

HRESULT DisplayTable(
					 LPMAPITABLE lpTable,
					 ObjectType tType,
					 CBaseDialog* lpHostDlg);

HRESULT DisplayTable(
					 LPMAPIPROP lpMAPIProp,
					 ULONG ulPropTag,
					 ObjectType tType,
					 CBaseDialog* lpHostDlg);

void UpdateMenuString(CWnd* cWnd, UINT uiMenuTag, UINT uidNewString);

void DisplayContextMenu(UINT uiClassMenu, UINT uiControlMenu, CWnd* pParent, int x, int y);

int GetEditHeight(HWND hwndEdit);
int GetTextHeight(HWND hwndEdit);

BOOL bShouldCancel(CWnd* cWnd, HRESULT hRes);

void DisplayMailboxTable(CParentWnd*	lpParent,
						 CMapiObjects* lpMapiObjects);
void DisplayPublicFolderTable(CParentWnd* lpParent,
							  CMapiObjects* lpMapiObjects);