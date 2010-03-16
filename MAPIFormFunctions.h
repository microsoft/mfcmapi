#pragma once
// MAPIFormFunctions.h : Stand alone MAPI Form functions

class CContentsTableListCtrl;

HRESULT CreateAndDisplayNewMailInFolder(
										LPMDB lpMDB,
										LPMAPISESSION lpMAPISession,
										CContentsTableListCtrl *lpContentsTableListCtrl,
										int iItem,
										LPCSTR szMessageClass,
										LPMAPIFOLDER lpFolder);
HRESULT OpenMessageModal(LPMAPIFOLDER lpParentFolder,
						 LPMAPISESSION lpMAPISession,
						 LPMDB lpMDB,
						 LPMESSAGE lpMessage);
HRESULT OpenMessageNonModal(
							HWND hwndParent,
							LPMDB lpMDB,
							LPMAPISESSION lpMAPISession,
							LPMAPIFOLDER lpSourceFolder,
							CContentsTableListCtrl *lpContentsTableListCtrl,
							int iItem,
							LPMESSAGE lpMessage,
							LONG lVerb,
							LPCRECT lpRect);