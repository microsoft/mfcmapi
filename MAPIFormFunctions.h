// MAPIFormFunctions.h : Stand alone MAPI Form functions

#pragma once

#include <MapiX.h>
#include <MapiUtil.h>
#include <MAPIform.h>
#include <MSPST.h>

#include <edkmdb.h>
#include <exchform.h>

//forward definitions
class CContentsTableListCtrl;

HRESULT	CreateAndDisplayNewMailInFolder(
										LPMDB lpMDB,
										LPMAPISESSION lpMAPISession,
										CContentsTableListCtrl *lpContentsTableListCtrl,
										int iItem,
										LPCSTR szMessageClass,
										LPMAPIFOLDER lpFolder);
HRESULT	OpenMessageModal(LPMAPIFOLDER lpParentFolder,
						 LPMAPISESSION lpMAPISession,
						 LPMDB lpMDB,
						 LPMESSAGE lpMessage);
HRESULT	OpenMessageNonModal(
							LPMDB lpMDB,
							LPMAPISESSION lpMAPISession,
							LPMAPIFOLDER lpSourceFolder,
							CContentsTableListCtrl *lpContentsTableListCtrl,
							int iItem,
							LPMESSAGE lpMessage,
							LONG lVerb,
							LPCRECT lpRect);