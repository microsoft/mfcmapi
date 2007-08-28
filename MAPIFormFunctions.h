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

HRESULT CallActivateNext(
						 LPMAPIFORMADVISESINK	lpFormAdviseSink,
						 LPCTSTR szClass,
						 ULONG ulStatus,
						 ULONG ulFlags,
						 LPPERSISTMESSAGE* lppPersist);
HRESULT	CreateAndDisplayNewMailInFolder(
										LPMDB lpMDB,
										LPMAPISESSION lpMAPISession,
										CContentsTableListCtrl *lpContentsTableListCtrl,
										int iItem,
										LPCTSTR szMessageClass,
										LPMAPIFOLDER lpFolder);
HRESULT LoadForm(
				 LPMAPIMESSAGESITE lpMessageSite,
				 LPMESSAGE lpMessage,
				 LPMAPIFOLDER lpFolder,
				 LPCTSTR szMessageClass,
				 ULONG ulMessageStatus,
				 ULONG ulMessageFlags,
				 LPMAPIFORM* lppForm);
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