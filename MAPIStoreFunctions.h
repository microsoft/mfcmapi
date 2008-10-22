// MAPIStoreFunctions.h : Stand alone MAPI Store functions

#pragma once

HRESULT CallOpenMsgStore(
						 LPMAPISESSION	lpSession,
						 ULONG_PTR		ulUIParam,
						 LPSBinary		lpEID,
						 ULONG			ulFlags,
						 LPMDB*			lpMDB);
HRESULT BuildServerDN(
					  LPCTSTR szServerName,
					  LPCTSTR szPost,
					  LPTSTR* lpszServerDN);
HRESULT GetMailboxTable(
						LPMDB lpMDB,
						LPCTSTR szServerName,
						ULONG ulOffset,
						LPMAPITABLE* lpMailboxTable);
HRESULT GetMailboxTable1(
						LPMDB lpMDB,
						LPCTSTR szServerDN,
						ULONG ulFlags,
						LPMAPITABLE* lpMailboxTable);
HRESULT GetMailboxTable3(
						LPMDB lpMDB,
						LPCTSTR szServerDN,
						ULONG ulOffset,
						ULONG ulFlags,
						LPMAPITABLE* lpMailboxTable);
HRESULT GetMailboxTable5(
						LPMDB lpMDB,
						LPCTSTR szServerDN,
						ULONG ulOffset,
						ULONG ulFlags,
						LPGUID lpGuidMDB,
						LPMAPITABLE* lpMailboxTable);
HRESULT GetPublicFolderTable1(
						LPMDB lpMDB,
						LPCTSTR szServerDN,
						ULONG ulFlags,
						LPMAPITABLE* lpPFTable);
HRESULT GetPublicFolderTable4(
						LPMDB lpMDB,
						LPCTSTR szServerDN,
						ULONG ulOffset,
						ULONG ulFlags,
						LPMAPITABLE* lpPFTable);
HRESULT GetPublicFolderTable5(
						LPMDB lpMDB,
						LPCTSTR szServerDN,
						ULONG ulOffset,
						ULONG ulFlags,
						LPGUID lpGuidMDB,
						LPMAPITABLE* lpPFTable);
HRESULT GetServerName(LPMAPISESSION lpSession, LPTSTR* szServerName);
HRESULT HrMailboxLogon(
					   LPMAPISESSION	lplhSession,	// ptr to MAPI session handle
					   LPMDB			lpMDB,			// ptr to message store
					   LPCTSTR			lpszMsgStoreDN,	// ptr to message store DN
					   LPCTSTR			lpszMailboxDN,  // ptr to mailbox DN
					   ULONG			ulFlags,		// desired flags for CreateStoreEntryID
					   LPMDB*			lppMailboxMDB);	// ptr to mailbox message store ptr
HRESULT	OpenDefaultMessageStore(
								LPMAPISESSION lpMAPISession,
								LPMDB* lppDefaultMDB);
HRESULT OpenOtherUsersMailbox(
							  LPMAPISESSION	lpMAPISession,
							  LPMDB lpMDB,
							  LPCTSTR szServerName,
							  LPCTSTR szMailboxDN,
							  ULONG ulFlags, // desired flags for CreateStoreEntryID
							  LPMDB* lppOtherUserMDB);
HRESULT OpenOtherUsersMailboxFromGal(
									 LPMAPISESSION	lpMAPISession,
									 LPADRBOOK lpAddrBook,
									 LPMDB* lppOtherUserMDB);
HRESULT OpenMessageStoreGUID(
							  LPMAPISESSION	lpMAPISession,
							  LPCSTR lpGUID,
							  LPMDB* lppMDB);
HRESULT OpenPublicMessageStore(
							   LPMAPISESSION lpMAPISession,
							   ULONG ulFlags, // Flags for CreateStoreEntryID
							   LPMDB* lppPublicMDB);
HRESULT OpenStoreFromMAPIProp(LPMAPISESSION lpMAPISession, LPMAPIPROP lpMAPIProp, LPMDB* lpMDB);

BOOL StoreSupportsManageStore(LPMDB lpMDB);

HRESULT HrUnWrapMDB(LPMDB lpMDBIn, LPMDB* lppMDBOut);