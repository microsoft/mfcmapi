// MAPIStoreFunctions.h : Stand alone MAPI Store functions

#pragma once

_Check_return_ HRESULT CallOpenMsgStore(
						 _In_ LPMAPISESSION	lpSession,
						 _In_ ULONG_PTR		ulUIParam,
						 _In_ LPSBinary		lpEID,
						 ULONG			ulFlags,
						 _Deref_out_ LPMDB*			lpMDB);
_Check_return_ HRESULT BuildServerDN(
					  _In_z_ LPCTSTR szServerName,
					  _In_z_ LPCTSTR szPost,
					  _Deref_out_z_ LPTSTR* lpszServerDN);
_Check_return_ HRESULT GetMailboxTable(
						_In_ LPMDB lpMDB,
						_In_z_ LPCTSTR szServerName,
						ULONG ulOffset,
						_Deref_out_opt_ LPMAPITABLE* lpMailboxTable);
_Check_return_ HRESULT GetMailboxTable1(
						 _In_ LPMDB lpMDB,
						 _In_z_ LPCTSTR szServerDN,
						 ULONG ulFlags,
						 _Deref_out_opt_ LPMAPITABLE* lpMailboxTable);
_Check_return_ HRESULT GetMailboxTable3(
						 _In_ LPMDB lpMDB,
						 _In_z_ LPCTSTR szServerDN,
						 ULONG ulOffset,
						 ULONG ulFlags,
						 _Deref_out_opt_ LPMAPITABLE* lpMailboxTable);
_Check_return_ HRESULT GetMailboxTable5(
						 _In_ LPMDB lpMDB,
						 _In_z_ LPCTSTR szServerDN,
						 ULONG ulOffset,
						 ULONG ulFlags,
						 _In_opt_ LPGUID lpGuidMDB,
						 _Deref_out_opt_ LPMAPITABLE* lpMailboxTable);
_Check_return_ HRESULT GetPublicFolderTable1(
							  _In_ LPMDB lpMDB,
							  _In_z_ LPCTSTR szServerDN,
							  ULONG ulFlags,
							  _Deref_out_opt_ LPMAPITABLE* lpPFTable);
_Check_return_ HRESULT GetPublicFolderTable4(
							  _In_ LPMDB lpMDB,
							  _In_z_ LPCTSTR szServerDN,
							  ULONG ulOffset,
							  ULONG ulFlags,
							  _Deref_out_opt_ LPMAPITABLE* lpPFTable);
_Check_return_ HRESULT GetPublicFolderTable5(
							  _In_ LPMDB lpMDB,
							  _In_z_ LPCTSTR szServerDN,
							  ULONG ulOffset,
							  ULONG ulFlags,
							  _In_opt_ LPGUID lpGuidMDB,
							  _Deref_out_opt_ LPMAPITABLE* lpPFTable);
_Check_return_ HRESULT GetServerName(_In_ LPMAPISESSION lpSession, _Deref_out_opt_z_ LPTSTR* szServerName);
_Check_return_ HRESULT HrMailboxLogon(
					   _In_ LPMAPISESSION		lplhSession,	// ptr to MAPI session handle
					   _In_ LPMDB				lpMDB,			// ptr to message store
					   _In_z_ LPCTSTR			lpszMsgStoreDN,	// ptr to message store DN
					   _In_opt_z_ LPCTSTR		lpszMailboxDN,  // ptr to mailbox DN
					   ULONG					ulFlags,		// desired flags for CreateStoreEntryID
					   _Deref_out_opt_ LPMDB*	lppMailboxMDB);	// ptr to mailbox message store ptr
_Check_return_ HRESULT OpenDefaultMessageStore(
								_In_ LPMAPISESSION lpMAPISession,
								_Deref_out_ LPMDB* lppDefaultMDB);
_Check_return_ HRESULT OpenOtherUsersMailbox(
							  _In_ LPMAPISESSION lpMAPISession,
							  _In_ LPMDB lpMDB,
							  _In_opt_z_ LPCTSTR szServerName,
							  _In_z_ LPCTSTR szMailboxDN,
							  ULONG ulFlags, // desired flags for CreateStoreEntryID
							  _Deref_out_opt_ LPMDB* lppOtherUserMDB);
_Check_return_ HRESULT OpenOtherUsersMailboxFromGal(
									 _In_ LPMAPISESSION lpMAPISession,
									 _In_ LPADRBOOK lpAddrBook,
									 _Deref_out_opt_ LPMDB* lppOtherUserMDB);
_Check_return_ HRESULT OpenMessageStoreGUID(
							 _In_ LPMAPISESSION lpMAPISession,
							 _In_z_ LPCSTR lpGUID,
							 _Deref_out_opt_ LPMDB* lppMDB);
_Check_return_ HRESULT OpenPublicMessageStore(
							   _In_ LPMAPISESSION lpMAPISession,
							   ULONG ulFlags, // Flags for CreateStoreEntryID
							   _Deref_out_opt_ LPMDB* lppPublicMDB);
_Check_return_ HRESULT OpenStoreFromMAPIProp(_In_ LPMAPISESSION lpMAPISession, _In_ LPMAPIPROP lpMAPIProp, _Deref_out_ LPMDB* lpMDB);

_Check_return_ bool StoreSupportsManageStore(_In_ LPMDB lpMDB);

_Check_return_ HRESULT HrUnWrapMDB(_In_ LPMDB lpMDBIn, _Deref_out_ LPMDB* lppMDBOut);