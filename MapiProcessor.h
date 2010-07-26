#pragma once

/*
CMAPIProcessor

Generic class for processing MAPI objects
This class handles the nitty-gritty work of walking through contents and hierarchy tables
It calls worker functions to do the customizable work on each object
These worker functions are intended to be overridden by specialized classes inheriting from this class
*/

typedef struct _FolderNode	FAR * LPFOLDERNODE;

typedef struct _FolderNode
{
	LPSBinary		lpFolderEID;
	LPTSTR			szFolderOffsetPath;
	LPFOLDERNODE	lpNextFolder;
} FolderNode;

class CMAPIProcessor
{
public:
	CMAPIProcessor();
	virtual ~CMAPIProcessor();

	// Initialization
	void InitSession(_In_ LPMAPISESSION lpSession);
	void InitMDB(_In_ LPMDB lpMDB);
	void InitFolder(_In_ LPMAPIFOLDER lpFolder);

	// Processing functions
	void ProcessMailboxTable(_In_z_ LPCTSTR szExchangeServerName);
	void ProcessStore();
	void ProcessFolders(BOOL bDoRegular, BOOL bDoAssociated, BOOL bDoDescent);
	void ProcessMessage(_In_ LPMESSAGE lpMessage, _In_opt_ LPVOID lpParentMessageData);

protected:
	LPMAPISESSION	m_lpSession;
	LPMDB			m_lpMDB;
	LPMAPIFOLDER	m_lpFolder;
	LPTSTR			m_szFolderOffset; // Offset to the folder, including trailing slash

private:
	// Worker functions (dump messages, scan for something, etc)
	virtual void BeginMailboxTableWork(_In_z_ LPCTSTR szExchangeServerName);
	virtual void DoMailboxTablePerRowWork(_In_ LPMDB lpMDB, _In_ LPSRow lpSRow, ULONG ulCurRow);
	virtual void EndMailboxTableWork();

	virtual void BeginStoreWork();
	virtual void EndStoreWork();

	virtual void BeginProcessFoldersWork();
	virtual void DoProcessFoldersPerFolderWork();
	virtual void EndProcessFoldersWork();

	virtual void BeginFolderWork();
	virtual void DoFolderPerHierarchyTableRowWork(_In_ LPSRow lpSRow);
	virtual void EndFolderWork();

	virtual void BeginContentsTableWork(ULONG ulFlags, ULONG ulCountRows);
	virtual void DoContentsTablePerRowWork(_In_ LPSRow lpSRow, ULONG ulCurRow);
	virtual void EndContentsTableWork();

	// lpData is allocated and returned by BeginMessageWork
	// If used, it should be cleaned up EndMessageWork
	// This allows implementations of these functions to avoid global variables
	virtual void BeginMessageWork(_In_ LPMESSAGE lpMessage, _In_opt_ LPVOID lpParentMessageData, _Deref_out_ LPVOID* lpData);
	virtual void BeginRecipientWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData);
	virtual void DoMessagePerRecipientWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData, _In_ LPSRow lpSRow, ULONG ulCurRow);
	virtual void EndRecipientWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData);
	virtual void BeginAttachmentWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData);
	virtual void DoMessagePerAttachmentWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData, _In_ LPSRow lpSRow, _In_ LPATTACH lpAttach, ULONG ulCurRow);
	virtual void EndAttachmentWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData);
	virtual void EndMessageWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData);

	void ProcessFolder(BOOL bDoRegular, BOOL bDoAssociated, BOOL bDoDescent);
	void ProcessContentsTable(ULONG ulFlags);
	void ProcessRecipients(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData);
	void ProcessAttachments(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData);

	// FolderList functions
	// Add a new node to the end of the folder list
	void AddFolderToFolderList(_In_opt_ LPSBinary lpFolderEID, _In_z_ LPCTSTR szFolderOffsetPath);

	// Call OpenEntry on the first folder in the list, remove it from the list
	void OpenFirstFolderInList();

	// Clean up the list
	void FreeFolderList();

	// Folder list pointers
	LPFOLDERNODE m_lpListHead;
	LPFOLDERNODE m_lpListTail;
};
