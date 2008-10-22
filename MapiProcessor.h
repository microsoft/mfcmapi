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
	void InitSession(LPMAPISESSION lpSession);
	void InitMDB(LPMDB lpMDB);
	void InitFolder(LPMAPIFOLDER lpFolder);

	// Processing functions
	void ProcessMailboxTable(LPCTSTR szExchangeServerName);
	void ProcessStore();
	void ProcessFolders(BOOL bDoRegular, BOOL bDoAssociated, BOOL bDoDescent);
	void ProcessMessage(LPMESSAGE lpMessage, LPVOID lpParentMessageData);

protected:
	LPMAPISESSION	m_lpSession;
	LPMDB			m_lpMDB;
	LPMAPIFOLDER	m_lpFolder;
	LPTSTR			m_szFolderOffset; // Offset to the folder, including trailing slash

private:
	// Worker functions (dump messages, scan for something, etc)
	virtual void BeginMailboxTableWork(LPCTSTR szExchangeServerName);
	virtual void DoMailboxTablePerRowWork(LPMDB lpMDB, LPSRow lpSRow, ULONG ulCurRow);
	virtual void EndMailboxTableWork();

	virtual void BeginStoreWork();
	virtual void EndStoreWork();

	virtual void BeginProcessFoldersWork();
	virtual void DoProcessFoldersPerFolderWork();
	virtual void EndProcessFoldersWork();

	virtual void BeginFolderWork();
	virtual void DoFolderPerHierarchyTableRowWork(LPSRow lpSRow);
	virtual void EndFolderWork();

	virtual void BeginContentsTableWork(ULONG ulFlags, ULONG ulCountRows);
	virtual void DoContentsTablePerRowWork(LPSRow lpSRow, ULONG ulCurRow);
	virtual void EndContentsTableWork();

	// lpData is allocated and returned by BeginMessageWork
	// If used, it should be cleaned up EndMessageWork
	// This allows implementations of these functions to avoid global variables
	virtual void BeginMessageWork(LPMESSAGE lpMessage, LPVOID lpParentMessageData, LPVOID* lpData);
	virtual void BeginRecipientWork(LPMESSAGE lpMessage, LPVOID lpData);
	virtual void DoMessagePerRecipientWork(LPMESSAGE lpMessage, LPVOID lpData, LPSRow lpSRow, ULONG ulCurRow);
	virtual void EndRecipientWork(LPMESSAGE lpMessage, LPVOID lpData);
	virtual void BeginAttachmentWork(LPMESSAGE lpMessage, LPVOID lpData);
	virtual void DoMessagePerAttachmentWork(LPMESSAGE lpMessage, LPVOID lpData, LPSRow lpSRow, LPATTACH lpAttach, ULONG ulCurRow);
	virtual void EndAttachmentWork(LPMESSAGE lpMessage, LPVOID lpData);
	virtual void EndMessageWork(LPMESSAGE lpMessage, LPVOID lpData);

	void ProcessFolder(BOOL bDoRegular, BOOL bDoAssociated, BOOL bDoDescent);
	void ProcessContentsTable(ULONG ulFlags);
	void ProcessRecipients(LPMESSAGE lpMessage, LPVOID lpData);
	void ProcessAttachments(LPMESSAGE lpMessage, LPVOID lpData);

	// FolderList functions
	// Add a new node to the end of the folder list
	void AddFolderToFolderList(LPSBinary lpFolderEID, LPCTSTR szFolderOffsetPath);

	// Call OpenEntry on the first folder in the list, remove it from the list
	void OpenFirstFolderInList();

	// Clean up the list
	void FreeFolderList();

	// Folder list pointers
	LPFOLDERNODE m_lpListHead;
	LPFOLDERNODE m_lpListTail;
};
