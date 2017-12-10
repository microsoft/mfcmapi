#pragma once
#include <deque>

/*
CMAPIProcessor

Generic class for processing MAPI objects
This class handles the nitty-gritty work of walking through contents and hierarchy tables
It calls worker functions to do the customizable work on each object
These worker functions are intended to be overridden by specialized classes inheriting from this class
*/

struct FolderNode
{
	LPSBinary lpFolderEID;
	wstring szFolderOffsetPath;
};

class CMAPIProcessor
{
public:
	CMAPIProcessor();
	virtual ~CMAPIProcessor();

	// Initialization
	void InitSession(_In_ LPMAPISESSION lpSession);
	void InitMDB(_In_ LPMDB lpMDB);
	void InitFolder(_In_ LPMAPIFOLDER lpFolder);
	void InitFolderContentsRestriction(_In_opt_ LPSRestriction lpRes);
	void InitMaxOutput(_In_ ULONG ulCount);
	void InitSortOrder(_In_ const LPSSortOrderSet lpSort);

	// Processing functions
	void ProcessMailboxTable(_In_ const wstring& szExchangeServerName);
	void ProcessStore();
	void ProcessFolders(bool bDoRegular, bool bDoAssociated, bool bDoDescent);
	void ProcessMessage(_In_ LPMESSAGE lpMessage, bool bHasAttach, _In_opt_ LPVOID lpParentMessageData);

protected:
	LPMAPISESSION m_lpSession;
	LPMDB m_lpMDB;
	LPMAPIFOLDER m_lpFolder;
	wstring m_szFolderOffset; // Offset to the folder, including trailing slash

private:
	// Worker functions (dump messages, scan for something, etc)
	virtual void BeginMailboxTableWork(_In_ const wstring& szExchangeServerName);
	virtual void DoMailboxTablePerRowWork(_In_ LPMDB lpMDB, _In_ const LPSRow lpSRow, ULONG ulCurRow);
	virtual void EndMailboxTableWork();

	virtual void BeginStoreWork();
	virtual void EndStoreWork();

	virtual bool ContinueProcessingFolders();
	virtual bool ShouldProcessContentsTable();
	virtual void BeginProcessFoldersWork();
	virtual void DoProcessFoldersPerFolderWork();
	virtual void EndProcessFoldersWork();

	virtual void BeginFolderWork();
	virtual void DoFolderPerHierarchyTableRowWork(_In_ const LPSRow lpSRow);
	virtual void EndFolderWork();

	virtual void BeginContentsTableWork(ULONG ulFlags, ULONG ulCountRows);
	// Called from ProcessContentsTable. If true is returned, ProcessContentsTable will continue
	// processing the message, calling OpenEntry and ProcessMessage. If false is returned,
	// ProcessContentsTable will stop processing the message.
	virtual bool DoContentsTablePerRowWork(_In_ const LPSRow lpSRow, ULONG ulCurRow);
	virtual void EndContentsTableWork();

	// lpData is allocated and returned by BeginMessageWork
	// If used, it should be cleaned up EndMessageWork
	// This allows implementations of these functions to avoid global variables
	// The Begin functions return true if work should continue
	// Do and End functions will only be called if Begin returned true
	// If BeginMessageWork returns false, we'll never call the recipient or attachment functions
	virtual bool BeginMessageWork(_In_ LPMESSAGE lpMessage, _In_opt_ LPVOID lpParentMessageData, _Deref_out_opt_ LPVOID* lpData);
	virtual bool BeginRecipientWork(_In_ LPMESSAGE lpMessage, _In_opt_ LPVOID lpData);
	virtual void DoMessagePerRecipientWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData, _In_ const LPSRow lpSRow, ULONG ulCurRow);
	virtual void EndRecipientWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData);
	virtual bool BeginAttachmentWork(_In_ LPMESSAGE lpMessage, _In_opt_ LPVOID lpData);
	virtual void DoMessagePerAttachmentWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData, _In_ const LPSRow lpSRow, _In_ LPATTACH lpAttach, ULONG ulCurRow);
	virtual void EndAttachmentWork(_In_ LPMESSAGE lpMessage, _In_opt_ LPVOID lpData);
	virtual void EndMessageWork(_In_ LPMESSAGE lpMessage, _In_opt_ LPVOID lpData);

	void ProcessFolder(bool bDoRegular, bool bDoAssociated, bool bDoDescent);
	void ProcessContentsTable(ULONG ulFlags);
	void ProcessRecipients(_In_ LPMESSAGE lpMessage, _In_opt_ LPVOID lpData);
	void ProcessAttachments(_In_ LPMESSAGE lpMessage, bool bHasAttach, _In_opt_ LPVOID lpData);

	// FolderList functions
	// Add a new node to the end of the folder list
	void AddFolderToFolderList(_In_opt_ const LPSBinary lpFolderEID, _In_ const wstring& szFolderOffsetPath);

	// Call OpenEntry on the first folder in the list, remove it from the list
	void OpenFirstFolderInList();

	// Clean up the list
	void FreeFolderList();

	// Folder list
	std::deque<FolderNode> m_List;

	LPSRestriction m_lpResFolderContents;
	LPSSortOrderSet m_lpSort;
	ULONG m_ulCount; // Limit on the number of messages processed per folder
};
