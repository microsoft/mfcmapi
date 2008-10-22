#pragma once
// DumpStore.h : Processes a store/folder to dump to disk

#include "MAPIProcessor.h"

typedef struct _MessageData	FAR * LPMESSAGEDATA;

typedef struct _MessageData
{
	TCHAR	szFilePath[MAX_PATH]; // Holds file name prepended with path
	FILE* fMessageProps;
	ULONG ulCurAttNum;
} MessageData;


class CDumpStore : public CMAPIProcessor
{
public:
	CDumpStore();
	virtual ~CDumpStore();

	void InitMessagePath(LPCTSTR szMessageFileName);
	void InitFolderPathRoot(LPCTSTR szFolderPathRoot);
	void InitMailboxTablePathRoot(LPCTSTR szMailboxTablePathRoot);

private:
	// Worker functions (dump messages, scan for something, etc)
	virtual void BeginMailboxTableWork(LPCTSTR szExchangeServerName);
	virtual void DoMailboxTablePerRowWork(LPMDB lpMDB, LPSRow lpSRow, ULONG ulCurRow);
	virtual void EndMailboxTableWork();

	virtual void BeginStoreWork();
	virtual void EndStoreWork();

	virtual void BeginFolderWork();
	virtual void DoFolderPerHierarchyTableRowWork(LPSRow lpSRow);
	virtual void EndFolderWork();

	virtual void BeginContentsTableWork(ULONG ulFlags, ULONG ulCountRows);
	virtual void DoContentsTablePerRowWork(LPSRow lpSRow, ULONG ulCurRow);
	virtual void EndContentsTableWork();

	virtual void BeginMessageWork(LPMESSAGE lpMessage, LPVOID lpParentMessageData, LPVOID* lpData);
	virtual void BeginRecipientWork(LPMESSAGE lpMessage, LPVOID lpData);
	virtual void DoMessagePerRecipientWork(LPMESSAGE lpMessage, LPVOID lpData, LPSRow lpSRow, ULONG ulCurRow);
	virtual void EndRecipientWork(LPMESSAGE lpMessage, LPVOID lpData);
	virtual void BeginAttachmentWork(LPMESSAGE lpMessage, LPVOID lpData);
	virtual void DoMessagePerAttachmentWork(LPMESSAGE lpMessage, LPVOID lpData, LPSRow lpSRow, LPATTACH lpAttach, ULONG ulCurRow);
	virtual void EndAttachmentWork(LPMESSAGE lpMessage, LPVOID lpData);
	virtual void EndMessageWork(LPMESSAGE lpMessage, LPVOID lpData);

	TCHAR m_szMailboxTablePathRoot[MAX_PATH];
	TCHAR m_szFolderPathRoot[MAX_PATH];
	TCHAR m_szMessageFileName[MAX_PATH];
	LPTSTR m_szFolderPath; // Root plus offset

	FILE* m_fFolderProps;
	FILE* m_fFolderContents;
	FILE* m_fMailboxTable;
};
