#pragma once
// DumpStore.h : Processes a store/folder to dump to disk

#include "MAPIProcessor.h"

struct MessageData
{
	WCHAR	szFilePath[MAX_PATH]; // Holds file name prepended with path
	FILE* fMessageProps;
	ULONG ulCurAttNum;
};
typedef MessageData* LPMESSAGEDATA;

class CDumpStore : public CMAPIProcessor
{
public:
	CDumpStore();
	virtual ~CDumpStore();

	void InitMessagePath(_In_z_ LPCWSTR szMessageFileName);
	void InitFolderPathRoot(_In_z_ LPCWSTR szFolderPathRoot);
	void InitMailboxTablePathRoot(_In_z_ LPCWSTR szMailboxTablePathRoot);
	void EnableMSG();
	void EnableList();
	void DisableStreamRetry();

private:
	// Worker functions (dump messages, scan for something, etc)
	virtual void BeginMailboxTableWork(_In_z_ LPCTSTR szExchangeServerName);
	virtual void DoMailboxTablePerRowWork(_In_ LPMDB lpMDB, _In_ LPSRow lpSRow, ULONG ulCurRow);
	virtual void EndMailboxTableWork();

	virtual void BeginStoreWork();
	virtual void EndStoreWork();

	virtual void BeginFolderWork();
	virtual void DoFolderPerHierarchyTableRowWork(_In_ LPSRow lpSRow);
	virtual void EndFolderWork();

	virtual void BeginContentsTableWork(ULONG ulFlags, ULONG ulCountRows);
	virtual bool DoContentsTablePerRowWork(_In_ LPSRow lpSRow, ULONG ulCurRow);
	virtual void EndContentsTableWork();

	virtual bool BeginMessageWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpParentMessageData, _Deref_out_opt_ LPVOID* lpData);
	virtual bool BeginRecipientWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData);
	virtual void DoMessagePerRecipientWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData, _In_ LPSRow lpSRow, ULONG ulCurRow);
	virtual void EndRecipientWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData);
	virtual bool BeginAttachmentWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData);
	virtual void DoMessagePerAttachmentWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData, _In_ LPSRow lpSRow, _In_ LPATTACH lpAttach, ULONG ulCurRow);
	virtual void EndAttachmentWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData);
	virtual void EndMessageWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData);

	WCHAR m_szMailboxTablePathRoot[MAX_PATH];
	WCHAR m_szFolderPathRoot[MAX_PATH];
	WCHAR m_szMessageFileName[MAX_PATH];
	LPWSTR m_szFolderPath; // Root plus offset

	FILE* m_fFolderProps;
	FILE* m_fFolderContents;
	FILE* m_fMailboxTable;

	bool m_bOutputMSG;
	bool m_bOutputList;
	bool m_bRetryStreamProps;
};
