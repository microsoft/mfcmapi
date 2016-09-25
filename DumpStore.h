#pragma once
// DumpStore.h : Processes a store/folder to dump to disk

#include "MAPIProcessor.h"

struct MessageData
{
	wstring szFilePath; // Holds file name prepended with path
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
	void DisableEmbeddedAttachments();

private:
	// Worker functions (dump messages, scan for something, etc)
	void BeginMailboxTableWork(_In_ wstring szExchangeServerName) override;
	void DoMailboxTablePerRowWork(_In_ LPMDB lpMDB, _In_ LPSRow lpSRow, ULONG ulCurRow) override;
	void EndMailboxTableWork() override;

	void BeginStoreWork() override;
	void EndStoreWork() override;

	void BeginFolderWork() override;
	void DoFolderPerHierarchyTableRowWork(_In_ LPSRow lpSRow) override;
	void EndFolderWork() override;

	void BeginContentsTableWork(ULONG ulFlags, ULONG ulCountRows) override;
	bool DoContentsTablePerRowWork(_In_ LPSRow lpSRow, ULONG ulCurRow) override;
	void EndContentsTableWork() override;

	bool BeginMessageWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpParentMessageData, _Deref_out_opt_ LPVOID* lpData) override;
	bool BeginRecipientWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData) override;
	void DoMessagePerRecipientWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData, _In_ LPSRow lpSRow, ULONG ulCurRow) override;
	void EndRecipientWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData) override;
	bool BeginAttachmentWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData) override;
	void DoMessagePerAttachmentWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData, _In_ LPSRow lpSRow, _In_ LPATTACH lpAttach, ULONG ulCurRow) override;
	void EndAttachmentWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData) override;
	void EndMessageWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData) override;

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
	bool m_bOutputAttachments;
};
