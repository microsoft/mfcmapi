#pragma once
// Processes a store/folder to dump to disk
#include <core/mapi/processor/mapiProcessor.h>

namespace mapi::processor
{
	void SaveFolderContentsToTXT(
		_In_ LPMDB lpMDB,
		_In_ LPMAPIFOLDER lpFolder,
		bool bRegular,
		bool bAssoc,
		bool bDescend,
		HWND hWnd);

	struct MessageData
	{
		std::wstring szFilePath; // Holds file name prepended with path
		FILE* fMessageProps{};
		ULONG ulCurAttNum{};
	};
	typedef MessageData* LPMESSAGEDATA;

	class dumpStore : public mapiProcessor
	{
	public:
		dumpStore() noexcept;
		~dumpStore();

		void InitMessagePath(_In_ const std::wstring& szMessageFileName);
		void InitFolderPathRoot(_In_ const std::wstring& szFolderPathRoot);
		void InitMailboxTablePathRoot(_In_ const std::wstring& szMailboxTablePathRoot);
		void InitProperties(const std::vector<std::wstring>& properties);
		void InitNamedProperties(const std::vector<std::wstring>& namedProperties);
		void EnableMSG() noexcept;
		void EnableList() noexcept;
		void DisableStreamRetry() noexcept;
		void DisableEmbeddedAttachments() noexcept;

	private:
		// Worker functions (dump messages, scan for something, etc)
		void BeginMailboxTableWork(_In_ const std::wstring& szExchangeServerName) override;
		void DoMailboxTablePerRowWork(_In_ LPMDB lpMDB, _In_ const _SRow* lpSRow, ULONG ulCurRow) override;
		void EndMailboxTableWork() override;

		void BeginStoreWork() noexcept override;
		void EndStoreWork() noexcept override;

		void BeginFolderWork() override;
		void DoFolderPerHierarchyTableRowWork(_In_ const _SRow* lpSRow) override;
		void EndFolderWork() override;

		void BeginContentsTableWork(ULONG ulFlags, ULONG ulCountRows) override;
		bool DoContentsTablePerRowWork(_In_ const _SRow* lpSRow, ULONG ulCurRow) override;
		void EndContentsTableWork() override;

		bool BeginMessageWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpParentMessageData, _Deref_out_opt_ LPVOID* lpData)
			override;
		bool BeginRecipientWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData) override;
		void DoMessagePerRecipientWork(
			_In_ LPMESSAGE lpMessage,
			_In_ LPVOID lpData,
			_In_ const _SRow* lpSRow,
			ULONG ulCurRow) override;
		void EndRecipientWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData) override;
		bool BeginAttachmentWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData) override;
		void DoMessagePerAttachmentWork(
			_In_ LPMESSAGE lpMessage,
			_In_ LPVOID lpData,
			_In_ const _SRow* lpSRow,
			_In_ LPATTACH lpAttach,
			ULONG ulCurRow) override;
		void EndAttachmentWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData) override;
		void EndMessageWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData) override;

		void InitializeInterestingTagArray();
		bool MessageHasInterestingProperties(_In_ LPMESSAGE lpMessage);

		std::wstring m_szMailboxTablePathRoot;
		std::wstring m_szFolderPathRoot;
		std::wstring m_szMessageFileName;
		std::wstring m_szFolderPath; // Root plus offset

		FILE* m_fFolderProps;
		FILE* m_fFolderContents;
		FILE* m_fMailboxTable;

		bool m_bOutputMSG;
		bool m_bOutputList;
		bool m_bRetryStreamProps;
		bool m_bOutputAttachments;

		std::vector<std::wstring> m_properties;
		std::vector<std::wstring> m_namedProperties;
		LPSPropTagArray m_lpInterestingPropTags = nullptr;
	};
} // namespace mapi::processor