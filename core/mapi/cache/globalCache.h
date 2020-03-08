#pragma once

namespace cache
{
// For GetBufferStatus
#define BUFFER_EMPTY ((ULONG) 0x00000000)
#define BUFFER_MESSAGES ((ULONG) 0x00000001)
#define BUFFER_FOLDER ((ULONG) 0x00000002)
#define BUFFER_PARENTFOLDER ((ULONG) 0x00000004)
#define BUFFER_ABENTRIES ((ULONG) 0x00000008)
#define BUFFER_PROPTAG ((ULONG) 0x00000010)
#define BUFFER_SOURCEPROPOBJ ((ULONG) 0x00000020)
#define BUFFER_ATTACHMENTS ((ULONG) 0x00000040)
#define BUFFER_PROFILE ((ULONG) 0x00000080)

	// A single instance cache of objects available to all
	class CGlobalCache
	{
	public:
		static CGlobalCache& getInstance() noexcept
		{
			static CGlobalCache instance;
			return instance;
		}

	private:
		CGlobalCache() noexcept;
		virtual ~CGlobalCache();

	public:
		CGlobalCache(CGlobalCache const&) = delete;
		void operator=(CGlobalCache const&) = delete;

		void MAPIInitialize(ULONG ulFlags);
		void MAPIUninitialize() noexcept;
		_Check_return_ bool bMAPIInitialized() const noexcept;

		void SetABEntriesToCopy(_In_ LPENTRYLIST lpEBEntriesToCopy) noexcept;
		_Check_return_ LPENTRYLIST GetABEntriesToCopy() const noexcept;

		void SetMessagesToCopy(_In_ LPENTRYLIST lpMessagesToCopy, _In_ LPMAPIFOLDER lpSourceParent) noexcept;
		_Check_return_ LPENTRYLIST GetMessagesToCopy() const noexcept;

		void SetFolderToCopy(_In_ LPMAPIFOLDER lpFolderToCopy, _In_ LPMAPIFOLDER lpSourceParent) noexcept;
		_Check_return_ LPMAPIFOLDER GetFolderToCopy() const noexcept;

		void SetPropertyToCopy(ULONG ulPropTag, _In_ LPMAPIPROP lpSourcePropObject) noexcept;
		_Check_return_ ULONG GetPropertyToCopy() const noexcept;
		_Check_return_ LPMAPIPROP GetSourcePropObject() const noexcept;

		void SetAttachmentsToCopy(_In_ LPMESSAGE lpMessage, _In_ const std::vector<ULONG>& attNumList);
		_Check_return_ std::vector<ULONG> GetAttachmentsToCopy() const;

		void SetProfileToCopy(_In_ const std::wstring& szProfileName);
		_Check_return_ std::wstring GetProfileToCopy() const;

		_Check_return_ LPMAPIFOLDER GetSourceParentFolder() const noexcept;

		_Check_return_ ULONG GetBufferStatus() const noexcept;

	private:
		void EmptyBuffer() noexcept;

		LPENTRYLIST m_lpAddressEntriesToCopy;
		LPENTRYLIST m_lpMessagesToCopy;
		LPMAPIFOLDER m_lpFolderToCopy;
		ULONG m_ulPropTagToCopy;
		std::vector<ULONG> m_attachmentsToCopy;
		std::wstring m_szProfileToCopy;
		LPMAPIFOLDER m_lpSourceParent;
		LPMAPIPROP m_lpSourcePropObject;
		bool m_bMAPIInitialized;
	};
} // namespace cache