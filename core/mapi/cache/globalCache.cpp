#include <core/stdafx.h>
#include <core/mapi/cache/globalCache.h>
#include <core/utility/output.h>
#include <core/utility/error.h>
#include <core/utility/clipboard.h>

namespace cache
{
	static std::wstring GCCLASS = L"CGlobalCache"; // STRING_OK

	CGlobalCache::CGlobalCache() noexcept { TRACE_CONSTRUCTOR(GCCLASS); }

	CGlobalCache::~CGlobalCache()
	{
		TRACE_DESTRUCTOR(GCCLASS);
		EmptyBuffer();
		MAPIUninitialize();
	}

	void CGlobalCache::MAPIInitialize(ULONG ulFlags)
	{
		if (!m_bMAPIInitialized)
		{
			MAPIINIT_0 mapiInit = {MAPI_INIT_VERSION, ulFlags};
			const auto hRes = WC_MAPI(::MAPIInitialize(&mapiInit));
			if (SUCCEEDED(hRes))
			{
				m_bMAPIInitialized = true;
			}
			else
			{
				error::ErrDialog(
					__FILE__, __LINE__, IDS_EDMAPIINITIALIZEFAILED, hRes, error::ErrorNameFromErrorCode(hRes).c_str());
			}
		}
	}

	void CGlobalCache::MAPIUninitialize() noexcept
	{
		if (m_bMAPIInitialized)
		{
			::MAPIUninitialize();
			m_bMAPIInitialized = false;
		}
	}

	_Check_return_ bool CGlobalCache::bMAPIInitialized() const noexcept { return m_bMAPIInitialized; }

	void CGlobalCache::EmptyBuffer() noexcept
	{
		if (m_lpAddressEntriesToCopy) MAPIFreeBuffer(m_lpAddressEntriesToCopy);
		if (m_lpMessagesToCopy) MAPIFreeBuffer(m_lpMessagesToCopy);
		m_attachmentsToCopy.clear();
		m_szProfileToCopy.clear();
		if (m_lpFolderToCopy) m_lpFolderToCopy->Release();
		if (m_lpSourceParent) m_lpSourceParent->Release();
		if (m_lpSourcePropObject) m_lpSourcePropObject->Release();

		m_lpAddressEntriesToCopy = nullptr;
		m_lpMessagesToCopy = nullptr;
		m_lpFolderToCopy = nullptr;
		m_propToCopy = {};
		m_sourcePropBag = {};
		m_lpSourceParent = nullptr;
		m_lpSourcePropObject = nullptr;
	}

	void CGlobalCache::SetABEntriesToCopy(_In_ LPENTRYLIST lpEBEntriesToCopy) noexcept
	{
		EmptyBuffer();
		m_lpAddressEntriesToCopy = lpEBEntriesToCopy;
	}

	_Check_return_ LPENTRYLIST CGlobalCache::GetABEntriesToCopy() const noexcept { return m_lpAddressEntriesToCopy; }

	void CGlobalCache::SetMessagesToCopy(_In_ LPENTRYLIST lpMessagesToCopy, _In_ LPMAPIFOLDER lpSourceParent) noexcept
	{
		EmptyBuffer();
		m_lpMessagesToCopy = lpMessagesToCopy;
		m_lpSourceParent = lpSourceParent;
		if (m_lpSourceParent) m_lpSourceParent->AddRef();
	}

	_Check_return_ LPENTRYLIST CGlobalCache::GetMessagesToCopy() const noexcept { return m_lpMessagesToCopy; }

	void CGlobalCache::SetFolderToCopy(_In_ LPMAPIFOLDER lpFolderToCopy, _In_ LPMAPIFOLDER lpSourceParent) noexcept
	{
		EmptyBuffer();
		m_lpFolderToCopy = lpFolderToCopy;
		if (m_lpFolderToCopy) m_lpFolderToCopy->AddRef();
		m_lpSourceParent = lpSourceParent;
		if (m_lpSourceParent) m_lpSourceParent->AddRef();
	}

	_Check_return_ LPMAPIFOLDER CGlobalCache::GetFolderToCopy() const noexcept
	{
		if (m_lpFolderToCopy) m_lpFolderToCopy->AddRef();
		return m_lpFolderToCopy;
	}

	_Check_return_ LPMAPIFOLDER CGlobalCache::GetSourceParentFolder() const noexcept
	{
		if (m_lpSourceParent) m_lpSourceParent->AddRef();
		return m_lpSourceParent;
	}

	void CGlobalCache::SetPropertyToCopy(
		std::shared_ptr<sortlistdata::propModelData> prop,
		std::shared_ptr<propertybag::IMAPIPropertyBag> propBag) noexcept
	{
		EmptyBuffer();
		m_propToCopy = prop;
		m_sourcePropBag = propBag;
		clipboard::CopyTo(m_propToCopy->ToString());
	}

	_Check_return_ std::shared_ptr<sortlistdata::propModelData> CGlobalCache::GetPropertyToCopy() const noexcept
	{
		return m_propToCopy;
	}

	_Check_return_ std::shared_ptr<propertybag::IMAPIPropertyBag> CGlobalCache::GetSourcePropBag() const noexcept
	{
		return m_sourcePropBag;
	}

	_Check_return_ LPMAPIPROP CGlobalCache::GetSourcePropObject() const noexcept
	{
		if (m_lpSourcePropObject) m_lpSourcePropObject->AddRef();
		return m_lpSourcePropObject;
	}

	void CGlobalCache::SetAttachmentsToCopy(_In_ LPMESSAGE lpMessage, _In_ const std::vector<ULONG>& attNumList)
	{
		EmptyBuffer();
		m_lpSourcePropObject = lpMessage;
		m_attachmentsToCopy = attNumList;
		if (m_lpSourcePropObject) m_lpSourcePropObject->AddRef();
	}

	_Check_return_ std::vector<ULONG> CGlobalCache::GetAttachmentsToCopy() const { return m_attachmentsToCopy; }

	void CGlobalCache::SetProfileToCopy(_In_ const std::wstring& szProfileName) { m_szProfileToCopy = szProfileName; }

	_Check_return_ std::wstring CGlobalCache::GetProfileToCopy() const { return m_szProfileToCopy; }

	_Check_return_ ULONG CGlobalCache::GetBufferStatus() const noexcept
	{
		auto ulStatus = BUFFER_EMPTY;
		if (m_lpMessagesToCopy) ulStatus |= BUFFER_MESSAGES;
		if (m_lpFolderToCopy) ulStatus |= BUFFER_FOLDER;
		if (m_lpSourceParent) ulStatus |= BUFFER_PARENTFOLDER;
		if (m_lpAddressEntriesToCopy) ulStatus |= BUFFER_ABENTRIES;
		if (m_propToCopy) ulStatus |= BUFFER_PROPTAG;
		if (m_lpSourcePropObject) ulStatus |= BUFFER_SOURCEPROPOBJ;
		if (!m_attachmentsToCopy.empty()) ulStatus |= BUFFER_ATTACHMENTS;
		if (!m_szProfileToCopy.empty()) ulStatus |= BUFFER_PROFILE;
		return ulStatus;
	}
} // namespace cache