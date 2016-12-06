#include "stdafx.h"
#include "GlobalCache.h"

static wstring GCCLASS = L"CGlobalCache"; // STRING_OK

CGlobalCache::CGlobalCache()
{
	TRACE_CONSTRUCTOR(GCCLASS);
	m_bMAPIInitialized = false;

	m_lpMessagesToCopy = nullptr;
	m_lpFolderToCopy = nullptr;
	m_lpAddressEntriesToCopy = nullptr;
	m_ulPropTagToCopy = 0;

	m_lpSourceParent = nullptr;
	m_lpSourcePropObject = nullptr;
}

CGlobalCache::~CGlobalCache()
{
	TRACE_DESTRUCTOR(GCCLASS);
	EmptyBuffer();
	MAPIUninitialize();
}

void CGlobalCache::MAPIInitialize(ULONG ulFlags)
{
	auto hRes = S_OK;
	if (!m_bMAPIInitialized)
	{
		MAPIINIT_0 mapiInit = { MAPI_INIT_VERSION, ulFlags };
		WC_MAPI(::MAPIInitialize(&mapiInit));
		if (SUCCEEDED(hRes))
		{
			m_bMAPIInitialized = true;
		}
		else
		{
			ErrDialog(__FILE__, __LINE__,
				IDS_EDMAPIINITIALIZEFAILED,
				hRes,
				ErrorNameFromErrorCode(hRes).c_str());
		}
	}
}

void CGlobalCache::MAPIUninitialize()
{
	if (m_bMAPIInitialized)
	{
		::MAPIUninitialize();
		m_bMAPIInitialized = false;
	}
}

_Check_return_ bool CGlobalCache::bMAPIInitialized() const
{
	return m_bMAPIInitialized;
}

void CGlobalCache::EmptyBuffer()
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
	m_ulPropTagToCopy = 0;
	m_lpSourceParent = nullptr;
	m_lpSourcePropObject = nullptr;
}

void CGlobalCache::SetABEntriesToCopy(_In_ LPENTRYLIST lpEBEntriesToCopy)
{
	EmptyBuffer();
	m_lpAddressEntriesToCopy = lpEBEntriesToCopy;
}

_Check_return_ LPENTRYLIST CGlobalCache::GetABEntriesToCopy() const
{
	return m_lpAddressEntriesToCopy;
}

void CGlobalCache::SetMessagesToCopy(_In_ LPENTRYLIST lpMessagesToCopy, _In_ LPMAPIFOLDER lpSourceParent)
{
	EmptyBuffer();
	m_lpMessagesToCopy = lpMessagesToCopy;
	m_lpSourceParent = lpSourceParent;
	if (m_lpSourceParent) m_lpSourceParent->AddRef();
}

_Check_return_ LPENTRYLIST CGlobalCache::GetMessagesToCopy() const
{
	return m_lpMessagesToCopy;
}

void CGlobalCache::SetFolderToCopy(_In_ LPMAPIFOLDER lpFolderToCopy, _In_ LPMAPIFOLDER lpSourceParent)
{
	EmptyBuffer();
	m_lpFolderToCopy = lpFolderToCopy;
	if (m_lpFolderToCopy) m_lpFolderToCopy->AddRef();
	m_lpSourceParent = lpSourceParent;
	if (m_lpSourceParent) m_lpSourceParent->AddRef();
}

_Check_return_ LPMAPIFOLDER CGlobalCache::GetFolderToCopy() const
{
	if (m_lpFolderToCopy) m_lpFolderToCopy->AddRef();
	return m_lpFolderToCopy;
}

_Check_return_ LPMAPIFOLDER CGlobalCache::GetSourceParentFolder() const
{
	if (m_lpSourceParent) m_lpSourceParent->AddRef();
	return m_lpSourceParent;
}

void CGlobalCache::SetPropertyToCopy(ULONG ulPropTag, _In_ LPMAPIPROP lpSourcePropObject)
{
	EmptyBuffer();
	m_ulPropTagToCopy = ulPropTag;
	m_lpSourcePropObject = lpSourcePropObject;
	if (m_lpSourcePropObject) m_lpSourcePropObject->AddRef();
}

_Check_return_ ULONG CGlobalCache::GetPropertyToCopy() const
{
	return m_ulPropTagToCopy;
}

_Check_return_ LPMAPIPROP CGlobalCache::GetSourcePropObject() const
{
	if (m_lpSourcePropObject) m_lpSourcePropObject->AddRef();
	return m_lpSourcePropObject;
}

void CGlobalCache::SetAttachmentsToCopy(_In_ LPMESSAGE lpMessage, _In_ const vector<ULONG>& attNumList)
{
	EmptyBuffer();
	m_lpSourcePropObject = lpMessage;
	m_attachmentsToCopy = attNumList;
	if (m_lpSourcePropObject) m_lpSourcePropObject->AddRef();
}

_Check_return_ vector<ULONG> CGlobalCache::GetAttachmentsToCopy() const
{
	return m_attachmentsToCopy;
}

void CGlobalCache::SetProfileToCopy(_In_ string szProfileName)
{
	m_szProfileToCopy = szProfileName;
}

_Check_return_ string CGlobalCache::GetProfileToCopy() const
{
	return m_szProfileToCopy;
}

_Check_return_ ULONG CGlobalCache::GetBufferStatus() const
{
	auto ulStatus = BUFFER_EMPTY;
	if (m_lpMessagesToCopy) ulStatus |= BUFFER_MESSAGES;
	if (m_lpFolderToCopy) ulStatus |= BUFFER_FOLDER;
	if (m_lpSourceParent) ulStatus |= BUFFER_PARENTFOLDER;
	if (m_lpAddressEntriesToCopy) ulStatus |= BUFFER_ABENTRIES;
	if (m_ulPropTagToCopy) ulStatus |= BUFFER_PROPTAG;
	if (m_lpSourcePropObject) ulStatus |= BUFFER_SOURCEPROPOBJ;
	if (!m_attachmentsToCopy.empty()) ulStatus |= BUFFER_ATTACHMENTS;
	if (!m_szProfileToCopy.empty()) ulStatus |= BUFFER_PROFILE;
	return ulStatus;
}