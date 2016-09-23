#pragma once

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
	CGlobalCache();
	virtual ~CGlobalCache();

	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	void MAPIInitialize(ULONG ulFlags);
	void MAPIUninitialize();
	_Check_return_ bool bMAPIInitialized() const;

	void SetABEntriesToCopy(_In_ LPENTRYLIST lpEBEntriesToCopy);
	_Check_return_ LPENTRYLIST GetABEntriesToCopy() const;

	void SetMessagesToCopy(_In_ LPENTRYLIST lpMessagesToCopy, _In_ LPMAPIFOLDER lpSourceParent);
	_Check_return_ LPENTRYLIST GetMessagesToCopy() const;

	void SetFolderToCopy(_In_ LPMAPIFOLDER lpFolderToCopy, _In_ LPMAPIFOLDER lpSourceParent);
	_Check_return_ LPMAPIFOLDER GetFolderToCopy() const;

	void SetPropertyToCopy(ULONG ulPropTag, _In_ LPMAPIPROP lpSourcePropObject);
	_Check_return_ ULONG GetPropertyToCopy() const;
	_Check_return_ LPMAPIPROP GetSourcePropObject() const;

	void SetAttachmentsToCopy(_In_ LPMESSAGE lpMessage, _In_ vector<ULONG> attNumList);
	_Check_return_ _In_ vector<ULONG> GetAttachmentsToCopy() const;

	void SetProfileToCopy(_In_ string szProfileName);
	_Check_return_ string GetProfileToCopy() const;

	_Check_return_ LPMAPIFOLDER GetSourceParentFolder() const;

	_Check_return_ ULONG GetBufferStatus() const;

private:
	void EmptyBuffer();

	LONG m_cRef;
	LPENTRYLIST m_lpAddressEntriesToCopy;
	LPENTRYLIST m_lpMessagesToCopy;
	LPMAPIFOLDER m_lpFolderToCopy;
	ULONG m_ulPropTagToCopy;
	vector<ULONG> m_attachmentsToCopy;
	string m_szProfileToCopy;
	LPMAPIFOLDER m_lpSourceParent;
	LPMAPIPROP m_lpSourcePropObject;
	bool m_bMAPIInitialized;
};