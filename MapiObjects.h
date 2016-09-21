#pragma once
// MapiObjects.h: interface for the CMapiObjects class.

class CGlobalCache;

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

class CMapiObjects
{
public:
	CMapiObjects(
		_In_opt_ CMapiObjects* OldMapiObjects);
	virtual ~CMapiObjects();

	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	_Check_return_ LPADRBOOK GetAddrBook(bool bForceOpen);
	_Check_return_ LPMDB GetMDB() const;
	_Check_return_ LPMAPISESSION GetSession() const;
	_Check_return_ LPMAPISESSION LogonGetSession(_In_ HWND hWnd);

	void SetAddrBook(_In_opt_ LPADRBOOK lpAddrBook);
	void SetMDB(_In_opt_ LPMDB lppMDB);
	void MAPILogonEx(_In_ HWND hwnd, _In_opt_z_ LPTSTR szProfileName, ULONG ulFlags);
	void Logoff(_In_ HWND hwnd, ULONG ulFlags);

	// For copy buffer
	_Check_return_ LPENTRYLIST GetABEntriesToCopy() const;
	_Check_return_ LPENTRYLIST GetMessagesToCopy() const;
	_Check_return_ LPMAPIFOLDER GetFolderToCopy() const;
	_Check_return_ LPMAPIFOLDER GetSourceParentFolder() const;

	_Check_return_ ULONG GetPropertyToCopy() const;
	_Check_return_ LPMAPIPROP GetSourcePropObject() const;

	_Check_return_ ULONG* GetAttachmentsToCopy() const;
	_Check_return_ ULONG GetNumAttachments() const;

	_Check_return_ string GetProfileToCopy() const;

	void SetABEntriesToCopy(_In_ LPENTRYLIST lpEBEntriesToCopy) const;
	void SetMessagesToCopy(_In_ LPENTRYLIST lpMessagesToCopy, _In_ LPMAPIFOLDER lpSourceParent) const;
	void SetFolderToCopy(_In_ LPMAPIFOLDER lpFolderToCopy, _In_ LPMAPIFOLDER lpSourceParent) const;
	void SetPropertyToCopy(ULONG ulPropTag, _In_ LPMAPIPROP lpSourcePropObject) const;
	void SetAttachmentsToCopy(_In_ LPMESSAGE lpMessage, ULONG ulNumSelected, _In_ ULONG* lpAttNumList) const;
	void SetProfileToCopy(_In_ LPCSTR szProfileName) const;

	_Check_return_ ULONG GetBufferStatus() const;

	void MAPIInitialize(ULONG ulFlags) const;
	void MAPIUninitialize() const;
	_Check_return_ bool bMAPIInitialized() const;

private:
	LONG m_cRef;
	LPMDB m_lpMDB;
	LPMAPISESSION m_lpMAPISession;
	LPADRBOOK m_lpAddrBook;
	CGlobalCache* m_lpGlobalCache;
};