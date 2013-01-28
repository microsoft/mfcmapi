// MapiObjects.cpp: implementation of the CMapiObjects class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "MapiObjects.h"
#include "MAPIFunctions.h"

static TCHAR* GCCLASS = _T("CGlobalCache"); // STRING_OK

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
	_Check_return_ bool bMAPIInitialized();

	void SetABEntriesToCopy(_In_ LPENTRYLIST lpEBEntriesToCopy);
	_Check_return_ LPENTRYLIST GetABEntriesToCopy();

	void SetMessagesToCopy(_In_ LPENTRYLIST lpMessagesToCopy, _In_ LPMAPIFOLDER lpSourceParent);
	_Check_return_ LPENTRYLIST GetMessagesToCopy();

	void SetFolderToCopy(_In_ LPMAPIFOLDER lpFolderToCopy, _In_ LPMAPIFOLDER lpSourceParent);
	_Check_return_ LPMAPIFOLDER GetFolderToCopy();

	void SetPropertyToCopy(ULONG ulPropTag, _In_ LPMAPIPROP lpSourcePropObject);
	_Check_return_ ULONG GetPropertyToCopy();
	_Check_return_ LPMAPIPROP GetSourcePropObject();

	void SetAttachmentsToCopy(_In_ LPMESSAGE lpMessage, ULONG ulNumAttachments, _In_ ULONG* lpAttNumList);
	_Check_return_ ULONG* GetAttachmentsToCopy();
	_Check_return_ ULONG GetNumAttachments();

	void SetProfileToCopy(_In_ LPSTR szProfileName);
	_Check_return_ LPSTR GetProfileToCopy();

	_Check_return_ LPMAPIFOLDER GetSourceParentFolder();

	_Check_return_ ULONG GetBufferStatus();

private:
	void EmptyBuffer();

	LONG			m_cRef;
	LPENTRYLIST		m_lpAddressEntriesToCopy;
	LPENTRYLIST		m_lpMessagesToCopy;
	LPMAPIFOLDER	m_lpFolderToCopy;
	ULONG			m_ulPropTagToCopy;
	ULONG*			m_lpulAttachmentsToCopy;
	ULONG			m_ulNumAttachments;
	LPSTR			m_szProfileToCopy;
	LPMAPIFOLDER	m_lpSourceParent;
	LPMAPIPROP		m_lpSourcePropObject;
	bool			m_bMAPIInitialized;
};

CGlobalCache::CGlobalCache()
{
	TRACE_CONSTRUCTOR(GCCLASS);
	m_cRef = 1;
	m_bMAPIInitialized = false;

	m_lpMessagesToCopy = NULL;
	m_lpFolderToCopy = NULL;
	m_lpAddressEntriesToCopy = NULL;
	m_ulPropTagToCopy = NULL;

	m_lpSourceParent = NULL;
	m_lpSourcePropObject = NULL;

	m_lpulAttachmentsToCopy = NULL;
	m_ulNumAttachments = NULL;

	m_szProfileToCopy = NULL;
} // CGlobalCache::CGlobalCache

CGlobalCache::~CGlobalCache()
{
	TRACE_DESTRUCTOR(GCCLASS);
	EmptyBuffer();
	CGlobalCache::MAPIUninitialize();
} // CGlobalCache::~CGlobalCache

STDMETHODIMP_(ULONG) CGlobalCache::AddRef()
{
	LONG lCount = InterlockedIncrement(&m_cRef);
	TRACE_ADDREF(GCCLASS,lCount);
	return lCount;
} // CGlobalCache::AddRef

STDMETHODIMP_(ULONG) CGlobalCache::Release()
{
	LONG lCount = InterlockedDecrement(&m_cRef);
	TRACE_RELEASE(GCCLASS,lCount);
	if (!lCount)  delete this;
	return lCount;
} // CGlobalCache::Release

void CGlobalCache::MAPIInitialize(ULONG ulFlags)
{
	HRESULT hRes = S_OK;
	if (!m_bMAPIInitialized)
	{
		MAPIINIT_0 mapiInit = {MAPI_INIT_VERSION,ulFlags};
		WC_MAPI(::MAPIInitialize(&mapiInit));
		if (SUCCEEDED(hRes))
		{
			m_bMAPIInitialized = true;
		}
		else
		{
			ErrDialog(__FILE__,__LINE__,
				IDS_EDMAPIINITIALIZEFAILED,
				hRes,
				ErrorNameFromErrorCode(hRes));
		}
	}
} // CGlobalCache::MAPIInitialize

void CGlobalCache::MAPIUninitialize()
{
	if (m_bMAPIInitialized)
	{
		::MAPIUninitialize();
		m_bMAPIInitialized = false;
	}
} // CGlobalCache::MAPIUninitialize

_Check_return_ bool CGlobalCache::bMAPIInitialized()
{
	return m_bMAPIInitialized;
} // CGlobalCache::bMAPIInitialized

void CGlobalCache::EmptyBuffer()
{
	if (m_lpAddressEntriesToCopy) MAPIFreeBuffer(m_lpAddressEntriesToCopy);
	if (m_lpMessagesToCopy) MAPIFreeBuffer(m_lpMessagesToCopy);
	if (m_lpulAttachmentsToCopy) MAPIFreeBuffer(m_lpulAttachmentsToCopy);
	if (m_szProfileToCopy) MAPIFreeBuffer(m_szProfileToCopy);
	if (m_lpFolderToCopy) m_lpFolderToCopy->Release();
	if (m_lpSourceParent) m_lpSourceParent->Release();
	if (m_lpSourcePropObject) m_lpSourcePropObject->Release();

	m_lpAddressEntriesToCopy = NULL;
	m_lpMessagesToCopy = NULL;
	m_lpFolderToCopy = NULL;
	m_ulPropTagToCopy = NULL;
	m_lpSourceParent = NULL;
	m_lpSourcePropObject = NULL;
	m_lpulAttachmentsToCopy = NULL;
	m_ulNumAttachments = NULL;
	m_szProfileToCopy = NULL;
} // CGlobalCache::EmptyBuffer

void CGlobalCache::SetABEntriesToCopy(_In_ LPENTRYLIST lpEBEntriesToCopy)
{
	EmptyBuffer();
	m_lpAddressEntriesToCopy = lpEBEntriesToCopy;
} // CGlobalCache::SetABEntriesToCopy

_Check_return_ LPENTRYLIST CGlobalCache::GetABEntriesToCopy()
{
	return m_lpAddressEntriesToCopy;
} // CGlobalCache::GetABEntriesToCopy

void CGlobalCache::SetMessagesToCopy(_In_ LPENTRYLIST lpMessagesToCopy, _In_ LPMAPIFOLDER lpSourceParent)
{
	EmptyBuffer();
	m_lpMessagesToCopy = lpMessagesToCopy;
	m_lpSourceParent = lpSourceParent;
	if (m_lpSourceParent) m_lpSourceParent->AddRef();
} // CGlobalCache::SetMessagesToCopy

_Check_return_ LPENTRYLIST CGlobalCache::GetMessagesToCopy()
{
	return m_lpMessagesToCopy;
} // CGlobalCache::GetMessagesToCopy

void CGlobalCache::SetFolderToCopy(_In_ LPMAPIFOLDER lpFolderToCopy, _In_ LPMAPIFOLDER lpSourceParent)
{
	EmptyBuffer();
	m_lpFolderToCopy = lpFolderToCopy;
	if (m_lpFolderToCopy) m_lpFolderToCopy->AddRef();
	m_lpSourceParent = lpSourceParent;
	if (m_lpSourceParent) m_lpSourceParent->AddRef();
} // CGlobalCache::SetFolderToCopy

_Check_return_ LPMAPIFOLDER CGlobalCache::GetFolderToCopy()
{
	if (m_lpFolderToCopy) m_lpFolderToCopy->AddRef();
	return m_lpFolderToCopy;
} // CGlobalCache::GetFolderToCopy

_Check_return_ LPMAPIFOLDER CGlobalCache::GetSourceParentFolder()
{
	if (m_lpSourceParent) m_lpSourceParent->AddRef();
	return m_lpSourceParent;
} // CGlobalCache::GetSourceParentFolder

void CGlobalCache::SetPropertyToCopy(ULONG ulPropTag, _In_ LPMAPIPROP lpSourcePropObject)
{
	EmptyBuffer();
	m_ulPropTagToCopy = ulPropTag;
	m_lpSourcePropObject = lpSourcePropObject;
	if (m_lpSourcePropObject) m_lpSourcePropObject->AddRef();
} // CGlobalCache::SetPropertyToCopy

_Check_return_ ULONG CGlobalCache::GetPropertyToCopy()
{
	return m_ulPropTagToCopy;
} // CGlobalCache::GetPropertyToCopy

_Check_return_ LPMAPIPROP CGlobalCache::GetSourcePropObject()
{
	if (m_lpSourcePropObject) m_lpSourcePropObject->AddRef();
	return m_lpSourcePropObject;
} // CGlobalCache::GetSourcePropObject

void CGlobalCache::SetAttachmentsToCopy(_In_ LPMESSAGE lpMessage, ULONG ulNumAttachments, _In_ ULONG* lpAttNumList)
{
	EmptyBuffer();
	m_lpulAttachmentsToCopy = lpAttNumList;
	m_ulNumAttachments = ulNumAttachments;
	m_lpSourcePropObject = lpMessage;
	if (m_lpSourcePropObject) m_lpSourcePropObject->AddRef();
} // CGlobalCache::SetAttachmentsToCopy

_Check_return_ ULONG* CGlobalCache::GetAttachmentsToCopy()
{
	return m_lpulAttachmentsToCopy;
} // CGlobalCache::GetAttachmentsToCopy

_Check_return_ ULONG CGlobalCache::GetNumAttachments()
{
	return m_ulNumAttachments;
} // CGlobalCache::GetNumAttachments

void CGlobalCache::SetProfileToCopy(_In_ LPSTR szProfileName)
{
	(void) CopyStringA(&m_szProfileToCopy, szProfileName, NULL);
} // CGlobalCache::SetProfileToCopy

_Check_return_ LPSTR CGlobalCache::GetProfileToCopy()
{
	return m_szProfileToCopy;
} // CGlobalCache::GetProfileToCopy

_Check_return_ ULONG CGlobalCache::GetBufferStatus()
{
	ULONG ulStatus = BUFFER_EMPTY;
	if (m_lpMessagesToCopy)			ulStatus |= BUFFER_MESSAGES;
	if (m_lpFolderToCopy)			ulStatus |= BUFFER_FOLDER;
	if (m_lpSourceParent)			ulStatus |= BUFFER_PARENTFOLDER;
	if (m_lpAddressEntriesToCopy)	ulStatus |= BUFFER_ABENTRIES;
	if (m_ulPropTagToCopy)			ulStatus |= BUFFER_PROPTAG;
	if (m_lpSourcePropObject)		ulStatus |= BUFFER_SOURCEPROPOBJ;
	if (m_lpulAttachmentsToCopy)	ulStatus |= BUFFER_ATTACHMENTS;
	if (m_szProfileToCopy)			ulStatus |= BUFFER_PROFILE;
	return ulStatus;
} // CGlobalCache::GetBufferStatus

static TCHAR* CLASS = _T("CMapiObjects");

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

// Pass an existing CMapiObjects to make a copy, pass NULL to create a new one from scratch
CMapiObjects::CMapiObjects(_In_opt_ CMapiObjects *OldMapiObjects)
{
	TRACE_CONSTRUCTOR(CLASS);
	DebugPrintEx(DBGConDes,CLASS,CLASS,_T("OldMapiObjects = %p\n"),OldMapiObjects);
	m_cRef = 1;

	m_lpAddrBook = NULL;
	m_lpMDB = NULL;
	m_lpMAPISession = NULL;

	// If we were passed a valid object, make copies of its interfaces.
	if (OldMapiObjects)
	{
		m_lpMAPISession = OldMapiObjects->m_lpMAPISession;
		if (m_lpMAPISession) m_lpMAPISession->AddRef();

		m_lpMDB = OldMapiObjects->m_lpMDB;
		if (m_lpMDB) m_lpMDB->AddRef();

		m_lpAddrBook = OldMapiObjects->m_lpAddrBook;
		if (m_lpAddrBook) m_lpAddrBook->AddRef();

		m_lpGlobalCache = OldMapiObjects->m_lpGlobalCache;
		if (m_lpGlobalCache) m_lpGlobalCache->AddRef();
	}
	else
	{
		m_lpGlobalCache = new CGlobalCache();
	}
} // CMapiObjects::CMapiObjects

CMapiObjects::~CMapiObjects()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_lpAddrBook) m_lpAddrBook->Release();
	if (m_lpMDB) m_lpMDB->Release();
	if (m_lpMAPISession) m_lpMAPISession->Release();

	// Must be last - uninitializes MAPI
	if (m_lpGlobalCache) m_lpGlobalCache->Release();
} // CMapiObjects::~CMapiObjects

STDMETHODIMP_(ULONG) CMapiObjects::AddRef()
{
	LONG lCount = InterlockedIncrement(&m_cRef);
	TRACE_ADDREF(CLASS,lCount);
	return lCount;
} // CMapiObjects::AddRef

STDMETHODIMP_(ULONG) CMapiObjects::Release()
{
	LONG lCount = InterlockedDecrement(&m_cRef);
	TRACE_RELEASE(CLASS,lCount);
	if (!lCount)  delete this;
	return lCount;
} // CMapiObjects::Release

void CMapiObjects::MAPILogonEx(_In_ HWND hwnd, _In_opt_z_ LPTSTR szProfileName, ULONG ulFlags)
{
	HRESULT hRes = S_OK;
	DebugPrint(DBGGeneric,_T("Logging on with MAPILogonEx, ulFlags = 0x%X\n"),ulFlags);

	this->MAPIInitialize(NULL);
	if (!bMAPIInitialized()) return;

	if (m_lpMAPISession) m_lpMAPISession->Release();
	m_lpMAPISession = NULL;

	EC_H_CANCEL(::MAPILogonEx(
		(ULONG_PTR)hwnd,
		szProfileName,
		0,
		ulFlags,
		&m_lpMAPISession));

	DebugPrint(DBGGeneric,_T("\tm_lpMAPISession set to %p\n"),m_lpMAPISession);
} // CMapiObjects::MAPILogonEx

void CMapiObjects::Logoff(_In_ HWND hwnd, ULONG ulFlags)
{
	HRESULT hRes = S_OK;
	DebugPrint(DBGGeneric,_T("Logging off of %p, ulFlags = 0x%08X\n"),m_lpMAPISession,ulFlags);

	if (m_lpMAPISession)
	{
		EC_MAPI(m_lpMAPISession->Logoff((ULONG_PTR)hwnd,ulFlags,NULL));
		m_lpMAPISession->Release();
		m_lpMAPISession = NULL;
	}
} // CMapiObjects::Logoff

_Check_return_ LPMAPISESSION CMapiObjects::GetSession()
{
	return m_lpMAPISession;
} // CMapiObjects::GetSession

_Check_return_ LPMAPISESSION CMapiObjects::LogonGetSession(_In_ HWND hWnd)
{
	if (m_lpMAPISession) return m_lpMAPISession;

	CMapiObjects::MAPILogonEx(hWnd, NULL, MAPI_EXTENDED | MAPI_LOGON_UI | MAPI_NEW_SESSION);

	return m_lpMAPISession;
} // CMapiObjects::LogonGetSession

void CMapiObjects::SetMDB(_In_opt_ LPMDB lpMDB)
{
	DebugPrintEx(DBGGeneric,CLASS,_T("SetMDB"),_T("replacing %p with %p\n"),m_lpMDB,lpMDB);
	if (m_lpMDB) m_lpMDB->Release();
	m_lpMDB = lpMDB;
	if (m_lpMDB) m_lpMDB->AddRef();
} // CMapiObjects::SetMDB

_Check_return_ LPMDB CMapiObjects::GetMDB()
{
	return m_lpMDB;
} // CMapiObjects::GetMDB

void CMapiObjects::SetAddrBook(_In_opt_ LPADRBOOK lpAddrBook)
{
	DebugPrintEx(DBGGeneric,CLASS,_T("SetAddrBook"),_T("replacing %p with %p\n"),m_lpAddrBook,lpAddrBook);
	if (m_lpAddrBook) m_lpAddrBook->Release();
	m_lpAddrBook = lpAddrBook;
	if (m_lpAddrBook) m_lpAddrBook->AddRef();
} // CMapiObjects::SetAddrBook

_Check_return_ LPADRBOOK CMapiObjects::GetAddrBook(bool bForceOpen)
{
	// if we haven't opened the address book yet and we have a session, open it now
	if (!m_lpAddrBook && m_lpMAPISession && bForceOpen)
	{
		HRESULT hRes = S_OK;
		EC_MAPI(m_lpMAPISession->OpenAddressBook(
			NULL,
			NULL,
			NULL,
			&m_lpAddrBook));
	}
	return m_lpAddrBook;
} // CMapiObjects::GetAddrBook

void CMapiObjects::MAPIInitialize(ULONG ulFlags)
{
	if (m_lpGlobalCache)
	{
		m_lpGlobalCache->MAPIInitialize(ulFlags);
	}
} // CMapiObjects::MAPIInitialize

void CMapiObjects::MAPIUninitialize()
{
	if (m_lpGlobalCache)
	{
		m_lpGlobalCache->MAPIUninitialize();
	}
} // CMapiObjects::MAPIUninitialize

_Check_return_ bool CMapiObjects::bMAPIInitialized()
{
	if (m_lpGlobalCache)
	{
		return m_lpGlobalCache->bMAPIInitialized();
	}
	return false;
} // CMapiObjects::bMAPIInitialized

void CMapiObjects::SetABEntriesToCopy(_In_ LPENTRYLIST lpEBEntriesToCopy)
{
	if (m_lpGlobalCache)
	{
		m_lpGlobalCache->SetABEntriesToCopy(lpEBEntriesToCopy);
	}
} // CMapiObjects::SetABEntriesToCopy

_Check_return_ LPENTRYLIST CMapiObjects::GetABEntriesToCopy()
{
	if (m_lpGlobalCache)
	{
		return m_lpGlobalCache->GetABEntriesToCopy();
	}
	return NULL;
} // CMapiObjects::GetABEntriesToCopy

void CMapiObjects::SetMessagesToCopy(_In_ LPENTRYLIST lpMessagesToCopy, _In_ LPMAPIFOLDER lpSourceParent)
{
	if (m_lpGlobalCache)
	{
		m_lpGlobalCache->SetMessagesToCopy(lpMessagesToCopy,lpSourceParent);
	}
} // CMapiObjects::SetMessagesToCopy

_Check_return_ LPENTRYLIST CMapiObjects::GetMessagesToCopy()
{
	if (m_lpGlobalCache)
	{
		return m_lpGlobalCache->GetMessagesToCopy();
	}
	return NULL;
} // CMapiObjects::GetMessagesToCopy

void CMapiObjects::SetFolderToCopy(_In_ LPMAPIFOLDER lpFolderToCopy, _In_ LPMAPIFOLDER lpSourceParent)
{
	if (m_lpGlobalCache)
	{
		m_lpGlobalCache->SetFolderToCopy(lpFolderToCopy,lpSourceParent);
	}
} // CMapiObjects::SetFolderToCopy

_Check_return_ LPMAPIFOLDER CMapiObjects::GetFolderToCopy()
{
	if (m_lpGlobalCache)
	{
		return m_lpGlobalCache->GetFolderToCopy();
	}
	return NULL;
} // CMapiObjects::GetFolderToCopy

_Check_return_ LPMAPIFOLDER CMapiObjects::GetSourceParentFolder()
{
	if (m_lpGlobalCache)
	{
		return m_lpGlobalCache->GetSourceParentFolder();
	}
	return NULL;
} // CMapiObjects::GetSourceParentFolder

void CMapiObjects::SetPropertyToCopy(ULONG ulPropTag, _In_ LPMAPIPROP lpSourcePropObject)
{
	if (m_lpGlobalCache)
	{
		m_lpGlobalCache->SetPropertyToCopy(ulPropTag,lpSourcePropObject);
	}
} // CMapiObjects::SetPropertyToCopy

_Check_return_ ULONG CMapiObjects::GetPropertyToCopy()
{
	if (m_lpGlobalCache)
	{
		return m_lpGlobalCache->GetPropertyToCopy();
	}
	return NULL;
} // CMapiObjects::GetPropertyToCopy

_Check_return_ LPMAPIPROP CMapiObjects::GetSourcePropObject()
{
	if (m_lpGlobalCache)
	{
		return m_lpGlobalCache->GetSourcePropObject();
	}
	return NULL;
} // CMapiObjects::GetSourcePropObject

void CMapiObjects::SetAttachmentsToCopy(_In_ LPMESSAGE lpMessage, ULONG ulNumAttachments, _In_ ULONG* lpAttNumList)
{
	if (m_lpGlobalCache)
	{
		m_lpGlobalCache->SetAttachmentsToCopy(lpMessage,ulNumAttachments,lpAttNumList);
	}
} // CMapiObjects::SetAttachmentsToCopy

_Check_return_ ULONG* CMapiObjects::GetAttachmentsToCopy()
{
	if (m_lpGlobalCache)
	{
		return m_lpGlobalCache->GetAttachmentsToCopy();
	}
	return NULL;
} // CMapiObjects::GetAttachmentsToCopy

_Check_return_ ULONG CMapiObjects::GetNumAttachments()
{
	if (m_lpGlobalCache)
	{
		return m_lpGlobalCache->GetNumAttachments();
	}
	return NULL;
} // CMapiObjects::GetNumAttachments

void CMapiObjects::SetProfileToCopy(_In_ LPSTR szProfileName)
{
	if (m_lpGlobalCache)
	{
		m_lpGlobalCache->SetProfileToCopy(szProfileName);
	}
} // CMapiObjects::SetProfileToCopy

_Check_return_ LPSTR CMapiObjects::GetProfileToCopy()
{
	if (m_lpGlobalCache)
	{
		return m_lpGlobalCache->GetProfileToCopy();
	}
	return NULL;
} // CMapiObjects::GetProfileToCopy

_Check_return_ ULONG CMapiObjects::GetBufferStatus()
{
	if (m_lpGlobalCache)
	{
		return m_lpGlobalCache->GetBufferStatus();
	}
	return BUFFER_EMPTY;
} // CMapiObjects::GetBufferStatus