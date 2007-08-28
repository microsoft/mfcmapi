// MapiObjects.cpp: implementation of the CMapiObjects class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "registry.h"

#include "MapiObjects.h"

#include "error.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

static TCHAR* GCCLASS = _T("CGlobalCache"); // STRING_OK

//A single instance cache of objects available to all
class CGlobalCache
{
public:
	CGlobalCache();
	virtual ~CGlobalCache();

	STDMETHODIMP_(ULONG) AddRef();
 	STDMETHODIMP_(ULONG) Release();

	void	MAPIInitialize(ULONG ulFlags);
	void	MAPIUninitialize();
	BOOL	bMAPIInitialized();

	void SetABEntriesToCopy(LPENTRYLIST lpEBEntriesToCopy);
	LPENTRYLIST GetABEntriesToCopy();

	void SetMessagesToCopy(LPENTRYLIST lpMessagesToCopy, LPMAPIFOLDER lpSourceParent);
	LPENTRYLIST GetMessagesToCopy();

	void SetFolderToCopy(LPMAPIFOLDER lpFolderToCopy, LPMAPIFOLDER lpSourceParent);
	LPMAPIFOLDER GetFolderToCopy();

	void SetPropertyToCopy(ULONG ulPropTag, LPMAPIPROP lpSourcePropObject);
	ULONG GetPropertyToCopy();
	LPMAPIPROP GetSourcePropObject();

	LPMAPIFOLDER GetSourceParentFolder();

	ULONG GetBufferStatus();

private:
	void EmptyBuffer();
	LONG			m_cRef;

	LPENTRYLIST		m_lpAddressEntriesToCopy;
	LPENTRYLIST		m_lpMessagesToCopy;
	LPMAPIFOLDER	m_lpFolderToCopy;
	ULONG			m_ulPropTagToCopy;
	LPMAPIFOLDER	m_lpSourceParent;
	LPMAPIPROP		m_lpSourcePropObject;
	BOOL			m_bMAPIInitialized;
};

CGlobalCache::CGlobalCache()
{
	TRACE_CONSTRUCTOR(GCCLASS);
	m_cRef = 1;
	m_bMAPIInitialized = FALSE;

	m_lpMessagesToCopy = NULL;
	m_lpFolderToCopy = NULL;
	m_lpAddressEntriesToCopy = NULL;
	m_ulPropTagToCopy = NULL;

	m_lpSourceParent = NULL;
	m_lpSourcePropObject = NULL;
}

CGlobalCache::~CGlobalCache()
{
	DebugPrintEx(DBGConDes,GCCLASS,GCCLASS,_T("(this = 0x%X) - Destructor\n"),this);
	EmptyBuffer();
	CGlobalCache::MAPIUninitialize();
}

STDMETHODIMP_(ULONG) CGlobalCache::AddRef()
{
	LONG lCount = InterlockedIncrement(&m_cRef);
	TRACE_ADDREF(GCCLASS,lCount);
	return lCount;
}

STDMETHODIMP_(ULONG) CGlobalCache::Release()
{
	LONG lCount = InterlockedDecrement(&m_cRef);
	TRACE_RELEASE(GCCLASS,lCount);
	if (!lCount)  delete this;
	return lCount;
}

void CGlobalCache::MAPIInitialize(ULONG ulFlags)
{
	HRESULT hRes = S_OK;
	if (!m_bMAPIInitialized)
	{
		MAPIINIT_0 mapiInit = {MAPI_INIT_VERSION,ulFlags};
		EC_H(::MAPIInitialize(&mapiInit));
		if (SUCCEEDED(hRes))
		{
			m_bMAPIInitialized = TRUE;
		}
	}
}// CGlobalCache::MAPIInitialize

void CGlobalCache::MAPIUninitialize()
{
	if (m_bMAPIInitialized)
	{
		::MAPIUninitialize();
		m_bMAPIInitialized = FALSE;
	}
}// CGlobalCache::MAPIUninitialize

BOOL CGlobalCache::bMAPIInitialized()
{
	return m_bMAPIInitialized;
}// CGlobalCache::bMAPIInitialized

void CGlobalCache::EmptyBuffer()
{
	MAPIFreeBuffer(m_lpAddressEntriesToCopy);
	MAPIFreeBuffer(m_lpMessagesToCopy);
	if (m_lpFolderToCopy) m_lpFolderToCopy->Release();
	if (m_lpSourceParent) m_lpSourceParent->Release();
	if (m_lpSourcePropObject) m_lpSourcePropObject->Release();

	m_lpAddressEntriesToCopy = NULL;
	m_lpMessagesToCopy = NULL;
	m_lpFolderToCopy = NULL;
	m_ulPropTagToCopy = NULL;
	m_lpSourceParent = NULL;
	m_lpSourcePropObject = NULL;
}

void CGlobalCache::SetABEntriesToCopy(LPENTRYLIST lpEBEntriesToCopy)
{
	EmptyBuffer();
	m_lpAddressEntriesToCopy = lpEBEntriesToCopy;
}

LPENTRYLIST CGlobalCache::GetABEntriesToCopy()
{
	return m_lpAddressEntriesToCopy;
}

void CGlobalCache::SetMessagesToCopy(LPENTRYLIST lpMessagesToCopy, LPMAPIFOLDER lpSourceParent)
{
	EmptyBuffer();
	m_lpMessagesToCopy = lpMessagesToCopy;
	m_lpSourceParent = lpSourceParent;
	if (m_lpSourceParent) m_lpSourceParent->AddRef();

}

LPENTRYLIST CGlobalCache::GetMessagesToCopy()
{
	return m_lpMessagesToCopy;
}

void CGlobalCache::SetFolderToCopy(LPMAPIFOLDER lpFolderToCopy, LPMAPIFOLDER lpSourceParent)
{
	EmptyBuffer();
	m_lpFolderToCopy = lpFolderToCopy;
	if (m_lpFolderToCopy) m_lpFolderToCopy->AddRef();
	m_lpSourceParent = lpSourceParent;
	if (m_lpSourceParent) m_lpSourceParent->AddRef();
}

LPMAPIFOLDER CGlobalCache::GetFolderToCopy()
{
	if (m_lpFolderToCopy) m_lpFolderToCopy->AddRef();
	return m_lpFolderToCopy;
}

LPMAPIFOLDER CGlobalCache::GetSourceParentFolder()
{
	if (m_lpSourceParent) m_lpSourceParent->AddRef();
	return m_lpSourceParent;
}

void CGlobalCache::SetPropertyToCopy(ULONG ulPropTag, LPMAPIPROP lpSourcePropObject)
{
	EmptyBuffer();
	m_ulPropTagToCopy = ulPropTag;
	m_lpSourcePropObject = lpSourcePropObject;
	if (m_lpSourcePropObject) m_lpSourcePropObject->AddRef();
}

ULONG CGlobalCache::GetPropertyToCopy()
{
	return m_ulPropTagToCopy;
}

LPMAPIPROP CGlobalCache::GetSourcePropObject()
{
	if (m_lpSourcePropObject) m_lpSourcePropObject->AddRef();
	return m_lpSourcePropObject;
}

ULONG CGlobalCache::GetBufferStatus()
{
	ULONG ulStatus = BUFFER_EMPTY;
	if (m_lpMessagesToCopy)			ulStatus |= BUFFER_MESSAGES;
	if (m_lpFolderToCopy)			ulStatus |= BUFFER_FOLDER;
	if (m_lpSourceParent)			ulStatus |= BUFFER_PARENTFOLDER;
	if (m_lpAddressEntriesToCopy)	ulStatus |= BUFFER_ABENTRIES;
	if (m_ulPropTagToCopy)			ulStatus |= BUFFER_PROPTAG;
	if (m_lpSourcePropObject)		ulStatus |= BUFFER_SOURCEPROPOBJ;
	return ulStatus;
}

static TCHAR* CLASS = _T("CMapiObjects");

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

//Pass an existing CMapiObjects to make a copy, pass NULL to create a new one from scratch
CMapiObjects::CMapiObjects(CMapiObjects *OldMapiObjects)
{
	TRACE_CONSTRUCTOR(CLASS);
	DebugPrintEx(DBGConDes,CLASS,CLASS,_T("OldMapiObjects = 0x%X\n"),OldMapiObjects);
	m_cRef = 1;

	m_lpAddrBook = NULL;
	m_lpMDB = NULL;
	m_lpMAPISession = NULL;
	m_lpProfAdmin = NULL;

	//If we were passed a valid object, make copies of its interfaces.
	if (OldMapiObjects)
	{
		m_lpMAPISession = OldMapiObjects->m_lpMAPISession;
		if (m_lpMAPISession) m_lpMAPISession->AddRef();

		m_lpMDB = OldMapiObjects->m_lpMDB;
		if (m_lpMDB) m_lpMDB->AddRef();

		m_lpAddrBook = OldMapiObjects->m_lpAddrBook;
		if (m_lpAddrBook) m_lpAddrBook->AddRef();

		m_lpProfAdmin = OldMapiObjects->m_lpProfAdmin;
		if (m_lpProfAdmin) m_lpProfAdmin->AddRef();

		m_lpGlobalCache = OldMapiObjects->m_lpGlobalCache;
		if (m_lpGlobalCache) m_lpGlobalCache->AddRef();
	}
	else
	{
		m_lpGlobalCache = new CGlobalCache();
	}
}

CMapiObjects::~CMapiObjects()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_lpAddrBook) m_lpAddrBook->Release();
	if (m_lpMDB) m_lpMDB->Release();
	if (m_lpMAPISession) m_lpMAPISession->Release();
	if (m_lpProfAdmin) m_lpProfAdmin->Release();

	// Must be last - uninitializes MAPI
	if (m_lpGlobalCache) m_lpGlobalCache->Release();
}

STDMETHODIMP_(ULONG) CMapiObjects::AddRef()
{
	LONG lCount = InterlockedIncrement(&m_cRef);
	TRACE_ADDREF(CLASS,lCount);
	return lCount;
}

STDMETHODIMP_(ULONG) CMapiObjects::Release()
{
	LONG lCount = InterlockedDecrement(&m_cRef);
	TRACE_RELEASE(CLASS,lCount);
	if (!lCount)  delete this;
	return lCount;
}

void CMapiObjects::MAPILogonEx(HWND hwnd,LPTSTR szProfileName, ULONG ulFlags)
{
	HRESULT hRes = S_OK;
	DebugPrint(DBGGeneric,_T("Logging on with MAPILogonEx, ulFlags = 0x%X\n"),ulFlags);

	this->MAPIInitialize(NULL);

	if (m_lpMAPISession) m_lpMAPISession->Release();
	m_lpMAPISession = NULL;

	EC_H_CANCEL(::MAPILogonEx(
		(ULONG_PTR)hwnd,
		szProfileName,
		0,
		ulFlags,
		&m_lpMAPISession));

	DebugPrint(DBGGeneric,_T("\tm_lpMAPISession set to 0x%X\n"),m_lpMAPISession);
}// CMapiObjects::MAPILogonEx

void CMapiObjects::Logoff()
{
	HRESULT hRes = S_OK;
	DebugPrint(DBGGeneric,_T("Logging off of 0x%X\n"),m_lpMAPISession);

	if (m_lpMAPISession)
	{
		EC_H(m_lpMAPISession->Logoff(NULL,MAPI_LOGOFF_UI,NULL));
		m_lpMAPISession->Release();
		m_lpMAPISession = NULL;
	}
}// CMapiObjects::Logoff

LPMAPISESSION CMapiObjects::GetSession()
{
	return m_lpMAPISession;
}

LPPROFADMIN CMapiObjects::GetProfAdmin()
{
	if (!m_lpProfAdmin)
	{
		HRESULT hRes = S_OK;
		this->MAPIInitialize(NULL);
		EC_H(MAPIAdminProfiles(0, &m_lpProfAdmin));
	}
	return m_lpProfAdmin;
}

void CMapiObjects::SetMDB(LPMDB lpMDB)
{
	DebugPrintEx(DBGGeneric,CLASS,_T("SetMDB"),_T("replacing 0x%X with 0x%X\n"),m_lpMDB,lpMDB);
	if (m_lpMDB) m_lpMDB->Release();
	m_lpMDB = lpMDB;
	if (m_lpMDB) m_lpMDB->AddRef();
}

LPMDB CMapiObjects::GetMDB()
{
	return m_lpMDB;
}

void CMapiObjects::SetAddrBook(LPADRBOOK lpAddrBook)
{
	DebugPrintEx(DBGGeneric,CLASS,_T("SetAddrBook"),_T("replacing 0x%X with 0x%X\n"),m_lpAddrBook,lpAddrBook);
	if (m_lpAddrBook) m_lpAddrBook->Release();
	m_lpAddrBook = lpAddrBook;
	if (m_lpAddrBook) m_lpAddrBook->AddRef();
}

LPADRBOOK CMapiObjects::GetAddrBook(BOOL bForceOpen)
{
	//if we haven't opened the address book yet and we have a session, open it now
	if (!m_lpAddrBook && m_lpMAPISession && bForceOpen)
	{
		HRESULT hRes = S_OK;
		EC_H(m_lpMAPISession->OpenAddressBook(
			NULL,
			NULL,
			NULL,
			&m_lpAddrBook));
	}
	return m_lpAddrBook;
}

void CMapiObjects::MAPIInitialize(ULONG ulFlags)
{
	if (m_lpGlobalCache)
	{
		m_lpGlobalCache->MAPIInitialize(ulFlags);
	}
}

void CMapiObjects::MAPIUninitialize()
{
	if (m_lpGlobalCache)
	{
		m_lpGlobalCache->MAPIUninitialize();
	}
}

BOOL CMapiObjects::bMAPIInitialized()
{
	if (m_lpGlobalCache)
	{
		return m_lpGlobalCache->bMAPIInitialized();
	}
	return false;
}

void CMapiObjects::SetABEntriesToCopy(LPENTRYLIST lpEBEntriesToCopy)
{
	if (m_lpGlobalCache)
	{
		m_lpGlobalCache->SetABEntriesToCopy(lpEBEntriesToCopy);
	}
}

LPENTRYLIST CMapiObjects::GetABEntriesToCopy()
{
	if (m_lpGlobalCache)
	{
		return m_lpGlobalCache->GetABEntriesToCopy();
	}
	return NULL;
}

void CMapiObjects::SetMessagesToCopy(LPENTRYLIST lpMessagesToCopy, LPMAPIFOLDER lpSourceParent)
{
	if (m_lpGlobalCache)
	{
		m_lpGlobalCache->SetMessagesToCopy(lpMessagesToCopy,lpSourceParent);
	}
}

LPENTRYLIST CMapiObjects::GetMessagesToCopy()
{
	if (m_lpGlobalCache)
	{
		return m_lpGlobalCache->GetMessagesToCopy();
	}
	return NULL;
}

void CMapiObjects::SetFolderToCopy(LPMAPIFOLDER lpFolderToCopy, LPMAPIFOLDER lpSourceParent)
{
	if (m_lpGlobalCache)
	{
		m_lpGlobalCache->SetFolderToCopy(lpFolderToCopy,lpSourceParent);
	}
}

LPMAPIFOLDER CMapiObjects::GetFolderToCopy()
{
	if (m_lpGlobalCache)
	{
		return m_lpGlobalCache->GetFolderToCopy();
	}
	return NULL;
}

LPMAPIFOLDER CMapiObjects::GetSourceParentFolder()
{
	if (m_lpGlobalCache)
	{
		return m_lpGlobalCache->GetSourceParentFolder();
	}
	return NULL;
}

void CMapiObjects::SetPropertyToCopy(ULONG ulPropTag, LPMAPIPROP lpSourcePropObject)
{
	if (m_lpGlobalCache)
	{
		m_lpGlobalCache->SetPropertyToCopy(ulPropTag,lpSourcePropObject);
	}
}

ULONG CMapiObjects::GetPropertyToCopy()
{
	if (m_lpGlobalCache)
	{
		return m_lpGlobalCache->GetPropertyToCopy();
	}
	return NULL;
}

LPMAPIPROP CMapiObjects::GetSourcePropObject()
{
	if (m_lpGlobalCache)
	{
		return m_lpGlobalCache->GetSourcePropObject();
	}
	return NULL;
}

ULONG CMapiObjects::GetBufferStatus()
{
	if (m_lpGlobalCache)
	{
		return m_lpGlobalCache->GetBufferStatus();
	}
	return BUFFER_EMPTY;
}
