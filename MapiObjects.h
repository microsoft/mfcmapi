// MapiObjects.h: interface for the CMapiObjects class.
//
//////////////////////////////////////////////////////////////////////

#pragma once

class CGlobalCache;

//For GetBufferStatus
#define BUFFER_EMPTY			((ULONG) 0x00000000)
#define BUFFER_MESSAGES			((ULONG) 0x00000001)
#define BUFFER_FOLDER			((ULONG) 0x00000002)
#define BUFFER_PARENTFOLDER		((ULONG) 0x00000004)
#define BUFFER_ABENTRIES		((ULONG) 0x00000008)
#define BUFFER_PROPTAG			((ULONG) 0x00000010)
#define BUFFER_SOURCEPROPOBJ	((ULONG) 0x00000020)

class CMapiObjects
{
public:
	CMapiObjects(
		CMapiObjects *OldMapiObjects);
	virtual ~CMapiObjects();

	STDMETHODIMP_(ULONG) AddRef();
 	STDMETHODIMP_(ULONG) Release();

	LPADRBOOK		GetAddrBook(BOOL bForceOpen);
	LPMDB			GetMDB();
	LPMAPISESSION	GetSession();
	LPPROFADMIN		GetProfAdmin();

	void	SetAddrBook(LPADRBOOK lpAddrBook);
	void	SetMDB(LPMDB lppMDB);
	void	MAPILogonEx(HWND hwnd,LPTSTR szProfileName, ULONG ulFlags);
	void	Logoff();

	//For copy buffer
	LPENTRYLIST		GetABEntriesToCopy();
	LPENTRYLIST		GetMessagesToCopy();
	LPMAPIFOLDER	GetFolderToCopy();
	LPMAPIFOLDER	GetSourceParentFolder();

	ULONG			GetPropertyToCopy();
	LPMAPIPROP		GetSourcePropObject();

	void	SetABEntriesToCopy(LPENTRYLIST lpEBEntriesToCopy);
	void	SetMessagesToCopy(LPENTRYLIST lpMessagesToCopy, LPMAPIFOLDER lpSourceParent);
	void	SetFolderToCopy(LPMAPIFOLDER lpFolderToCopy, LPMAPIFOLDER lpSourceParent);
	void	SetPropertyToCopy(ULONG ulPropTag, LPMAPIPROP lpSourcePropObject);

	ULONG	GetBufferStatus();

	void	MAPIInitialize(ULONG ulFlags);
	void	MAPIUninitialize();
	BOOL	bMAPIInitialized();

private:
	LONG			m_cRef;
	LPMDB			m_lpMDB;
	LPMAPISESSION	m_lpMAPISession;
	LPADRBOOK		m_lpAddrBook;
	LPPROFADMIN		m_lpProfAdmin;
	CGlobalCache*	m_lpGlobalCache;
};