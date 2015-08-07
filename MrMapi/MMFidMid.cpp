#include "stdafx.h"
#include "..\stdafx.h"
#include "MrMAPI.h"
#include "MMFidMid.h"
#include "..\MAPIFunctions.h"
#include "..\MapiProcessor.h"
#include "..\ExtraPropTags.h"
#include "..\SmartView\SmartView.h"
#include "..\String.h"

inline void PrintFolder(LPCWSTR szFid, LPCTSTR szFolder)
{
	printf("%-15ws %s\n", szFid, szFolder);
}

inline void PrintMessage(LPCWSTR szMid, bool fAssociated, LPCTSTR szSubject, LPCTSTR szClass)
{
	printf("     %-15ws %c %s (%s)\n", szMid, (fAssociated ? 'A' : 'R'), szSubject, szClass);
}

class CFindFidMid : public CMAPIProcessor
{
public:
	CFindFidMid();
	virtual ~CFindFidMid();
	void InitFidMid(_In_z_ LPCWSTR szFid, _In_z_ LPCWSTR szMid);

private:
	virtual bool ContinueProcessingFolders();
	virtual bool ShouldProcessContentsTable();
	virtual void BeginFolderWork();
	virtual void BeginContentsTableWork(ULONG ulFlags, ULONG ulCountRows);
	virtual bool DoContentsTablePerRowWork(_In_ LPSRow lpSRow, ULONG ulCurRow);

	LPCWSTR m_szFid;
	LPCWSTR m_szMid;
	LPWSTR m_szCurrentFid;

	bool m_fFIDMatch;
	bool m_fFIDExactMatch;
	bool m_fFIDPrinted;
	bool m_fAssociated;
};

CFindFidMid::CFindFidMid()
{
	m_szFid = NULL;
	m_szMid = NULL;
	m_szCurrentFid = NULL;

	m_fFIDMatch = false;
	m_fFIDExactMatch = false;
	m_fFIDPrinted = false;
	m_fAssociated = false;
}

CFindFidMid::~CFindFidMid()
{
	delete[] m_szCurrentFid;
	m_szCurrentFid = NULL;
}

void CFindFidMid::InitFidMid(_In_z_ LPCWSTR szFid, _In_z_ LPCWSTR szMid)
{
	m_szFid = szFid;
	m_szMid = szMid;
}

// --------------------------------------------------------------------------------- //

void CFindFidMid::BeginFolderWork()
{
	DebugPrint(DBGGeneric, "CFindFidMid::BeginFolderWork: m_szFolderOffset %s\n", m_szFolderOffset);
	m_fFIDMatch = false;
	m_fFIDExactMatch = false;
	m_fFIDPrinted = false;
	delete[] m_szCurrentFid;
	m_szCurrentFid = NULL;
	if (!m_lpFolder) return;

	HRESULT hRes = S_OK;
	ULONG ulProps = NULL;
	LPSPropValue lpProps = NULL;

	enum
	{
		ePR_DISPLAY_NAME_W,
		ePidTagFolderId,
		NUM_COLS
	};
	static const SizedSPropTagArray(NUM_COLS, sptaFolderProps) =
	{
		NUM_COLS,
		PR_DISPLAY_NAME_W,
		PidTagFolderId
	};

	WC_H_GETPROPS(m_lpFolder->GetProps(
		(LPSPropTagArray)&sptaFolderProps,
		NULL,
		&ulProps,
		&lpProps));
	DebugPrint(DBGGeneric, "CFindFidMid::DoFolderPerHierarchyTableRowWork: m_szFolderOffset %s\n", m_szFolderOffset);

	// Now get the FID
	LPWSTR lpszDisplayName = NULL;

	if (lpProps && PR_DISPLAY_NAME_W == lpProps[ePR_DISPLAY_NAME_W].ulPropTag)
	{
		lpszDisplayName = lpProps[ePR_DISPLAY_NAME_W].Value.lpszW;
	}

	if (lpProps && PidTagFolderId == lpProps[ePidTagFolderId].ulPropTag)
	{
		m_szCurrentFid = wstringToLPWSTR(FidMidToSzString(lpProps[ePidTagFolderId].Value.li.QuadPart, false));
		DebugPrint(DBGGeneric, "CFindFidMid::DoFolderPerHierarchyTableRowWork: Found FID %ws for %ws\n", m_szCurrentFid, lpszDisplayName);
	}
	else
	{
		// Nothing left to do if we can't find a fid.
		DebugPrint(DBGGeneric, "CFindFidMid::DoFolderPerHierarchyTableRowWork: No FID found for %ws\n", lpszDisplayName);
		return;
	}

	// If FidMidToSzString failed, we're done.
	if (!m_szCurrentFid) return;

	// Check for FID matches - no fid matches all folders
	if (NULL == m_szFid)
	{
		m_fFIDMatch = true;
		m_fFIDExactMatch = false;
	}
	// Passed in Fid matches the found Fid exactly, or matches the tail exactly
	// For instance, both 3-15632 and 15632 will match against an input fid of 15632
	else if (_wcsicmp(m_szFid, m_szCurrentFid) == 0
		|| (wcschr(m_szCurrentFid, '-') != 0 && _wcsicmp(m_szFid, wcschr(m_szCurrentFid, '-') + 1) == 0))
	{
		m_fFIDMatch = true;
		m_fFIDExactMatch = true;
	}

	if (m_fFIDMatch)
	{
		DebugPrint(DBGGeneric, "CFindFidMid::DoFolderPerHierarchyTableRowWork: Matched FID %ws\n", m_szFid);
		// Print out the folder
		if (m_szMid == NULL || m_fFIDExactMatch)
		{
			PrintFolder(m_szCurrentFid, m_szFolderOffset);
			m_fFIDPrinted = true;
		}
	}
}

bool CFindFidMid::ContinueProcessingFolders()
{
	// if we've found an exact match, we can stop
	return !m_fFIDExactMatch;
}

bool CFindFidMid::ShouldProcessContentsTable()
{
	// Only process a folder's contents table if both
	// 1 - we matched our fid, possibly because a fid wasn't passed in
	// 2 - We have a mid to look for, even if it's an empty string
	return m_fFIDMatch && m_szMid;
}

void CFindFidMid::BeginContentsTableWork(ULONG ulFlags, ULONG /*ulCountRows*/)
{
	m_fAssociated = ((ulFlags & MAPI_ASSOCIATED) == MAPI_ASSOCIATED);
}

bool CFindFidMid::DoContentsTablePerRowWork(_In_ LPSRow lpSRow, ULONG /*ulCurRow*/)
{
	if (!lpSRow) return false;

	LPSPropValue lpPropMid = NULL;
	LPSPropValue lpPropSubject = NULL;
	LPSPropValue lpPropClass = NULL;
	wstring lpszThisMid;
	LPTSTR lpszSubject = _T("");
	LPTSTR lpszClass = _T("");

	lpPropMid = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PidTagMid);
	if (lpPropMid)
	{
		lpszThisMid = FidMidToSzString(lpPropMid->Value.li.QuadPart, false);
		DebugPrint(DBGGeneric, "CFindFidMid::DoContentsTablePerRowWork: Found MID %ws\n", lpszThisMid);
	}
	else
	{
		// Nothing left to do if we can't find a mid
		DebugPrint(DBGGeneric, "CFindFidMid::DoContentsTablePerRowWork: No MID found\n");
		return false;
	}

	// Check for a MID match
	if (wcscmp(m_szMid, L"") == 0 || _wcsicmp(m_szMid, lpszThisMid.c_str()) == 0
		|| (wcschr(lpszThisMid.c_str(), L'-') != 0 && _wcsicmp(m_szMid, wcschr(lpszThisMid.c_str(), L'-') + 1) == 0))
	{
		// If we haven't already, print the folder info
		if (!m_fFIDPrinted)
		{
			PrintFolder(m_szCurrentFid, m_szFolderOffset);
			m_fFIDPrinted = true;
		}

		lpPropSubject = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_SUBJECT);
		if (lpPropSubject)
		{
			lpszSubject = lpPropSubject->Value.LPSZ;
		}

		lpPropClass = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_MESSAGE_CLASS);
		if (lpPropClass)
		{
			lpszClass = lpPropClass->Value.LPSZ;
		}

		PrintMessage(lpszThisMid.c_str(), m_fAssociated, lpszSubject, lpszClass);
		DebugPrint(DBGGeneric, "EnumMessages::ProcessRow: Matched MID %ws, \"%s\", \"%s\"\n", lpszThisMid, lpszSubject, lpszClass);
	}

	// Don't open the message - the row is enough
	return false;
}

void DumpFidMid(
	_In_z_ LPCWSTR lpszProfile,
	_In_ LPMDB lpMDB,
	_In_z_ LPCWSTR lpszFid,
	_In_z_ LPCWSTR lpszMid)
{
	// FID/MID lookups only succeed online, so go ahead and force it
	RegKeys[regKeyMAPI_NO_CACHE].ulCurDWORD = true;
	RegKeys[regkeyMDB_ONLINE].ulCurDWORD = true;

	DebugPrint(DBGGeneric, "DumpFidMid: Outputting from profile %ws. FID: %ws, MID: %ws\n", lpszProfile, lpszFid, lpszMid);
	HRESULT hRes = S_OK;
	LPMAPIFOLDER lpFolder = NULL;

	if (lpMDB)
	{
		// Open root container.
		WC_H(CallOpenEntry(
			lpMDB,
			NULL,
			NULL,
			NULL,
			NULL, // open root container
			NULL,
			MAPI_BEST_ACCESS,
			NULL,
			(LPUNKNOWN*)&lpFolder));
	}

	if (lpFolder)
	{
		CFindFidMid MyFindFidMid;
		MyFindFidMid.InitMDB(lpMDB);
		MyFindFidMid.InitFolder(lpFolder);
		MyFindFidMid.InitFidMid(lpszFid, lpszMid);
		MyFindFidMid.ProcessFolders(
			true,
			true,
			true);
	}

	if (lpFolder) lpFolder->Release();
}

void DoFidMid(_In_ MYOPTIONS ProgOpts)
{
	DumpFidMid(
		ProgOpts.lpszProfile.c_str(),
		ProgOpts.lpMDB,
		ProgOpts.lpszFid.c_str(),
		ProgOpts.lpszMid.c_str());
}