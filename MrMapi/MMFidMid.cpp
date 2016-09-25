#include "stdafx.h"
#include "MrMAPI.h"
#include "MMFidMid.h"
#include "../MAPIFunctions.h"
#include "../MapiProcessor.h"
#include "../ExtraPropTags.h"
#include "../SmartView/SmartView.h"
#include "../String.h"

void PrintFolder(wstring szFid, wstring szFolder)
{
	printf("%-15ws %ws\n", szFid.c_str(), szFolder.c_str());
}

void PrintMessage(LPCWSTR szMid, bool fAssociated, wstring szSubject, wstring szClass)
{
	wprintf(L" %-15ws %wc %ws (%ws)\n", szMid, fAssociated ? L'A' : L'R', szSubject.c_str(), szClass.c_str());
}

class CFindFidMid : public CMAPIProcessor
{
public:
	CFindFidMid();
	void InitFidMid(wstring szFid, wstring szMid, bool bMid);

private:
	bool ContinueProcessingFolders() override;
	bool ShouldProcessContentsTable() override;
	void BeginFolderWork() override;
	void BeginContentsTableWork(ULONG ulFlags, ULONG ulCountRows) override;
	bool DoContentsTablePerRowWork(_In_ LPSRow lpSRow, ULONG ulCurRow) override;

	wstring m_szFid;
	wstring m_szMid;
	wstring m_szCurrentFid;

	bool m_bMid;
	bool m_fFIDMatch;
	bool m_fFIDExactMatch;
	bool m_fFIDPrinted;
	bool m_fAssociated;
};

CFindFidMid::CFindFidMid()
{
	m_bMid = false;
	m_fFIDMatch = false;
	m_fFIDExactMatch = false;
	m_fFIDPrinted = false;
	m_fAssociated = false;
}

void CFindFidMid::InitFidMid(wstring szFid, wstring szMid, bool bMid)
{
	m_szFid = szFid;
	m_szMid = szMid;
	m_bMid = bMid;
}

// --------------------------------------------------------------------------------- //

// Passed in Fid matches the found Fid exactly, or matches the tail exactly
// For instance, both 3-15632 and 15632 will match against an input fid of 15632
bool MatchFid(wstring inputFid, wstring currentFid)
{
	if (_wcsicmp(inputFid.c_str(), currentFid.c_str()) == 0) return true;

	auto pos = currentFid.find('-');
	if (pos == string::npos) return false;
	auto trimmedFid = currentFid.substr(pos + 1, string::npos);
	if (_wcsicmp(inputFid.c_str(), trimmedFid.c_str()) == 0)
	{
		return true;
	}

	return false;
}

void CFindFidMid::BeginFolderWork()
{
	DebugPrint(DBGGeneric, L"CFindFidMid::BeginFolderWork: m_szFolderOffset %ws\n", m_szFolderOffset.c_str());
	m_fFIDMatch = false;
	m_fFIDExactMatch = false;
	m_fFIDPrinted = false;
	m_szCurrentFid.clear();
	if (!m_lpFolder) return;

	auto hRes = S_OK;
	ULONG ulProps = NULL;
	LPSPropValue lpProps = nullptr;

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
		LPSPropTagArray(&sptaFolderProps),
		NULL,
		&ulProps,
		&lpProps));
	DebugPrint(DBGGeneric, L"CFindFidMid::DoFolderPerHierarchyTableRowWork: m_szFolderOffset %ws\n", m_szFolderOffset.c_str());

	// Now get the FID
	LPWSTR lpszDisplayName = nullptr;

	if (lpProps && PR_DISPLAY_NAME_W == lpProps[ePR_DISPLAY_NAME_W].ulPropTag)
	{
		lpszDisplayName = lpProps[ePR_DISPLAY_NAME_W].Value.lpszW;
	}

	if (lpProps && PidTagFolderId == lpProps[ePidTagFolderId].ulPropTag)
	{
		m_szCurrentFid = FidMidToSzString(lpProps[ePidTagFolderId].Value.li.QuadPart, false);
		DebugPrint(DBGGeneric, L"CFindFidMid::DoFolderPerHierarchyTableRowWork: Found FID %ws for %ws\n", m_szCurrentFid.c_str(), lpszDisplayName);
	}
	else
	{
		// Nothing left to do if we can't find a fid.
		DebugPrint(DBGGeneric, L"CFindFidMid::DoFolderPerHierarchyTableRowWork: No FID found for %ws\n", lpszDisplayName);
		return;
	}

	// If FidMidToSzString failed, we're done.
	if (m_szCurrentFid.empty()) return;

	// Check for FID matches - no fid matches all folders
	if (m_szFid.empty())
	{
		m_fFIDMatch = true;
		m_fFIDExactMatch = false;
	}
	else if (MatchFid(m_szFid, m_szCurrentFid))
	{
		m_fFIDMatch = true;
		m_fFIDExactMatch = true;
	}

	if (m_fFIDMatch)
	{
		DebugPrint(DBGGeneric, L"CFindFidMid::DoFolderPerHierarchyTableRowWork: Matched FID %ws\n", m_szFid.c_str());
		// Print out the folder
		if (m_szMid.empty() || m_fFIDExactMatch)
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
	// 1 - We matched our fid, possibly because a fid wasn't passed in
	// 2 - We are in mid mode
	return m_fFIDMatch && m_bMid;
}

void CFindFidMid::BeginContentsTableWork(ULONG ulFlags, ULONG /*ulCountRows*/)
{
	m_fAssociated = (ulFlags & MAPI_ASSOCIATED) == MAPI_ASSOCIATED;
}

bool CFindFidMid::DoContentsTablePerRowWork(_In_ LPSRow lpSRow, ULONG /*ulCurRow*/)
{
	if (!lpSRow) return false;

	wstring lpszThisMid;
	wstring lpszSubject;
	wstring lpszClass;

	auto lpPropMid = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PidTagMid);
	if (lpPropMid)
	{
		lpszThisMid = FidMidToSzString(lpPropMid->Value.li.QuadPart, false);
		DebugPrint(DBGGeneric, L"CFindFidMid::DoContentsTablePerRowWork: Found MID %ws\n", lpszThisMid);
	}
	else
	{
		// Nothing left to do if we can't find a mid
		DebugPrint(DBGGeneric, L"CFindFidMid::DoContentsTablePerRowWork: No MID found\n");
		return false;
	}

	// Check for a MID match
	if (m_szMid.empty() || MatchFid(m_szMid, lpszThisMid))
	{
		// If we haven't already, print the folder info
		if (!m_fFIDPrinted)
		{
			PrintFolder(m_szCurrentFid, m_szFolderOffset);
			m_fFIDPrinted = true;
		}

		auto lpPropSubject = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_SUBJECT);
		if (lpPropSubject)
		{
			lpszSubject = LPCTSTRToWstring(lpPropSubject->Value.LPSZ);
		}

		auto lpPropClass = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_MESSAGE_CLASS);
		if (lpPropClass)
		{
			lpszClass = LPCTSTRToWstring(lpPropClass->Value.LPSZ);
		}

		PrintMessage(lpszThisMid.c_str(), m_fAssociated, lpszSubject, lpszClass);
		DebugPrint(DBGGeneric, L"EnumMessages::ProcessRow: Matched MID %ws, \"%ws\", \"%ws\"\n", lpszThisMid.c_str(), lpszSubject.c_str(), lpszClass.c_str());
	}

	// Don't open the message - the row is enough
	return false;
}

void DumpFidMid(
	_In_z_ LPCWSTR lpszProfile,
	_In_ LPMDB lpMDB,
	wstring lpszFid,
	wstring lpszMid,
	bool bMid)
{
	// FID/MID lookups only succeed online, so go ahead and force it
	RegKeys[regKeyMAPI_NO_CACHE].ulCurDWORD = true;
	RegKeys[regkeyMDB_ONLINE].ulCurDWORD = true;

	DebugPrint(DBGGeneric, L"DumpFidMid: Outputting from profile %ws. FID: %ws, MID: %ws\n", lpszProfile, lpszFid.c_str(), lpszMid.c_str());
	auto hRes = S_OK;
	LPMAPIFOLDER lpFolder = nullptr;

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
			reinterpret_cast<LPUNKNOWN*>(&lpFolder)));
	}

	if (lpFolder)
	{
		CFindFidMid MyFindFidMid;
		MyFindFidMid.InitMDB(lpMDB);
		MyFindFidMid.InitFolder(lpFolder);
		MyFindFidMid.InitFidMid(lpszFid, lpszMid, bMid);
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
		ProgOpts.lpszMid.c_str(),
		OPT_MID == (ProgOpts.ulOptions & OPT_MID));
}