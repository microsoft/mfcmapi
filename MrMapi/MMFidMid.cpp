#include <StdAfx.h>
#include <MrMapi/MMFidMid.h>
#include <MrMapi/mmcli.h>
#include <core/mapi/processor/mapiProcessor.h>
#include <core/mapi/extraPropTags.h>
#include <core/smartview/SmartView.h>
#include <core/utility/strings.h>
#include <core/utility/registry.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFunctions.h>

namespace mapiprocessor
{
	void PrintFolder(const std::wstring& szFid, const std::wstring& szFolder)
	{
		wprintf(L"%-15ws %ws\n", szFid.c_str(), szFolder.c_str());
	}

	void PrintMessage(
		const std::wstring& szMid,
		bool fAssociated,
		const std::wstring& szSubject,
		const std::wstring& szClass)
	{
		wprintf(
			L" %-15ws %wc %ws (%ws)\n", szMid.c_str(), fAssociated ? L'A' : L'R', szSubject.c_str(), szClass.c_str());
	}

	class CFindFidMid : public mapi::processor::mapiProcessor
	{
	public:
		CFindFidMid();
		void InitFidMid(const std::wstring& szFid, const std::wstring& szMid, bool bMid);

	private:
		bool ContinueProcessingFolders() override;
		bool ShouldProcessContentsTable() override;
		void BeginFolderWork() override;
		void BeginContentsTableWork(ULONG ulFlags, ULONG ulCountRows) override;
		bool DoContentsTablePerRowWork(_In_ const _SRow* lpSRow, ULONG ulCurRow) override;

		std::wstring m_szFid;
		std::wstring m_szMid;
		std::wstring m_szCurrentFid;

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

	void CFindFidMid::InitFidMid(const std::wstring& szFid, const std::wstring& szMid, bool bMid)
	{
		m_szFid = szFid;
		m_szMid = szMid;
		m_bMid = bMid;
	}

	// --------------------------------------------------------------------------------- //

	// Passed in Fid matches the found Fid exactly, or matches the tail exactly
	// For instance, both 3-15632 and 15632 will match against an input fid of 15632
	bool MatchFid(const std::wstring& inputFid, const std::wstring& currentFid)
	{
		if (_wcsicmp(inputFid.c_str(), currentFid.c_str()) == 0) return true;

		const auto pos = currentFid.find('-');
		if (pos == std::string::npos) return false;
		auto trimmedFid = currentFid.substr(pos + 1, std::string::npos);
		return _wcsicmp(inputFid.c_str(), trimmedFid.c_str()) == 0;
	}

	void CFindFidMid::BeginFolderWork()
	{
		output::DebugPrint(
			DBGGeneric, L"CFindFidMid::BeginFolderWork: m_szFolderOffset %ws\n", m_szFolderOffset.c_str());
		m_fFIDMatch = false;
		m_fFIDExactMatch = false;
		m_fFIDPrinted = false;
		m_szCurrentFid.clear();
		if (!m_lpFolder) return;

		ULONG ulProps = NULL;
		LPSPropValue lpProps = nullptr;

		enum
		{
			ePR_DISPLAY_NAME_W,
			ePidTagFolderId,
			NUM_COLS
		};
		static const SizedSPropTagArray(NUM_COLS, sptaFolderProps) = {NUM_COLS, {PR_DISPLAY_NAME_W, PidTagFolderId}};

		WC_H_GETPROPS_S(m_lpFolder->GetProps(LPSPropTagArray(&sptaFolderProps), NULL, &ulProps, &lpProps));
		output::DebugPrint(
			DBGGeneric,
			L"CFindFidMid::DoFolderPerHierarchyTableRowWork: m_szFolderOffset %ws\n",
			m_szFolderOffset.c_str());

		// Now get the FID
		LPWSTR lpszDisplayName = nullptr;

		if (lpProps && PR_DISPLAY_NAME_W == lpProps[ePR_DISPLAY_NAME_W].ulPropTag)
		{
			lpszDisplayName = lpProps[ePR_DISPLAY_NAME_W].Value.lpszW;
		}

		if (lpProps && PidTagFolderId == lpProps[ePidTagFolderId].ulPropTag)
		{
			m_szCurrentFid = smartview::FidMidToSzString(lpProps[ePidTagFolderId].Value.li.QuadPart, false);
			output::DebugPrint(
				DBGGeneric,
				L"CFindFidMid::DoFolderPerHierarchyTableRowWork: Found FID %ws for %ws\n",
				m_szCurrentFid.c_str(),
				lpszDisplayName);
		}
		else
		{
			// Nothing left to do if we can't find a fid.
			output::DebugPrint(
				DBGGeneric, L"CFindFidMid::DoFolderPerHierarchyTableRowWork: No FID found for %ws\n", lpszDisplayName);
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
			output::DebugPrint(
				DBGGeneric, L"CFindFidMid::DoFolderPerHierarchyTableRowWork: Matched FID %ws\n", m_szFid.c_str());
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

	bool CFindFidMid::DoContentsTablePerRowWork(_In_ const _SRow* lpSRow, ULONG /*ulCurRow*/)
	{
		if (!lpSRow) return false;

		std::wstring lpszThisMid;
		std::wstring lpszSubject;
		std::wstring lpszClass;

		const auto lpPropMid = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PidTagMid);
		if (lpPropMid)
		{
			lpszThisMid = smartview::FidMidToSzString(lpPropMid->Value.li.QuadPart, false);
			output::DebugPrint(
				DBGGeneric, L"CFindFidMid::DoContentsTablePerRowWork: Found MID %ws\n", lpszThisMid.c_str());
		}
		else
		{
			// Nothing left to do if we can't find a mid
			output::DebugPrint(DBGGeneric, L"CFindFidMid::DoContentsTablePerRowWork: No MID found\n");
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

			const auto lpPropSubject = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_SUBJECT);
			if (lpPropSubject)
			{
				lpszSubject = strings::LPCTSTRToWstring(lpPropSubject->Value.LPSZ);
			}

			const auto lpPropClass = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_MESSAGE_CLASS);
			if (lpPropClass)
			{
				lpszClass = strings::LPCTSTRToWstring(lpPropClass->Value.LPSZ);
			}

			PrintMessage(lpszThisMid, m_fAssociated, lpszSubject, lpszClass);
			output::DebugPrint(
				DBGGeneric,
				L"EnumMessages::ProcessRow: Matched MID %ws, \"%ws\", \"%ws\"\n",
				lpszThisMid.c_str(),
				lpszSubject.c_str(),
				lpszClass.c_str());
		}

		// Don't open the message - the row is enough
		return false;
	}

	void DumpFidMid(
		_In_ const std::wstring& lpszProfile,
		_In_ LPMDB lpMDB,
		_In_ const std::wstring& lpszFid,
		_In_ const std::wstring& lpszMid,
		bool bMid)
	{
		// FID/MID lookups only succeed online, so go ahead and force it
		registry::forceMapiNoCache = true;
		registry::forceMDBOnline = true;

		output::DebugPrint(
			DBGGeneric,
			L"DumpFidMid: Outputting from profile %ws. FID: %ws, MID: %ws\n",
			lpszProfile.c_str(),
			lpszFid.c_str(),
			lpszMid.c_str());

		if (lpMDB)
		{
			// Open root container.
			auto lpFolder = mapi::CallOpenEntry<LPMAPIFOLDER>(
				lpMDB,
				nullptr,
				nullptr,
				nullptr,
				nullptr, // open root container
				nullptr,
				MAPI_BEST_ACCESS,
				nullptr);
			if (lpFolder)
			{
				CFindFidMid MyFindFidMid;
				MyFindFidMid.InitMDB(lpMDB);
				MyFindFidMid.InitFolder(lpFolder);
				MyFindFidMid.InitFidMid(lpszFid, lpszMid, bMid);
				MyFindFidMid.ProcessFolders(true, true, true);
				lpFolder->Release();
			}
		}
	}
} // namespace mapiprocessor

void DoFidMid(_In_ LPMDB lpMDB)
{
	mapiprocessor::DumpFidMid(
		cli::switchProfile.at(0),
		lpMDB,
		cli::switchFid.at(0),
		cli::switchMid.at(0),
		cli::switchMid.isSet());
}