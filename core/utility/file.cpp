#include <core/stdafx.h>
#include <core/utility/file.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/utility/error.h>
#include <ShlObj.h>

namespace file
{
	std::wstring ShortenPath(const std::wstring& path)
	{
		if (!path.empty())
		{
			WCHAR szShortPath[MAX_PATH] = {};
			// Use the short path to give us as much room as possible
			WC_D_S(GetShortPathNameW(path.c_str(), szShortPath, _countof(szShortPath)));
			std::wstring ret = szShortPath;
			if (ret.back() != L'\\')
			{
				ret += L'\\';
			}

			return ret;
		}

		return path;
	}

	std::wstring GetDirectoryPath(HWND hWnd)
	{
		WCHAR szPath[MAX_PATH] = {0};
		BROWSEINFOW BrowseInfo = {};
		LPMALLOC lpMalloc = nullptr;

		EC_H_S(SHGetMalloc(&lpMalloc));

		if (!lpMalloc) return strings::emptystring;

		auto szInputString = strings::loadstring(IDS_DIRPROMPT);

		BrowseInfo.hwndOwner = hWnd;
		BrowseInfo.lpszTitle = szInputString.c_str();
		BrowseInfo.pszDisplayName = szPath;
		BrowseInfo.ulFlags = BIF_USENEWUI | BIF_RETURNONLYFSDIRS;

		// Note - I don't initialize COM for this call because MAPIInitialize does this
		const auto lpItemIdList = SHBrowseForFolderW(&BrowseInfo);
		if (lpItemIdList)
		{
			EC_B_S(SHGetPathFromIDListW(lpItemIdList, szPath));
			lpMalloc->Free(lpItemIdList);
		}

		lpMalloc->Release();

		auto path = ShortenPath(szPath);
		if (path.length() >= MAXMSGPATH)
		{
			error::ErrDialog(__FILE__, __LINE__, IDS_EDPATHTOOLONG, path.length(), MAXMSGPATH);
			return strings::emptystring;
		}

		return path;
	}

	// The file name we generate should be shorter than MAX_PATH
	// This includes our directory name too!
	std::wstring BuildFileNameAndPath(
		_In_ const std::wstring& szExt,
		_In_ const std::wstring& szSubj,
		_In_ const std::wstring& szRootPath,
		_In_opt_ const _SBinary* lpBin)
	{
		output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath ext = \"%ws\"\n", szExt.c_str());
		output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath subj = \"%ws\"\n", szSubj.c_str());
		output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath rootPath = \"%ws\"\n", szRootPath.c_str());

		// Set up the path portion of the output.
		auto cleanRoot = ShortenPath(szRootPath);
		output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath cleanRoot = \"%ws\"\n", cleanRoot.c_str());

		// If we don't have enough space for even the shortest filename, give up.
		if (cleanRoot.length() >= MAXMSGPATH) return strings::emptystring;

		// Work with a max path which allows us to add our extension.
		// Shrink the max path to allow for a -ATTACHxxx extension.
		const auto maxFile = MAX_PATH - cleanRoot.length() - szExt.length() - MAXATTACH;

		std::wstring cleanSubj;
		if (!szSubj.empty())
		{
			cleanSubj = strings::SanitizeFileName(szSubj);
		}
		else
		{
			// We must have failed to get a subject before. Make one up.
			cleanSubj = L"UnknownSubject"; // STRING_OK
		}

		output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath cleanSubj = \"%ws\"\n", cleanSubj.c_str());

		std::wstring szBin;
		if (lpBin && lpBin->cb)
		{
			szBin = L"_" + strings::BinToHexString(lpBin, false);
		}

		if (cleanSubj.length() + szBin.length() <= maxFile)
		{
			auto szFile = cleanRoot + cleanSubj + szBin + szExt;
			output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath fileOut= \"%ws\"\n", szFile.c_str());
			return szFile;
		}

		// We couldn't build the string we wanted, so try something shorter
		// Compute a shorter subject length that should fit.
		const auto maxSubj = maxFile - min(MAXBIN, szBin.length()) - 1;
		auto szFile = cleanSubj.substr(0, maxSubj) + szBin.substr(0, MAXBIN);
		output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath shorter file = \"%ws\"\n", szFile.c_str());
		output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath new length = %d\n", szFile.length());

		if (szFile.length() >= maxFile)
		{
			szFile = cleanSubj.substr(0, MAXSUBJTIGHT) + szBin.substr(0, MAXBIN);
			output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath shorter file = \"%ws\"\n", szFile.c_str());
			output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath new length = %d\n", szFile.length());
		}

		if (szFile.length() >= maxFile)
		{
			output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath failed to build a string\n");
			return strings::emptystring;
		}

		auto szOut = cleanRoot + szFile + szExt;
		output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath fileOut= \"%ws\"\n", szOut.c_str());
		return szOut;
	}

	std::wstring GetModuleFileName(_In_opt_ HMODULE hModule)
	{
		auto buf = std::vector<wchar_t>();
		auto copied = DWORD();
		do
		{
			buf.resize(buf.size() + MAX_PATH);
			copied = EC_D(DWORD, ::GetModuleFileNameW(hModule, &buf.at(0), static_cast<DWORD>(buf.size())));
		} while (copied >= buf.size());

		buf.resize(copied);

		const auto path = std::wstring(buf.begin(), buf.end());

		return path;
	}

	std::wstring GetSystemDirectory()
	{
		auto buf = std::vector<wchar_t>();
		auto copied = DWORD();
		do
		{
			buf.resize(buf.size() + MAX_PATH);
			copied = EC_D(DWORD, ::GetSystemDirectoryW(&buf.at(0), static_cast<UINT>(buf.size())));
		} while (copied >= buf.size());

		buf.resize(copied);

		const auto path = std::wstring(buf.begin(), buf.end());

		return path;
	}

	std::map<std::wstring, std::wstring> GetFileVersionInfo(_In_opt_ HMODULE hModule)
	{
		auto versionStrings = std::map<std::wstring, std::wstring>();
		const auto szFullPath = file::GetModuleFileName(hModule);
		if (!szFullPath.empty())
		{
			auto dwVerInfoSize = EC_D(DWORD, GetFileVersionInfoSizeW(szFullPath.c_str(), nullptr));
			if (dwVerInfoSize)
			{
				// If we were able to get the information, process it.
				const auto pbData = new BYTE[dwVerInfoSize];
				if (pbData == nullptr) return {};

				auto hRes =
					EC_B(::GetFileVersionInfoW(szFullPath.c_str(), NULL, dwVerInfoSize, static_cast<void*>(pbData)));

				if (SUCCEEDED(hRes))
				{
					struct LANGANDCODEPAGE
					{
						WORD wLanguage;
						WORD wCodePage;
					}* lpTranslate = {nullptr};

					UINT cbTranslate = 0;

					// Read the list of languages and code pages.
					hRes = EC_B(VerQueryValueW(
						pbData,
						L"\\VarFileInfo\\Translation", // STRING_OK
						reinterpret_cast<LPVOID*>(&lpTranslate),
						&cbTranslate));

					// Read the file description for each language and code page.

					if (hRes == S_OK && lpTranslate)
					{
						for (UINT iCodePages = 0; iCodePages < cbTranslate / sizeof(LANGANDCODEPAGE); iCodePages++)
						{
							const auto szSubBlock = strings::format(
								L"\\StringFileInfo\\%04x%04x\\", // STRING_OK
								lpTranslate[iCodePages].wLanguage,
								lpTranslate[iCodePages].wCodePage);

							// Load all our strings
							for (auto iVerString = IDS_VER_FIRST; iVerString <= IDS_VER_LAST; iVerString++)
							{
								UINT cchVer = 0;
								wchar_t* lpszVer = nullptr;
								auto szVerString = strings::loadstring(iVerString);
								auto szQueryString = szSubBlock + szVerString;

								hRes = EC_B(VerQueryValueW(
									static_cast<void*>(pbData),
									szQueryString.c_str(),
									reinterpret_cast<void**>(&lpszVer),
									&cchVer));

								if (hRes == S_OK && cchVer && lpszVer)
								{
									versionStrings[szVerString] = lpszVer;
								}
							}
						}
					}
				}

				delete[] pbData;
			}
		}

		return versionStrings;
	}
} // namespace file
