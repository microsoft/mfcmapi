#include <core/stdafx.h>
#include <core/addin/addin.h>
#include <core/addin/mfcmapi.h>
#include <core/utility/import.h>
#include <mapistub/library/stubutils.h>
#include <core/utility/strings.h>
#include <core/utility/file.h>
#include <core/utility/registry.h>
#include <core/utility/output.h>
#include <core/utility/error.h>
#include <core/addin/addinfile.h>

// Our built in arrays, which get merged into the arrays declared in mfcmapi.h
#include <core/interpret/genTagArray.h>
#include <core/interpret/flagArray.h>
#include <core/interpret/guidArray.h>
#include <core/interpret/nameIDArray.h>
#include <core/interpret/propTypeArray.h>
#include <core/interpret/smartViewParsers.h>

std::vector<NAME_ARRAY_ENTRY_V2> PropTagArray;
std::vector<NAME_ARRAY_ENTRY> PropTypeArray;
std::vector<GUID_ARRAY_ENTRY> PropGuidArray;
std::vector<NAMEID_ARRAY_ENTRY> NameIDArray;
std::vector<FLAG_ARRAY_ENTRY> FlagArray;
std::vector<SMARTVIEW_PARSER_ARRAY_ENTRY> SmartViewParserArray;
std::vector<SMARTVIEW_PARSER_TYPE_ARRAY_ENTRY> SmartViewParserTypeArray;
std::vector<_AddIn> g_lpMyAddins;

namespace addin
{
	template <typename T> T GetFunction(HMODULE hMod, LPCSTR szFuncName)
	{
		return WC_D(T, reinterpret_cast<T>(GetProcAddress(hMod, szFuncName)));
	}

	_Check_return_ ULONG GetAddinVersion(HMODULE hMod)
	{
		const auto pfnGetAPIVersion = GetFunction<LPGETAPIVERSION>(hMod, szGetAPIVersion);
		if (pfnGetAPIVersion)
		{
			return pfnGetAPIVersion();
		}

		// Default case for unversioned add-ins
		return MFCMAPI_HEADER_V1;
	}

	// Load NAME_ARRAY_ENTRY style props from the add-in and upconvert them to V2
	void LoadLegacyPropTags(
		HMODULE hMod,
		_In_ ULONG* lpulPropTags, // Number of entries in lppPropTags
		_In_ LPNAME_ARRAY_ENTRY_V2* lppPropTags // Array of NAME_ARRAY_ENTRY_V2 structures
	)
	{
		const auto pfnGetPropTags = GetFunction<LPGETPROPTAGS>(hMod, szGetPropTags);
		if (pfnGetPropTags)
		{
			LPNAME_ARRAY_ENTRY lpPropTags = nullptr;
			pfnGetPropTags(lpulPropTags, &lpPropTags);
			if (lpPropTags && *lpulPropTags)
			{
				const auto lpPropTagsV2 = new (std::nothrow) NAME_ARRAY_ENTRY_V2[*lpulPropTags];
				if (lpPropTagsV2)
				{
					for (ULONG i = 0; i < *lpulPropTags; i++)
					{
						lpPropTagsV2[i].lpszName = lpPropTags[i].lpszName;
						lpPropTagsV2[i].ulValue = lpPropTags[i].ulValue;
						lpPropTagsV2[i].ulSortOrder = 0;
					}

					*lppPropTags = lpPropTagsV2;
				}
			}
		}
	}

	void LoadSingleAddIn(_In_ _AddIn& addIn, HMODULE hMod, _In_ LPLOADADDIN pfnLoadAddIn)
	{
		output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Loading AddIn\n");
		if (!pfnLoadAddIn) return;
		addIn.hMod = hMod;
		pfnLoadAddIn(&addIn.szName);
		if (addIn.szName)
		{
			output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Loading \"%ws\"\n", addIn.szName);
		}

		const auto ulVersion = GetAddinVersion(hMod);
		output::DebugPrint(output::dbgLevel::AddInPlumbing, L"AddIn version = %u\n", ulVersion);

		const auto pfnGetMenus = GetFunction<LPGETMENUS>(hMod, szGetMenus);
		if (pfnGetMenus)
		{
			pfnGetMenus(&addIn.ulMenu, &addIn.lpMenu);
			if (!addIn.ulMenu || !addIn.lpMenu)
			{
				output::DebugPrint(output::dbgLevel::AddInPlumbing, L"AddIn returned invalid menus\n");
				addIn.ulMenu = NULL;
				addIn.lpMenu = nullptr;
			}
			if (addIn.ulMenu && addIn.lpMenu)
			{
				for (ULONG ulMenu = 0; ulMenu < addIn.ulMenu; ulMenu++)
				{
					// Save off our add-in struct
					addIn.lpMenu[ulMenu].lpAddIn = &addIn;
					if (addIn.lpMenu[ulMenu].szMenu)
						output::DebugPrint(
							output::dbgLevel::AddInPlumbing, L"Menu: %ws\n", addIn.lpMenu[ulMenu].szMenu);
					if (addIn.lpMenu[ulMenu].szHelp)
						output::DebugPrint(
							output::dbgLevel::AddInPlumbing, L"Help: %ws\n", addIn.lpMenu[ulMenu].szHelp);
					output::DebugPrint(output::dbgLevel::AddInPlumbing, L"ID: 0x%08X\n", addIn.lpMenu[ulMenu].ulID);
					output::DebugPrint(
						output::dbgLevel::AddInPlumbing, L"Context: 0x%08X\n", addIn.lpMenu[ulMenu].ulContext);
					output::DebugPrint(
						output::dbgLevel::AddInPlumbing, L"Flags: 0x%08X\n", addIn.lpMenu[ulMenu].ulFlags);
				}
			}
		}

		addIn.bLegacyPropTags = false;
		const auto pfnGetPropTagsV2 = GetFunction<LPGETPROPTAGSV2>(hMod, szGetPropTagsV2);
		if (pfnGetPropTagsV2)
		{
			pfnGetPropTagsV2(&addIn.ulPropTags, &addIn.lpPropTags);
		}
		else
		{
			LoadLegacyPropTags(hMod, &addIn.ulPropTags, &addIn.lpPropTags);
			addIn.bLegacyPropTags = true;
		}

		const auto pfnGetPropTypes = GetFunction<LPGETPROPTYPES>(hMod, szGetPropTypes);
		if (pfnGetPropTypes)
		{
			pfnGetPropTypes(&addIn.ulPropTypes, &addIn.lpPropTypes);
		}

		const auto pfnGetPropGuids = GetFunction<LPGETPROPGUIDS>(hMod, szGetPropGuids);
		if (pfnGetPropGuids)
		{
			pfnGetPropGuids(&addIn.ulPropGuids, &addIn.lpPropGuids);
		}

		// v2 changed the LPNAMEID_ARRAY_ENTRY structure
		if (MFCMAPI_HEADER_V2 <= ulVersion)
		{
			const auto pfnGetNameIDs = GetFunction<LPGETNAMEIDS>(hMod, szGetNameIDs);
			if (pfnGetNameIDs)
			{
				pfnGetNameIDs(&addIn.ulNameIDs, &addIn.lpNameIDs);
			}
		}

		const auto pfnGetPropFlags = GetFunction<LPGETPROPFLAGS>(hMod, szGetPropFlags);
		if (pfnGetPropFlags)
		{
			pfnGetPropFlags(&addIn.ulPropFlags, &addIn.lpPropFlags);
		}

		const auto pfnGetSmartViewParserArray = GetFunction<LPGETSMARTVIEWPARSERARRAY>(hMod, szGetSmartViewParserArray);
		if (pfnGetSmartViewParserArray)
		{
			pfnGetSmartViewParserArray(&addIn.ulSmartViewParsers, &addIn.lpSmartViewParsers);
		}

		const auto pfnGetSmartViewParserTypeArray =
			GetFunction<LPGETSMARTVIEWPARSERTYPEARRAY>(hMod, szGetSmartViewParserTypeArray);
		if (pfnGetSmartViewParserTypeArray)
		{
			pfnGetSmartViewParserTypeArray(&addIn.ulSmartViewParserTypes, &addIn.lpSmartViewParserTypes);
		}

		output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Done loading AddIn\n");
	}

	namespace
	{
		constexpr auto APPROVAL_LIST = L"AddInApprovalList"; // STRING_OK
		constexpr auto DENIAL_LIST = L"AddInDenialList"; // STRING_OK
		constexpr auto LIST_SEPARATOR = L";"; // STRING_OK

		void ReadList(_In_ HKEY rootKey, _In_ const std::wstring& valueName, _Out_ std::vector<std::wstring>& values)
		{
			const auto serialized = registry::ReadStringFromRegistry(rootKey, valueName);
			size_t start = 0;
			while (start < serialized.size())
			{
				const auto end = serialized.find(L';', start);
				const auto length = end == std::wstring::npos ? serialized.size() - start : end - start;
				if (length) values.emplace_back(serialized.substr(start, length));
				if (end == std::wstring::npos) break;
				start = end + 1;
			}
		}

		class CApprovalCache final
		{
		public:
			CApprovalCache() : m_rootKey(registry::CreateRootKey())
			{
				if (m_rootKey)
				{
					ReadList(m_rootKey, APPROVAL_LIST, m_approved);
					ReadList(m_rootKey, DENIAL_LIST, m_denied);
				}
			}

			~CApprovalCache()
			{
				if (m_rootKey) EC_W32_S(RegCloseKey(m_rootKey));
			}

			bool IsApproved(_In_ const std::wstring& hash) const { return Contains(m_approved, hash); }
			bool IsDenied(_In_ const std::wstring& hash) const { return Contains(m_denied, hash); }

			void Add(_In_ const std::wstring& hash, bool approved)
			{
				auto& target = approved ? m_approved : m_denied;
				auto& other = approved ? m_denied : m_approved;
				if (!Contains(target, hash)) target.push_back(hash);
				other.erase(
					std::remove_if(
						other.begin(),
						other.end(),
						[&hash](const auto& value) { return strings::compareInsensitive(value, hash); }),
					other.end());
				m_dirty = true;
			}

			_Check_return_ bool Save()
			{
				if (!m_dirty) return true;
				if (!m_rootKey) return false;
				return WriteList(APPROVAL_LIST, m_approved) && WriteList(DENIAL_LIST, m_denied);
			}

		private:
			static bool Contains(const std::vector<std::wstring>& values, const std::wstring& hash)
			{
				return std::any_of(values.cbegin(), values.cend(), [&hash](const auto& value) {
					return strings::compareInsensitive(value, hash);
				});
			}

			bool WriteList(const std::wstring& valueName, const std::vector<std::wstring>& values)
			{
				if (values.empty())
				{
					const auto result = RegDeleteValueW(m_rootKey, valueName.c_str());
					if (result == ERROR_SUCCESS || result == ERROR_FILE_NOT_FOUND) return true;
					output::DebugPrint(
						output::dbgLevel::Generic,
						L"RegDeleteValueW(%ws) failed with %u\n",
						valueName.c_str(),
						static_cast<DWORD>(result));
					return false;
				}

				std::wstring serialized;
				for (const auto& value : values)
				{
					serialized += value;
					serialized += LIST_SEPARATOR;
				}

				if (serialized.size() >= MAXDWORD / sizeof(WCHAR))
				{
					output::DebugPrint(output::dbgLevel::Generic, L"Add-in decision cache is too large to store\n");
					return false;
				}

				const auto cbValue = static_cast<DWORD>((serialized.size() + 1) * sizeof(WCHAR));
				const auto result = RegSetValueExW(
					m_rootKey,
					valueName.c_str(),
					0,
					REG_SZ,
					reinterpret_cast<const BYTE*>(serialized.c_str()),
					cbValue);
				if (result == ERROR_SUCCESS) return true;
				output::DebugPrint(
					output::dbgLevel::Generic,
					L"RegSetValueExW(%ws) failed with %u\n",
					valueName.c_str(),
					static_cast<DWORD>(result));
				return false;
			}

			HKEY m_rootKey{};
			std::vector<std::wstring> m_approved;
			std::vector<std::wstring> m_denied;
			bool m_dirty{};
		};

		struct CPendingAddIn
		{
			AddInApprovalCandidate approval;
			std::wstring path;
			std::unique_ptr<CAddInFile> file;
		};

		void LoadApprovedAddIn(_In_ const std::wstring& path)
		{
			const auto flags = LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_DEFAULT_DIRS;
			const auto hMod = LoadLibraryExW(path.c_str(), nullptr, flags);
			if (!hMod)
			{
				output::DebugPrint(
					output::dbgLevel::AddInPlumbing,
					L"LoadLibraryExW(\"%ws\") failed with %u\n",
					path.c_str(),
					GetLastError());
				return;
			}

			const auto pfnLoadAddIn = GetFunction<LPLOADADDIN>(hMod, szLoadAddIn);
			if (pfnLoadAddIn && GetAddinVersion(hMod) == MFCMAPI_HEADER_CURRENT_VERSION)
			{
				g_lpMyAddins.emplace_back();
				LoadSingleAddIn(g_lpMyAddins.back(), hMod, pfnLoadAddIn);
			}
			else
			{
				output::DebugPrint(
					output::dbgLevel::AddInPlumbing, L"Approved DLL no longer has a compatible add-in export\n");
				FreeLibrary(hMod);
			}
		}
	} // namespace

	_Check_return_ bool ClearAddInApprovalCache()
	{
		const auto rootKey = registry::CreateRootKey();
		if (!rootKey) return false;

		bool success = true;
		for (const auto valueName : {APPROVAL_LIST, DENIAL_LIST})
		{
			const auto result = RegDeleteValueW(rootKey, valueName);
			if (result != ERROR_SUCCESS && result != ERROR_FILE_NOT_FOUND)
			{
				output::DebugPrint(
					output::dbgLevel::Generic,
					L"RegDeleteValueW(%ws) failed with %u\n",
					valueName,
					static_cast<DWORD>(result));
				success = false;
			}
		}

		EC_W32_S(RegCloseKey(rootKey));
		return success;
	}

	DWORD ValidateAddInDirectory(const std::wstring& directory)
	{
		if (directory.empty()) return ERROR_SUCCESS;
		if (directory.size() > MAX_PATH - 7) return ERROR_FILENAME_EXCED_RANGE;
		if (directory.size() < 3 ||
			!((directory[0] >= L'A' && directory[0] <= L'Z') || (directory[0] >= L'a' && directory[0] <= L'z')) ||
			directory[1] != L':' || directory[2] != L'\\')
			return ERROR_BAD_PATHNAME;
		if (directory.find_first_of(L"<>\"|?*/:", 2) != std::wstring::npos ||
			std::any_of(directory.begin(), directory.end(), [](wchar_t ch) { return ch < L' '; }))
			return ERROR_BAD_PATHNAME;

		for (size_t start = 3; start < directory.size();)
		{
			const auto end = directory.find(L'\\', start);
			const auto length = (end == std::wstring::npos ? directory.size() : end) - start;
			if (!length || directory[start + length - 1] == L'.' || directory[start + length - 1] == L' ')
				return ERROR_BAD_PATHNAME;
			if (end == std::wstring::npos) break;
			start = end + 1;
		}

		const auto root = directory.substr(0, 3);
		if (GetDriveTypeW(root.c_str()) != DRIVE_FIXED) return ERROR_NOT_SUPPORTED;

		const auto handle = CreateFileW(
			directory.c_str(),
			FILE_LIST_DIRECTORY | FILE_READ_ATTRIBUTES,
			FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
			nullptr,
			OPEN_EXISTING,
			FILE_FLAG_BACKUP_SEMANTICS,
			nullptr);
		if (handle == INVALID_HANDLE_VALUE) return GetLastError();

		auto result = static_cast<DWORD>(ERROR_SUCCESS);
		BY_HANDLE_FILE_INFORMATION info{};
		WCHAR resolved[MAX_PATH + 4]{};
		if (!GetFileInformationByHandle(handle, &info))
			result = GetLastError();
		else if (!(info.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
			result = ERROR_DIRECTORY;
		else
		{
			const auto length = GetFinalPathNameByHandleW(handle, resolved, _countof(resolved), FILE_NAME_NORMALIZED);
			if (!length)
				result = GetLastError();
			else if (length >= _countof(resolved))
				result = ERROR_FILENAME_EXCED_RANGE;
			else if (length < 7 || wcsncmp(resolved, L"\\\\?\\", 4) || resolved[5] != L':' || resolved[6] != L'\\')
				result = ERROR_NOT_SUPPORTED;
			else
			{
				const std::wstring resolvedRoot(resolved + 4, 3);
				if (GetDriveTypeW(resolvedRoot.c_str()) != DRIVE_FIXED) result = ERROR_NOT_SUPPORTED;
			}
		}

		WC_B_S(CloseHandle(handle));
		return result;
	}

	DWORD LoadAddIns(_In_ const AddInApprovalPrompt& approvalPrompt)
	{
		output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Loading AddIns\n");

		const std::wstring directory = registry::addInDirectory;
		auto result = ValidateAddInDirectory(directory);
		if (result != ERROR_SUCCESS)
		{
			output::DebugPrint(
				output::dbgLevel::AddInPlumbing,
				L"Invalid add-in directory \"%ws\" (error %u); add-ins will not be loaded\n",
				directory.c_str(),
				result);
			MergeAddInArrays();
			return result;
		}

		if (!directory.empty())
		{
			output::DebugPrint(
				output::dbgLevel::AddInPlumbing, L"Scanning add-in directory \"%ws\"\n", directory.c_str());
			const auto prefix = directory.back() == L'\\' ? directory : directory + L"\\";
			const auto spec = prefix + L"*.dll"; // STRING_OK
			WIN32_FIND_DATAW findData = {};
			const auto findHandle = FindFirstFileW(spec.c_str(), &findData);
			if (findHandle == INVALID_HANDLE_VALUE)
			{
				const auto findError = GetLastError();
				if (findError != ERROR_FILE_NOT_FOUND)
				{
					result = findError;
					output::DebugPrint(output::dbgLevel::AddInPlumbing, L"FindFirstFileW failed with %u\n", result);
				}
			}
			else
			{
				CApprovalCache cache;
				std::vector<CPendingAddIn> approved;
				std::vector<CPendingAddIn> unknown;
				auto findError = ERROR_NO_MORE_FILES;

				for (;;)
				{
					const auto path = prefix + findData.cFileName;
					auto candidateFile = std::make_unique<CAddInFile>();
					if (candidateFile->Open(path))
					{
						if (!candidateFile->HasCompatibleBitness())
						{
							output::DebugPrint(
								output::dbgLevel::AddInPlumbing,
								L"Skipping add-in with incompatible bitness \"%ws\"\n",
								path.c_str());
						}
						else
						{
							std::wstring hash;
							if (candidateFile->HasLoadAddInExport() && candidateFile->GetSHA256(hash))
							{
								candidateFile->CloseView();
								CPendingAddIn candidate{
									{findData.cFileName, std::move(hash)}, path, std::move(candidateFile)};
								if (!cache.IsDenied(candidate.approval.sha256) &&
									cache.IsApproved(candidate.approval.sha256))
								{
									approved.emplace_back(std::move(candidate));
								}
								else if (!cache.IsDenied(candidate.approval.sha256))
								{
									unknown.emplace_back(std::move(candidate));
								}
							}
							else
							{
								output::DebugPrint(
									output::dbgLevel::AddInPlumbing,
									L"Skipping non-add-in or unreadable image \"%ws\"\n",
									path.c_str());
							}
						}
					}
					else
					{
						output::DebugPrint(
							output::dbgLevel::AddInPlumbing,
							L"Could not inspect \"%ws\" (error %u)\n",
							path.c_str(),
							GetLastError());
					}

					if (!FindNextFileW(findHandle, &findData))
					{
						findError = GetLastError();
						break;
					}
				}

				FindClose(findHandle);
				if (findError != ERROR_NO_MORE_FILES)
				{
					result = findError;
					output::DebugPrint(output::dbgLevel::AddInPlumbing, L"FindNextFileW failed with %u\n", findError);
				}

				if (!unknown.empty() && approvalPrompt)
				{
					std::vector<AddInApprovalCandidate> promptCandidates;
					std::vector<size_t> decisionIndex;
					promptCandidates.reserve(unknown.size());
					decisionIndex.reserve(unknown.size());
					for (const auto& candidate : unknown)
					{
						const auto existing = std::find_if(
							promptCandidates.cbegin(),
							promptCandidates.cend(),
							[&candidate](const auto& promptCandidate) {
								return promptCandidate.sha256 == candidate.approval.sha256;
							});
						if (existing == promptCandidates.cend())
						{
							decisionIndex.push_back(promptCandidates.size());
							promptCandidates.push_back(candidate.approval);
						}
						else
						{
							decisionIndex.push_back(static_cast<size_t>(existing - promptCandidates.cbegin()));
						}
					}

					std::vector<bool> decisions;
					if (approvalPrompt(promptCandidates, decisions) && decisions.size() == promptCandidates.size())
					{
						for (size_t i = 0; i < promptCandidates.size(); ++i)
							cache.Add(promptCandidates[i].sha256, decisions[i]);

						if (!cache.Save())
						{
							output::DebugPrint(
								output::dbgLevel::AddInPlumbing,
								L"Could not persist add-in approval decisions; newly approved add-ins will not "
								L"load\n");
						}
						else
						{
							for (size_t i = 0; i < unknown.size(); ++i)
							{
								if (decisions[decisionIndex[i]]) approved.emplace_back(std::move(unknown[i]));
							}
						}
					}
					else
					{
						output::DebugPrint(
							output::dbgLevel::AddInPlumbing,
							L"Add-in approval was cancelled or returned invalid decisions\n");
					}
				}

				for (const auto& candidate : approved)
					LoadApprovedAddIn(candidate.path);
			}
		}
		else
		{
			output::DebugPrint(
				output::dbgLevel::AddInPlumbing, L"No add-in directory configured; bypassing add-in loading\n");
		}

		MergeAddInArrays();
		output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Done loading AddIns\n");
		return result;
	}

	void UnloadAddIns()
	{
		output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Unloading AddIns\n");

		for (const auto& addIn : g_lpMyAddins)
		{
			output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Freeing add-in\n");
			if (addIn.bLegacyPropTags)
			{
				delete[] addIn.lpPropTags;
			}

			if (addIn.hMod)
			{
				if (addIn.szName)
				{
					output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Unloading \"%ws\"\n", addIn.szName);
				}

				const auto pfnUnLoadAddIn = GetFunction<LPUNLOADADDIN>(addIn.hMod, szUnloadAddIn);
				if (pfnUnLoadAddIn) pfnUnLoadAddIn();

				FreeLibrary(addIn.hMod);
			}
		}

		output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Done unloading AddIns\n");
	}

	// Compare type arrays.
	int _cdecl compareTypes(_In_ void* /*pvlocale*/, _In_ const void* a1, _In_ const void* a2) noexcept
	{
		const auto lpType1 = static_cast<const NAME_ARRAY_ENTRY*>(a1);
		const auto lpType2 = static_cast<const NAME_ARRAY_ENTRY*>(a2);

		if (lpType1->ulValue > lpType2->ulValue) return 1;
		if (lpType1->ulValue == lpType2->ulValue)
		{
			return wcscmp(lpType1->lpszName, lpType2->lpszName);
		}

		return -1;
	}

	// Compare tag arrays. Pay no attention to sort order - we'll sort on sort order during output.
	int _cdecl compareTags(_In_ void* /*pvlocale*/, _In_ const void* a1, _In_ const void* a2) noexcept
	{
		const auto lpTag1 = static_cast<const NAME_ARRAY_ENTRY_V2*>(a1);
		const auto lpTag2 = static_cast<const NAME_ARRAY_ENTRY_V2*>(a2);

		const auto nameCmp = wcscmp(lpTag1->lpszName, lpTag2->lpszName);
		// If the names were the same, treat the entries as equal so we don't merge in two entries with the same name
		if (nameCmp == 0) return 0;

		if (lpTag1->ulValue > lpTag2->ulValue) return 1;
		if (lpTag1->ulValue == lpTag2->ulValue) return nameCmp;
		return -1;
	}

	int _cdecl compareNameID(_In_ void* /*pvlocale*/, _In_ const void* a1, _In_ const void* a2) noexcept
	{
		const auto lpID1 = static_cast<const NAMEID_ARRAY_ENTRY*>(a1);
		const auto lpID2 = static_cast<const NAMEID_ARRAY_ENTRY*>(a2);

		if (lpID1->lValue > lpID2->lValue) return 1;
		if (lpID1->lValue == lpID2->lValue)
		{
			const auto iCmp = wcscmp(lpID1->lpszName, lpID2->lpszName);
			if (iCmp) return iCmp;
			if (IsEqualGUID(*lpID1->lpGuid, *lpID2->lpGuid)) return 0;
		}

		return -1;
	}

	int _cdecl compareSmartViewParser(_In_ void* /*pvlocale*/, _In_ const void* a1, _In_ const void* a2) noexcept
	{
		const auto lpParser1 = static_cast<const SMARTVIEW_PARSER_ARRAY_ENTRY*>(a1);
		const auto lpParser2 = static_cast<const SMARTVIEW_PARSER_ARRAY_ENTRY*>(a2);

		if (lpParser1->ulIndex > lpParser2->ulIndex) return 1;
		if (lpParser1->ulIndex == lpParser2->ulIndex)
		{
			if (lpParser1->type > lpParser2->type) return 1;
			if (lpParser1->type == lpParser2->type)
			{
				if (lpParser1->bMV && !lpParser2->bMV) return 1;
				if (!lpParser1->bMV && lpParser2->bMV) return -1;
				return 0;
			}
		}

		return -1;
	}

	template <typename T>
	void MergeArrays(
		std::vector<T>& Target,
		_Inout_bytecap_x_(cSource* width) T* Source,
		_In_ size_t cSource,
		_In_ _CoreCrtSecureSearchSortCompareFunction Comparison)
	{
		// Sort the source array
		qsort_s(Source, cSource, sizeof(T), Comparison, nullptr);

		// Append any entries in the source not already in the target to the target
		for (ULONG i = 0; i < cSource; i++)
		{
			if (end(Target) == find_if(begin(Target), end(Target), [&](T& entry) {
					return Comparison(nullptr, &Source[i], &entry) == 0;
				}))
			{
				Target.push_back(Source[i]);
			}
		}

		// Stable sort the resulting array
		std::stable_sort(begin(Target), end(Target), [Comparison](const T& a, const T& b) -> bool {
			return Comparison(nullptr, &a, &b) < 0;
		});
	}

	// Flags are difficult to sort since we need to have a stable sort
	// Records with the same key must appear in the output in the same order as the input
	// qsort doesn't guarantee this, so we do it manually with an insertion sort
	void SortFlagArray(_In_count_(ulFlags) LPFLAG_ARRAY_ENTRY lpFlags, _In_ size_t ulFlags) noexcept
	{
		for (ULONG i = 1; i < ulFlags; i++)
		{
			const auto NextItem = lpFlags[i];
			auto iLoc = ULONG{};
			for (iLoc = i; iLoc > 0; iLoc--)
			{
				if (lpFlags[iLoc - 1].ulFlagName <= NextItem.ulFlagName) break;
				lpFlags[iLoc] = lpFlags[iLoc - 1];
			}

			lpFlags[iLoc] = NextItem;
		}
	}

	// Consults the end of the target array to find a match to the source
	// If no dupe is found, copies source to target
	void AppendFlagIfNotDupe(std::vector<FLAG_ARRAY_ENTRY>& target, FLAG_ARRAY_ENTRY source)
	{
		auto iTarget = target.size();

		while (iTarget)
		{
			iTarget--;
			// Stop searching when ulFlagName doesn't match
			// Assumes lpTarget is sorted
			if (target[iTarget].ulFlagName != source.ulFlagName) break;
			if (target[iTarget].lFlagValue == source.lFlagValue && target[iTarget].ulFlagType == source.ulFlagType &&
				strings::compareInsensitive(target[iTarget].lpszName, source.lpszName))
			{
				return;
			}
		}

		target.push_back(source);
	}

	// Similar to MergeArrays, but using AppendFlagIfNotDupe logic
	void MergeFlagArrays(std::vector<FLAG_ARRAY_ENTRY>& In1, _In_count_(cIn2) LPFLAG_ARRAY_ENTRY In2, _In_ size_t cIn2)
	{
		if (!In2) return;

		std::vector<FLAG_ARRAY_ENTRY> Out;

		size_t iIn1 = 0;
		size_t iIn2 = 0;

		while (iIn1 < In1.size() && iIn2 < cIn2)
		{
			// Add from In1 first, then In2 when In2 is bigger than In2
			if (In1[iIn1].ulFlagName <= In2[iIn2].ulFlagName)
			{
				AppendFlagIfNotDupe(Out, In1[iIn1++]);
			}
			else
			{
				AppendFlagIfNotDupe(Out, In2[iIn2++]);
			}
		}

		while (iIn1 < In1.size())
		{
			AppendFlagIfNotDupe(Out, In1[iIn1++]);
		}

		while (iIn2 < cIn2)
		{
			AppendFlagIfNotDupe(Out, In2[iIn2++]);
		}

		In1 = Out;
	}

	// Assumes built in arrays are already sorted!
	void MergeAddInArrays()
	{
		output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Loading default arrays\n");
		PropTagArray = std::vector<NAME_ARRAY_ENTRY_V2>(std::begin(g_PropTagArray), std::end(g_PropTagArray));
		PropTypeArray = std::vector<NAME_ARRAY_ENTRY>(std::begin(g_PropTypeArray), std::end(g_PropTypeArray));
		PropGuidArray =
			std::vector<GUID_ARRAY_ENTRY>(std::begin(guid::g_PropGuidArray), std::end(guid::g_PropGuidArray));
		NameIDArray = std::vector<NAMEID_ARRAY_ENTRY>(std::begin(g_NameIDArray), std::end(g_NameIDArray));
		FlagArray = std::vector<FLAG_ARRAY_ENTRY>(std::begin(g_FlagArray), std::end(g_FlagArray));
		SmartViewParserArray = std::vector<SMARTVIEW_PARSER_ARRAY_ENTRY>(
			std::begin(smartview::g_SmartViewParserArray), std::end(smartview::g_SmartViewParserArray));
		SmartViewParserTypeArray = std::vector<SMARTVIEW_PARSER_TYPE_ARRAY_ENTRY>(
			std::begin(smartview::g_SmartViewParserTypeArray), std::end(smartview::g_SmartViewParserTypeArray));

		output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Found 0x%08X built in prop tags.\n", PropTagArray.size());
		output::DebugPrint(
			output::dbgLevel::AddInPlumbing, L"Found 0x%08X built in prop types.\n", PropTypeArray.size());
		output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Found 0x%08X built in guids.\n", PropGuidArray.size());
		output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Found 0x%08X built in named ids.\n", NameIDArray.size());
		output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Found 0x%08X built in flags.\n", FlagArray.size());
		output::DebugPrint(
			output::dbgLevel::AddInPlumbing,
			L"Found 0x%08X built in Smart View parsers.\n",
			SmartViewParserArray.size());
		output::DebugPrint(
			output::dbgLevel::AddInPlumbing,
			L"Found 0x%08X built in Smart View parser types.\n",
			SmartViewParserTypeArray.size());

		// No add-in == nothing to merge
		if (g_lpMyAddins.empty()) return;

		output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Merging Add-In arrays\n");
		for (const auto& addIn : g_lpMyAddins)
		{
			output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Looking at %ws\n", addIn.szName);
			output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Found 0x%08X prop tags.\n", addIn.ulPropTags);
			output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Found 0x%08X prop types.\n", addIn.ulPropTypes);
			output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Found 0x%08X guids.\n", addIn.ulPropGuids);
			output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Found 0x%08X named ids.\n", addIn.ulNameIDs);
			output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Found 0x%08X flags.\n", addIn.ulPropFlags);
			output::DebugPrint(
				output::dbgLevel::AddInPlumbing, L"Found 0x%08X Smart View parsers.\n", addIn.ulSmartViewParsers);
			output::DebugPrint(
				output::dbgLevel::AddInPlumbing,
				L"Found 0x%08X Smart View parser types.\n",
				addIn.ulSmartViewParserTypes);

			if (addIn.ulPropTypes)
			{
				MergeArrays<NAME_ARRAY_ENTRY>(PropTypeArray, addIn.lpPropTypes, addIn.ulPropTypes, compareTypes);
			}

			if (addIn.ulPropTags)
			{
				MergeArrays<NAME_ARRAY_ENTRY_V2>(PropTagArray, addIn.lpPropTags, addIn.ulPropTags, compareTags);
			}

			if (addIn.ulNameIDs)
			{
				MergeArrays<NAMEID_ARRAY_ENTRY>(NameIDArray, addIn.lpNameIDs, addIn.ulNameIDs, compareNameID);
			}

			if (addIn.ulPropFlags)
			{
				SortFlagArray(addIn.lpPropFlags, addIn.ulPropFlags);
				MergeFlagArrays(FlagArray, addIn.lpPropFlags, addIn.ulPropFlags);
			}

			if (addIn.ulSmartViewParsers)
			{
				MergeArrays<SMARTVIEW_PARSER_ARRAY_ENTRY>(
					SmartViewParserArray, addIn.lpSmartViewParsers, addIn.ulSmartViewParsers, compareSmartViewParser);
			}

			// We add our new parsers to the end of the array, assigning ids starting with parserType::END
			static auto s_nextParser = static_cast<int>(parserType::END);
			if (addIn.ulSmartViewParserTypes)
			{
				for (ULONG i = 0; i < addIn.ulSmartViewParserTypes; i++)
				{
					SMARTVIEW_PARSER_TYPE_ARRAY_ENTRY addinType{};
					addinType.type = static_cast<parserType>(s_nextParser);
					addinType.lpszName = addIn.lpSmartViewParserTypes[i];
					SmartViewParserTypeArray.push_back(addinType);
					s_nextParser++;
				}
			}

			if (addIn.ulPropGuids)
			{
				// Copy guids from addIn.lpPropGuids, checking for dupes on the way
				for (ULONG i = 0; i < addIn.ulPropGuids; i++)
				{
					auto bDupe = false;
					// Since this array isn't sorted, we have to compare against all valid entries for dupes
					for (const auto& guid : PropGuidArray)
					{
						if (IsEqualGUID(*addIn.lpPropGuids[i].lpGuid, *guid.lpGuid))
						{
							bDupe = true;
							break;
						}
					}

					if (!bDupe)
					{
						PropGuidArray.push_back(addIn.lpPropGuids[i]);
					}
				}
			}
		}

		output::DebugPrint(output::dbgLevel::AddInPlumbing, L"After merge, 0x%08X prop tags.\n", PropTagArray.size());
		output::DebugPrint(output::dbgLevel::AddInPlumbing, L"After merge, 0x%08X prop types.\n", PropTypeArray.size());
		output::DebugPrint(output::dbgLevel::AddInPlumbing, L"After merge, 0x%08X guids.\n", PropGuidArray.size());
		output::DebugPrint(output::dbgLevel::AddInPlumbing, L"After merge, 0x%08X flags.\n", FlagArray.size());
		output::DebugPrint(
			output::dbgLevel::AddInPlumbing, L"After merge, 0x%08X Smart View parsers.\n", SmartViewParserArray.size());
		output::DebugPrint(
			output::dbgLevel::AddInPlumbing,
			L"After merge, 0x%08X Smart View parser types.\n",
			SmartViewParserTypeArray.size());

		output::DebugPrint(output::dbgLevel::AddInPlumbing, L"Done merging add-in arrays\n");
	}

	__declspec(dllexport) void __cdecl AddInLog(bool bPrintThreadTime, _Printf_format_string_ LPWSTR szMsg, ...)
	{
		if (!fIsSet(output::dbgLevel::AddIn)) return;

		auto argList = va_list{};
		va_start(argList, szMsg);
		const auto szAddInLogString = strings::formatV(szMsg, argList);
		va_end(argList);

		output::Output(output::dbgLevel::AddIn, nullptr, bPrintThreadTime, szAddInLogString);
	}

	__declspec(dllexport) void __cdecl GetMAPIModule(_In_ HMODULE* lphModule, bool bForce)
	{
		if (!lphModule) return;
		*lphModule = mapistub::GetMAPIHandle();
		if (!*lphModule && bForce)
		{
			// No MAPI loaded - load it
			*lphModule = mapistub::GetPrivateMAPI();
		}
	}

	std::wstring AddInStructTypeToString(parserType parser)
	{
		for (const auto& smartViewParserType : SmartViewParserTypeArray)
		{
			if (smartViewParserType.type == parser)
			{
				return smartViewParserType.lpszName;
			}
		}

		return L"";
	}

	std::wstring AddInSmartView(parserType iStructType, ULONG cbBin, _In_count_(cbBin) const BYTE* lpBin)
	{
		// Don't let add-ins hijack our built in types
		if (iStructType < parserType::END) return L"";

		auto szStructType = AddInStructTypeToString(iStructType);
		if (szStructType.empty()) return L"";

		for (const auto& addIn : g_lpMyAddins)
		{
			if (addIn.ulSmartViewParserTypes)
			{
				for (ULONG i = 0; i < addIn.ulSmartViewParserTypes; i++)
				{
					if (szStructType == addIn.lpSmartViewParserTypes[i])
					{
						const auto pfnSmartViewParse = GetFunction<LPSMARTVIEWPARSE>(addIn.hMod, szSmartViewParse);
						const auto pfnFreeParse = GetFunction<LPFREEPARSE>(addIn.hMod, szFreeParse);

						if (pfnSmartViewParse && pfnFreeParse)
						{
							const auto szParse = pfnSmartViewParse(szStructType.c_str(), cbBin, lpBin);
							std::wstring szRet = szParse;

							pfnFreeParse(szParse);
							return szRet;
						}
					}
				}
			}
		}

		return L"No parser found";
	}
} // namespace addin