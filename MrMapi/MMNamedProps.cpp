#include <StdAfx.h>
#include <MrMapi/MMNamedProps.h>
#include <MrMapi/mmcli.h>
#include <core/mapi/processor/mapiProcessor.h>
#include <core/mapi/extraPropTags.h>
#include <core/smartview/SmartView.h>
#include <core/utility/strings.h>
#include <core/utility/registry.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/cache/namedProps.h>
#include <core/addin/mfcmapi.h>
#include <chrono>
#include <unordered_set>
#include <core/interpret/guid.h>

template <typename... Arguments>
void WriteOutput(_In_opt_ FILE* fOut, _Printf_format_string_ LPCWSTR szMsg, Arguments&&... args)
{
	const auto szString = strings::format(szMsg, std::forward<Arguments>(args)...);
	if (fOut)
	{
		output::Output(output::dbgLevel::NoDebug, fOut, false, szString);
	}
	else
	{
		wprintf(L"%ws", szString.c_str());
	}
}

void PrintGuid(FILE* fOut, const GUID& lpGuid)
{
	WriteOutput(
		fOut,
		L"%.8lX-%.4X-%.4X-%.2X%.2X-%.2X%.2X%.2X%.2X%.2X%.2X",
		lpGuid.Data1,
		lpGuid.Data2,
		lpGuid.Data3,
		lpGuid.Data4[0],
		lpGuid.Data4[1],
		lpGuid.Data4[2],
		lpGuid.Data4[3],
		lpGuid.Data4[4],
		lpGuid.Data4[5],
		lpGuid.Data4[6],
		lpGuid.Data4[7]);
}

/// <summary>
/// Hash implementation for GUIDs so they can be stored in std::unordered_list
/// </summary>
class GuidHasher
{
public:
	size_t operator()(const GUID& guid) const
	{
		std::uint32_t hash =
			guid.Data1 ^ (((int) guid.Data2 << 16) | (int) guid.Data3) ^ (((int) guid.Data4[2] << 24) | guid.Data4[7]);
		return std::hash<std::uint32_t>()(hash);
	}
};

/// <summary>
/// Comparer class used for comparing two GUIDs (whether lhs < rhs).
/// Used with std::map where GUID is the key
/// </summary>
class GuidLess
{
public:
	bool operator()(const GUID& lhs, const GUID& rhs) const { return memcmp(&lhs, &rhs, sizeof(GUID)) < 0; }
};

/// <summary>
/// Analyzes named properties to discover patterns that indicate incorrect usage of named properties by applications.
/// </summary>
/// <param name="fOut">The output file pointer</param>
/// <param name="nameToNumberMap">The named property list</param>
void AnalyzeNamedProps(
	_In_opt_ FILE* fOut,
	std::vector<std::shared_ptr<cache::namedPropCacheEntry>> const& nameToNumberMap)
{
	const std::uint32_t namedPropsQuota = 16384;
	const std::uint32_t NamedPropsLeakPatternCheckPercentageThreshold = 30;
	const std::uint32_t NamedPropsLeakPatternGuidPercentageThreshold = 20;
	const std::uint32_t NamedPropsLeakPatternGuidPercentageThresholdForInternetHeaders = 60;
	const std::uint32_t NamedPropsLeakPatternNamePrefixLength = 10;
	const std::uint32_t NamedPropsLeakPatternNamePrefixPercentageThreshold = 10;

	std::unordered_set<GUID, GuidHasher> leakPatternGuids;
	std::unordered_set<std::wstring> leakPatternNamePrefixes;

	// Simple detection for some misbehaving applications which cause named property exhaustion.
	// This detection is not comprehensive, it just attempts to recognize two known patterns.
	// If the application doesn't follow any of those two patterns (e.g. just registers random names with random guids)
	// then we won't detect the pattern, and manual analysis will have to be done.
	if (nameToNumberMap.size() > namedPropsQuota / 100 * NamedPropsLeakPatternCheckPercentageThreshold)
	{
		std::map<GUID, int, GuidLess> guidCounts;
		std::map<std::wstring, int> namePrefixCounts;

		// Loop through all properties and discovery leaky GUID patterns and leaky prefixes
		for (const auto& entry : nameToNumberMap)
		{
			const auto registeredPropName = entry->getMapiNameId();
			const auto isInternetHeadersNamespace =
				memcmp(registeredPropName->lpguid, &guid::PS_INTERNET_HEADERS, sizeof(GUID)) == 0;

			// Check for a GUID pattern. We consider abnormal if we see more than
			// 2K names within the same GUID namespace. We allow larger limit for InternetHeaders
			// because some mailboxes migrated from older versions have more properties in this namespace.
			int count = ++guidCounts[*registeredPropName->lpguid];

			if ((isInternetHeadersNamespace &&
				 count > namedPropsQuota / 100 * NamedPropsLeakPatternGuidPercentageThresholdForInternetHeaders) ||
				(false == isInternetHeadersNamespace &&
				 count > namedPropsQuota / 100 * NamedPropsLeakPatternGuidPercentageThreshold))
			{
				leakPatternGuids.emplace(*registeredPropName->lpguid);
			}

			// If this is a string name and name is long enough, check for a name prefix pattern. We consider abnormal if we see more than
			// 1K names with the same 10-character prefix.
			if (registeredPropName->ulKind == MNID_STRING && registeredPropName->Kind.lpwstrName != nullptr)
			{
				std::wstring namePrefix = registeredPropName->Kind.lpwstrName;
				if (namePrefix.size() > NamedPropsLeakPatternNamePrefixLength)
				{
					namePrefix = namePrefix.substr(0, NamedPropsLeakPatternNamePrefixLength);
				}

				count = ++namePrefixCounts[namePrefix];
				if (count > namedPropsQuota / 100 * NamedPropsLeakPatternNamePrefixPercentageThreshold)
				{
					leakPatternNamePrefixes.emplace(namePrefix);
				}
			}
		}
	}

	// Print out leaky property-set GUIDs
	WriteOutput(fOut, L"\nDetected %zu property sets with a large number of properties.\n", leakPatternGuids.size());
	if (!leakPatternGuids.empty())
	{
		WriteOutput(
			fOut,
			L"This typically indicates the incorrect use of named properties by a custom application.\nCheck the full "
			L"list of properties in the property set(s) listed below to help identify the application at fault.\n");
	}

	for (const auto& guid : leakPatternGuids)
	{
		const auto guidString = guid::GUIDToStringAndName(guid);
		WriteOutput(fOut, L"%ws\n", guidString.c_str());
	}

	// Print out leaky prefixes
	WriteOutput(
		fOut,
		L"\nDetected %zu large sets of unique properties sharing a common prefix.\n",
		leakPatternNamePrefixes.size());
	if (!leakPatternNamePrefixes.empty())
	{
		WriteOutput(fOut, L"This typically indicates the incorrect use of named properties by a custom application.\n");
	}

	for (const auto& prefix : leakPatternNamePrefixes)
	{
		WriteOutput(fOut, L"%ws\n", prefix.c_str());
	}
}

/// <summary>
/// Performs the full named property analysis and output.
/// </summary>
/// <param name="lpSession">The MAPI session.</param>
/// <param name="lpMDB">The MAPI message store.</param>
void DoNamedProps(_In_opt_ LPMAPISESSION lpSession, _In_opt_ LPMDB lpMDB)
{
	if (!lpSession || !lpMDB) return;
	registry::cacheNamedProps = false;

	// Ignore the reg key that disables smart view parsing
	registry::doSmartView = true;

	// See if they have specified an output file
	FILE* fOut = nullptr;
	const auto output = cli::switchOutput[0];
	if (!output.empty())
	{
		fOut = output::MyOpenFileMode(output, L"wb");
		if (!fOut)
		{
			wprintf(L"Cannot open output file %ws\n", output.c_str());
			return exit(-1);
		}
	}

	wprintf(L"Dumping named properties...\n\n");

	// Print the name of the profile and the parsed store entry-id so we can tell which mailbox this is
	const auto now = std::chrono::system_clock::to_time_t(std::chrono::system_clock::now());
	WriteOutput(fOut, L"Timestamp: %hs\n", ctime(&now));
	WriteOutput(fOut, L"Profile Name: %ws\n", mapi::GetProfileName(lpSession).c_str());
	LPSPropValue lpStoreEntryId = nullptr;
	const auto result = WC_H(HrGetOneProp(lpMDB, PR_STORE_ENTRYID, &lpStoreEntryId));
	if (SUCCEEDED(result))
	{
		const auto block = smartview::InterpretBinary(lpStoreEntryId->Value.bin, parserType::ENTRYID, lpMDB);
		const auto storeEntryId = strings::StripCarriage(block->toString());
		WriteOutput(fOut, L"%ws\n\n", storeEntryId.c_str());

		MAPIFreeBuffer(lpStoreEntryId);
	}

	auto names = cache::GetNamesFromIDs(lpMDB, nullptr, 0);

	// Start by analyzing the properties
	AnalyzeNamedProps(fOut, names);

	// Now output all of them
	WriteOutput(fOut, L"\nProperty List:\n");
	std::sort(
		names.begin(),
		names.end(),
		[](std::shared_ptr<cache::namedPropCacheEntry>& i, std::shared_ptr<cache::namedPropCacheEntry>& j) {
			// If the Property set GUIDs don't match, sort by that
			const auto iNameId = i->getMapiNameId();
			const auto jNameId = j->getMapiNameId();
			const auto res = memcmp(iNameId->lpguid, jNameId->lpguid, sizeof(GUID));
			if (res) return res < 0;

			// If they are different kinds, use that next
			if (iNameId->ulKind != jNameId->ulKind)
			{
				return iNameId->ulKind < jNameId->ulKind;
			}

			switch (iNameId->ulKind)
			{
			case MNID_ID:
				return iNameId->Kind.lID < jNameId->Kind.lID;
			case MNID_STRING:
				return std::wstring(iNameId->Kind.lpwstrName) < std::wstring(jNameId->Kind.lpwstrName);
			}

			// This shouldn't ever actually hit
			return i->getPropID() < j->getPropID();
		});

	for (const auto& name : names)
	{
		WriteOutput(fOut, L"[%08X] (", PROP_TAG(PT_UNSPECIFIED, name->getPropID()));
		const auto nameId = name->getMapiNameId();
		PrintGuid(fOut, *nameId->lpguid);
		switch (nameId->ulKind)
		{
		case MNID_ID:
			WriteOutput(fOut, L":0x%04X)\n", nameId->Kind.lID);
			break;

		case MNID_STRING:
			WriteOutput(fOut, L":%ws)\n", nameId->Kind.lpwstrName);
			break;
		}
	}

	wprintf(L"Completed.\n");
	if (fOut) fclose(fOut);
}