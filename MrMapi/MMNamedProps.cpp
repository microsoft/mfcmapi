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

template<typename... Arguments>
void WriteOutput(FILE* fOut, LPCWSTR szMsg, Arguments&&... args)
{
	auto szString = strings::format(szMsg, std::forward<Arguments>(args)...);
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

class GuidHasher
{
public:
	size_t operator()(const GUID& guid) const
	{
		// TODO: Figure out how not to take a depdendency on RPC
		RPC_STATUS status;
		std::uint16_t hash = UuidHash(const_cast<UUID*>(&guid), &status);
		return std::hash<std::uint16_t>()(hash);
	}
};

class GuidLess
{
public:
	bool operator()(const GUID& lhs, const GUID& rhs) const
	{
		return memcmp(&lhs, &rhs, sizeof(GUID)) < 0;
	}
};

void AnalyzeNamedProps(std::vector<std::shared_ptr<cache::namedPropCacheEntry>> const& nameToNumberMap)
{
	const long namedPropsQuota = 16384;
	const std::uint32_t NamedPropsLeakPatternCheckPercentageThreshold = 30;
	const std::uint32_t NamedPropsLeakPatternGuidPercentageThreshold = 20;
	const std::uint32_t NamedPropsLeakPatternGuidPercentageThresholdForInternetHeaders = 60;
	const std::uint32_t NamedPropsLeakPatternNamePrefixLength = 10;
	const std::uint32_t NamedPropsLeakPatternNamePrefixPercentageThreshold = 10;

	std::unordered_set<GUID, GuidHasher> leakPatternGuids;
	std::unordered_set<std::wstring> leakPatternNamePrefixes;

	// Simple protection from some bogus application which cause named property exhaustion.
	// This protection is not bullet-proof, it just attempts to recognize two known bogus patterns.
	// If bogus application doesn't follow any of those two patterns (e.g. just registers random names with random guids)
	// then no luck for such mailbox - it will hit the quota and should be dealt with manually.
	if constexpr (true || nameToNumberMap.size() > namedPropsQuota / 100 * NamedPropsLeakPatternCheckPercentageThreshold)
	{
		std::map<GUID, int, GuidLess> guidCounts;
		std::map<std::wstring, int> namePrefixCounts;

		// Loop through all properties and discovery leaky GUID patterns and leaky prefixes
		for (auto &entry : nameToNumberMap)
		{
			auto registeredPropName = entry->getMapiNameId();
			bool isInternetHeadersNamespace = memcmp(registeredPropName->lpguid, &guid::PS_INTERNET_HEADERS, sizeof(GUID)) == 0;

			// Check for a GUID pattern. We consider abnormal if we see more than
			// 2K names within the same GUID namespace. We allow larger limit for InternetHeaders
			// because some mailboxes migrated from E14 have more properties in this namespace.
			int count = guidCounts[*registeredPropName->lpguid];
			guidCounts[*registeredPropName->lpguid] = ++count;

			if ((isInternetHeadersNamespace && count > namedPropsQuota / 100 * NamedPropsLeakPatternGuidPercentageThresholdForInternetHeaders)
				|| (false == isInternetHeadersNamespace && count > namedPropsQuota / 100 * NamedPropsLeakPatternGuidPercentageThreshold))
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

				count = namePrefixCounts[namePrefix];
				namePrefixCounts[namePrefix] = ++count;

				if (count > namedPropsQuota / 100 * NamedPropsLeakPatternNamePrefixPercentageThreshold)
				{
					leakPatternNamePrefixes.emplace(namePrefix);
				}
			}
		}
	}

	wprintf(L"Detected %zu leaky GUIDs\n", leakPatternGuids.size());
	for (auto &guid : leakPatternGuids)
	{
		auto guidString = guid::GUIDToStringAndName(guid);
		wprintf(L"%ws\n", guidString.c_str());
	}

	wprintf(L"Detected %zu leaky Prefixes\n", leakPatternNamePrefixes.size());
	for (auto &prefix : leakPatternNamePrefixes)
	{
		wprintf(L"%ws\n", prefix.c_str());
	}
}

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
	auto now = std::chrono::system_clock::to_time_t(std::chrono::system_clock::now());
	WriteOutput(fOut, L"Timestamp: %hs\n", ctime(&now));
	WriteOutput(fOut, L"Profile Name: %ws\n", mapi::GetProfileName(lpSession).c_str());
	LPSPropValue lpStoreEntryId = nullptr;
	HRESULT result = HrGetOneProp(lpMDB, PR_STORE_ENTRYID, &lpStoreEntryId);
	if (SUCCEEDED(result))
	{
		auto block = smartview::InterpretBinary(lpStoreEntryId->Value.bin, parserType::ENTRYID, lpMDB);
		auto storeEntryId = strings::StripCarriage(block->toString());
		WriteOutput(fOut, L"%ws\n\n", storeEntryId.c_str());

		MAPIFreeBuffer(lpStoreEntryId);
	}

	auto names = cache::GetNamesFromIDs(lpMDB, nullptr, 0);

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

	AnalyzeNamedProps(names);

	wprintf(L"Completed.\n");
	if (fOut) fclose(fOut);
}