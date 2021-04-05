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

	wprintf(L"Completed.\n");
	if (fOut) fclose(fOut);
}