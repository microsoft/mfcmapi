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

void PrintGuid(const GUID& lpGuid)
{
	wprintf(
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

void DoNamedProps(_In_opt_ LPMDB lpMDB)
{
	if (!lpMDB) return;
	registry::cacheNamedProps = false;

	wprintf(L"Dumping named properties...\n");
	wprintf(L"\n");

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
		wprintf(L"[%08x] (", name->getPropID());
		const auto nameId = name->getMapiNameId();
		PrintGuid(*nameId->lpguid);
		switch (nameId->ulKind)
		{
		case MNID_ID:
			wprintf(L":0x%08x)\n", nameId->Kind.lID);
			break;

		case MNID_STRING:
			wprintf(L":%ws)\n", nameId->Kind.lpwstrName);
			break;
		}
	}
}