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

class NamedPropEntry
{
public:
	NamedPropEntry(ULONG propTag, LPGUID lpGuid, LPWSTR name)
		: Kind{MNID_STRING}, PropTag{propTag}, PropSetGuid{*lpGuid}, PropName{name}
	{
	}

	NamedPropEntry(ULONG propTag, LPGUID lpGuid, ULONG id)
		: Kind{MNID_ID}, PropTag{propTag}, PropSetGuid{*lpGuid}, PropId{id}
	{
	}

	ULONG PropTag{0};
	GUID PropSetGuid{0};
	ULONG Kind{0};
	std::wstring PropName;
	ULONG PropId{0};
};

void DoNamedProps(_In_opt_ LPMDB lpMDB)
{
	auto hRes = S_OK;
	wprintf(L"Dumping named properties...\n");

	ULONG cPropNames = 0;
	LPSPropTagArray pProps = nullptr;
	LPMAPINAMEID* pNames = nullptr;
	hRes = WC_MAPI(lpMDB->GetNamesFromIDs(&pProps, nullptr, 0, &cPropNames, &pNames));

	if (SUCCEEDED(hRes))
	{
		wprintf(L"\n");
		std::vector<NamedPropEntry> props;
		for (ULONG i = 0; i < cPropNames; i++)
		{
			switch (pNames[i]->ulKind)
			{
			case MNID_STRING:
				props.emplace_back(pProps->aulPropTag[i], pNames[i]->lpguid, pNames[i]->Kind.lpwstrName);
				break;
			case MNID_ID:
				props.emplace_back(pProps->aulPropTag[i], pNames[i]->lpguid, pNames[i]->Kind.lID);
				break;
			}
		}

		std::sort(props.begin(), props.end(), [](NamedPropEntry& i, NamedPropEntry& j) {
			// If the Property set GUIDs don't match, sort by that
			auto res = memcmp(&i.PropSetGuid, &j.PropSetGuid, sizeof(GUID));
			if (res) return res < 0;

			// If they are different kinds, use that next
			if (i.Kind != j.Kind)
			{
				return i.Kind < j.Kind;
			}

			switch (i.Kind)
			{
			case MNID_ID:
				return i.PropId < j.PropId;
			case MNID_STRING:
				return i.PropName < j.PropName;
			}

			// This shouldn't ever actually hit
			return i.PropTag < j.PropTag;
		});

		for (size_t i = 0; i < props.size(); i++)
		{
			wprintf(L"[%08x] (", props[i].PropTag);
			PrintGuid(props[i].PropSetGuid);
			switch (props[i].Kind)
			{
			case MNID_ID:
				wprintf(L":0x%08x)\n", props[i].PropId);
				break;

			case MNID_STRING:
				wprintf(L":%ws)\n", props[i].PropName.c_str());
				break;
			}
		}
	}
	else
	{
		wprintf(L"FAILED. Error 0x%x\n", hRes);
	}

	if (pProps) MAPIFreeBuffer(pProps);
	if (pNames) MAPIFreeBuffer(pNames);
}