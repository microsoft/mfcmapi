#include <core/stdafx.h>
#include <core/mapi/mapiOutput.h>
#include <core/utility/output.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>
#include <core/interpret/guid.h>
#include <core/interpret/proptype.h>
#include <core/interpret/proptags.h>

namespace output
{
	void outputBinary(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ const SBinary& bin)
	{
		CHKPARAM;
		EARLYABORT;

		Output(ulDbgLvl, fFile, false, strings::BinToHexString(&bin, true));

		Output(ulDbgLvl, fFile, false, L"\n");
	}

	void outputEntryList(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPENTRYLIST lpEntryList)
	{
		CHKPARAM;
		EARLYABORT;
		if (!lpEntryList) return;

		Outputf(ulDbgLvl, fFile, true, L"Dumping %u entry IDs.\n", lpEntryList->cValues);

		std::wstring szFlags;
		std::wstring szPropNum;

		for (ULONG i = 0; i < lpEntryList->cValues; i++)
		{
			Outputf(ulDbgLvl, fFile, true, L"lpEntryList->lpbin[%u]\n\t", i);
			outputBinary(ulDbgLvl, fFile, lpEntryList->lpbin[i]);
		}

		Outputf(ulDbgLvl, fFile, true, L"End dumping entry list.\n");
	}

	void outputFormPropArray(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPIFORMPROPARRAY lpMAPIFormPropArray)
	{
		Outputf(ulDbgLvl, fFile, true, L"\t0x%X Properties:\n", lpMAPIFormPropArray->cProps);
		for (ULONG i = 0; i < lpMAPIFormPropArray->cProps; i++)
		{
			Outputf(ulDbgLvl, fFile, true, L"\t\tProperty 0x%X\n", i);

			if (lpMAPIFormPropArray->aFormProp[i].ulFlags == MAPI_UNICODE)
			{
				Outputf(
					ulDbgLvl,
					fFile,
					true,
					L"\t\tProperty Name: %ws\n\t\tProperty Type: %ws\n\t\tSpecial Type: 0x%X\n\t\tNum Vals: 0x%X\n", // STRING_OK
					LPWSTR(lpMAPIFormPropArray->aFormProp[i].pszDisplayName),
					proptype::TypeToString(lpMAPIFormPropArray->aFormProp[i].nPropType).c_str(),
					lpMAPIFormPropArray->aFormProp[i].nSpecialType,
					lpMAPIFormPropArray->aFormProp[i].u.s1.cfpevAvailable);
			}
			else
			{
				Outputf(
					ulDbgLvl,
					fFile,
					true,
					L"\t\tProperty Name: %hs\n\t\tProperty Type: %ws\n\t\tSpecial Type: 0x%X\n\t\tNum Vals: 0x%X\n", // STRING_OK
					LPSTR(lpMAPIFormPropArray->aFormProp[i].pszDisplayName),
					proptype::TypeToString(lpMAPIFormPropArray->aFormProp[i].nPropType).c_str(),
					lpMAPIFormPropArray->aFormProp[i].nSpecialType,
					lpMAPIFormPropArray->aFormProp[i].u.s1.cfpevAvailable);
			}

			outputNamedPropID(ulDbgLvl, fFile, &lpMAPIFormPropArray->aFormProp[i].u.s1.nmidIdx);
			for (ULONG j = 0; j < lpMAPIFormPropArray->aFormProp[i].u.s1.cfpevAvailable; j++)
			{
				if (lpMAPIFormPropArray->aFormProp[i].ulFlags == MAPI_UNICODE)
				{
					Outputf(
						ulDbgLvl,
						fFile,
						true,
						L"\t\t\tEnum 0x%X\nEnumVal Name: %ws\t\t\t\nEnumVal enumeration: 0x%X\n", // STRING_OK
						j,
						LPWSTR(lpMAPIFormPropArray->aFormProp[i].u.s1.pfpevAvailable[j].pszDisplayName),
						lpMAPIFormPropArray->aFormProp[i].u.s1.pfpevAvailable[j].nVal);
				}
				else
				{
					Outputf(
						ulDbgLvl,
						fFile,
						true,
						L"\t\t\tEnum 0x%X\nEnumVal Name: %hs\t\t\t\nEnumVal enumeration: 0x%X\n", // STRING_OK
						j,
						LPSTR(lpMAPIFormPropArray->aFormProp[i].u.s1.pfpevAvailable[j].pszDisplayName),
						lpMAPIFormPropArray->aFormProp[i].u.s1.pfpevAvailable[j].nVal);
				}
			}

			outputNamedPropID(ulDbgLvl, fFile, &lpMAPIFormPropArray->aFormProp[i].nmid);
		}
	}

	void outputNamedPropID(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPINAMEID lpName)
	{
		CHKPARAM;
		EARLYABORT;

		if (!lpName) return;

		if (lpName->ulKind == MNID_ID)
		{
			Outputf(
				ulDbgLvl,
				fFile,
				true,
				L"\t\t: nmid ID: 0x%X\n", // STRING_OK
				lpName->Kind.lID);
		}
		else
		{
			Outputf(
				ulDbgLvl,
				fFile,
				true,
				L"\t\t: nmid Name: %ws\n", // STRING_OK
				lpName->Kind.lpwstrName);
		}

		Output(ulDbgLvl, fFile, false, guid::GUIDToStringAndName(lpName->lpguid));
		Output(ulDbgLvl, fFile, false, L"\n");
	}

	void outputPropTagArray(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSPropTagArray lpTagsToDump)
	{
		CHKPARAM;
		EARLYABORT;
		if (!lpTagsToDump) return;

		Outputf(
			ulDbgLvl,
			fFile,
			true,
			L"\tProp tag list, %u props\n", // STRING_OK
			lpTagsToDump->cValues);
		for (ULONG uCurProp = 0; uCurProp < lpTagsToDump->cValues; uCurProp++)
		{
			Outputf(
				ulDbgLvl,
				fFile,
				true,
				L"\t\tProp: %u = %ws\n", // STRING_OK
				uCurProp,
				proptags::TagToString(lpTagsToDump->aulPropTag[uCurProp], nullptr, false, true).c_str());
		}

		Output(ulDbgLvl, fFile, true, L"\tEnd Prop Tag List\n");
	}
} // namespace output
