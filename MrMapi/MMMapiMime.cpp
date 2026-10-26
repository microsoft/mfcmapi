#include <StdAfx.h>
#include <MrMapi/MMMapiMime.h>
#include <MrMapi/mmcli.h>
#include <core/mapi/mapiMime.h>
#include <core/utility/import.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

void DoMAPIMIME(_In_opt_ LPMAPISESSION lpMAPISession)
{
	const auto input = cli::switchInput[0];
	const auto output = cli::switchOutput[0];
	const auto ulWrapLines = cli::switchWrap.atULONG(0);
	const auto ulEncodingType = cli::switchEncoding.atULONG(0);
	const auto convertFlags = static_cast<CCSFLAGS>(cli::switchCCSFFlags.atULONG(0));

	const auto ulCodePage = cli::switchCharset.atULONG(0);
	const auto cSetType = static_cast<CHARSETTYPE>(cli::switchCharset.atULONG(1));
	const auto cSetApplyType = static_cast<CSETAPPLYTYPE>(cli::switchCharset.atULONG(2));

	wprintf_s(L"Message File Converter\n");
	wprintf_s(L"Options specified:\n");
	wprintf_s(L"   Input File: %ws\n", input.c_str());
	wprintf_s(L"   Output File: %ws\n", output.c_str());
	wprintf_s(L"   Conversion Type: ");
	if (cli::switchMAPI.isSet())
	{
		wprintf_s(L"MAPI -> MIME\n");

		wprintf_s(L"   Save Format: %ws\n", cli::switchRFC822.isSet() ? L"RFC822" : L"RFC1521");

		if (cli::switchWrap.isSet())
		{
			wprintf_s(L"   Line Wrap: ");
			if (ulWrapLines == 0)
				wprintf_s(L"OFF\n");
			else
				wprintf_s(L"%lu\n", ulWrapLines);
		}
	}
	else if (cli::switchMAPI.isSet())
	{
		wprintf_s(L"MIME -> MAPI\n");
		if (cli::switchUnicode.isSet())
		{
			wprintf_s(L"   Building Unicode MSG file\n");
		}

		if (cli::switchCharset.isSet())
		{
			wprintf_s(L"   CodePage: %lu\n", ulCodePage);
			wprintf_s(L"   CharSetType: ");
			switch (cSetType)
			{
			case CHARSET_BODY:
				wprintf_s(L"CHARSET_BODY");
				break;
			case CHARSET_HEADER:
				wprintf_s(L"CHARSET_HEADER");
				break;
			case CHARSET_WEB:
				wprintf_s(L"CHARSET_WEB");
				break;
			}

			wprintf_s(L"\n");
			wprintf_s(L"   CharSetApplyType: ");
			switch (cSetApplyType)
			{
			case CSET_APPLY_UNTAGGED:
				wprintf_s(L"CSET_APPLY_UNTAGGED");
				break;
			case CSET_APPLY_ALL:
				wprintf_s(L"CSET_APPLY_ALL");
				break;
			case CSET_APPLY_TAG_ALL:
				wprintf_s(L"CSET_APPLY_TAG_ALL");
				break;
			}

			wprintf_s(L"\n");
		}
	}

	if (!cli::switchCCSFFlags.empty())
	{
		auto szFlags = flags::InterpretFlags(flagCcsf, convertFlags);
		if (!szFlags.empty())
		{
			wprintf_s(L"   Conversion Flags: %ws\n", szFlags.c_str());
		}
	}

	if (cli::switchEncoding.isSet())
	{
		auto szType = flags::InterpretFlags(flagIet, ulEncodingType);
		if (!szType.empty())
		{
			wprintf_s(L"   Encoding Type: %ws\n", szType.c_str());
		}
	}

	if (cli::switchAddressBook.isSet())
	{
		wprintf_s(L"   Using Address Book\n");
	}

	auto hRes = S_OK;

	LPADRBOOK lpAdrBook = nullptr;
	if (cli::switchAddressBook.isSet() && lpMAPISession)
	{
		hRes = WC_MAPI(lpMAPISession->OpenAddressBook(NULL, nullptr, AB_NO_DIALOG, &lpAdrBook));
		if (FAILED(hRes)) wprintf_s(L"OpenAddressBook returned an error: 0x%08lx\n", hRes);
	}

	if (cli::switchMIME.isSet())
	{
		// Source file is MSG, target is EML
		hRes = WC_H(mapi::mapimime::ConvertMSGToEML(
			input.c_str(),
			output.c_str(),
			convertFlags,
			cli::switchEncoding.isSet() ? static_cast<ENCODINGTYPE>(ulEncodingType) : IET_UNKNOWN,
			cli::switchRFC822.isSet() ? SAVE_RFC822 : SAVE_RFC1521,
			cli::switchWrap.isSet() ? ulWrapLines : USE_DEFAULT_WRAPPING,
			lpAdrBook));
	}
	else if (cli::switchMAPI.isSet())
	{
		// Source file is EML, target is MSG
		HCHARSET hCharSet = nullptr;
		if (cli::switchCharset.isSet())
		{
			hRes = WC_H(import::MyMimeOleGetCodePageCharset(ulCodePage, cSetType, &hCharSet));
			if (FAILED(hRes))
			{
				wprintf_s(L"MimeOleGetCodePageCharset returned 0x%08lX\n", hRes);
			}
		}

		if (SUCCEEDED(hRes))
		{
			hRes = WC_H(mapi::mapimime::ConvertEMLToMSG(
				input.c_str(),
				output.c_str(),
				convertFlags,
				cli::switchCharset.isSet(),
				hCharSet,
				cSetApplyType,
				lpAdrBook,
				cli::switchUnicode.isSet()));
		}
	}

	if (SUCCEEDED(hRes))
	{
		wprintf_s(L"File converted successfully.\n");
	}
	else if (REGDB_E_CLASSNOTREG == hRes)
	{
		wprintf_s(L"MAPI <-> MIME converter not found. Perhaps Outlook is not installed.\n");
	}
	else
	{
		wprintf_s(L"Conversion returned an error: 0x%08lx\n", hRes);
	}

	if (lpAdrBook) lpAdrBook->Release();
}