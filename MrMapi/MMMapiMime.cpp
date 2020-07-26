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

	wprintf(L"Message File Converter\n");
	wprintf(L"Options specified:\n");
	wprintf(L"   Input File: %ws\n", input.c_str());
	wprintf(L"   Output File: %ws\n", output.c_str());
	wprintf(L"   Conversion Type: ");
	if (cli::switchMAPI.isSet())
	{
		wprintf(L"MAPI -> MIME\n");

		wprintf(L"   Save Format: %ws\n", cli::switchRFC822.isSet() ? L"RFC822" : L"RFC1521");

		if (cli::switchWrap.isSet())
		{
			wprintf(L"   Line Wrap: ");
			if (ulWrapLines == 0)
				wprintf(L"OFF\n");
			else
				wprintf(L"%lu\n", ulWrapLines);
		}
	}
	else if (cli::switchMAPI.isSet())
	{
		wprintf(L"MIME -> MAPI\n");
		if (cli::switchUnicode.isSet())
		{
			wprintf(L"   Building Unicode MSG file\n");
		}

		if (cli::switchCharset.isSet())
		{
			wprintf(L"   CodePage: %lu\n", ulCodePage);
			wprintf(L"   CharSetType: ");
			switch (cSetType)
			{
			case CHARSET_BODY:
				wprintf(L"CHARSET_BODY");
				break;
			case CHARSET_HEADER:
				wprintf(L"CHARSET_HEADER");
				break;
			case CHARSET_WEB:
				wprintf(L"CHARSET_WEB");
				break;
			}

			wprintf(L"\n");
			wprintf(L"   CharSetApplyType: ");
			switch (cSetApplyType)
			{
			case CSET_APPLY_UNTAGGED:
				wprintf(L"CSET_APPLY_UNTAGGED");
				break;
			case CSET_APPLY_ALL:
				wprintf(L"CSET_APPLY_ALL");
				break;
			case CSET_APPLY_TAG_ALL:
				wprintf(L"CSET_APPLY_TAG_ALL");
				break;
			}

			wprintf(L"\n");
		}
	}

	if (!cli::switchCCSFFlags.empty())
	{
		auto szFlags = flags::InterpretFlags(flagCcsf, convertFlags);
		if (!szFlags.empty())
		{
			wprintf(L"   Conversion Flags: %ws\n", szFlags.c_str());
		}
	}

	if (cli::switchEncoding.isSet())
	{
		auto szType = flags::InterpretFlags(flagIet, ulEncodingType);
		if (!szType.empty())
		{
			wprintf(L"   Encoding Type: %ws\n", szType.c_str());
		}
	}

	if (cli::switchAddressBook.isSet())
	{
		wprintf(L"   Using Address Book\n");
	}

	auto hRes = S_OK;

	LPADRBOOK lpAdrBook = nullptr;
	if (cli::switchAddressBook.isSet() && lpMAPISession)
	{
		hRes = WC_MAPI(lpMAPISession->OpenAddressBook(NULL, nullptr, AB_NO_DIALOG, &lpAdrBook));
		if (FAILED(hRes)) wprintf(L"OpenAddressBook returned an error: 0x%08lx\n", hRes);
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
				wprintf(L"MimeOleGetCodePageCharset returned 0x%08lX\n", hRes);
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
		wprintf(L"File converted successfully.\n");
	}
	else if (REGDB_E_CLASSNOTREG == hRes)
	{
		wprintf(L"MAPI <-> MIME converter not found. Perhaps Outlook is not installed.\n");
	}
	else
	{
		wprintf(L"Conversion returned an error: 0x%08lx\n", hRes);
	}

	if (lpAdrBook) lpAdrBook->Release();
}