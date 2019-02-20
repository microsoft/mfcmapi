#include <StdAfx.h>
#include <MrMapi/MMMapiMime.h>
#include <MrMapi/mmcli.h>
#include <core/mapi/mapiMime.h>
#include <core/utility/import.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

void DoMAPIMIME(_In_ cli::MYOPTIONS ProgOpts)
{
	const auto input = cli::switchInput.getArg(0);
	const auto output = cli::switchOutput.getArg(0);
	const auto ulWrapLines = cli::switchWrap.getArgAsULONG(0);
	const auto ulEncodingType = cli::switchEncoding.getArgAsULONG(0);
	const auto convertFlags = static_cast<CCSFLAGS>(cli::switchCCSFFlags.getArgAsULONG(0));

	const auto ulCodePage = cli::switchCharset.getArgAsULONG(0);
	const auto cSetType = static_cast<CHARSETTYPE>(cli::switchCharset.getArgAsULONG(1));
	const auto cSetApplyType = static_cast<CSETAPPLYTYPE>(cli::switchCharset.getArgAsULONG(2));

	printf("Message File Converter\n");
	printf("Options specified:\n");
	printf("   Input File: %ws\n", input.c_str());
	printf("   Output File: %ws\n", output.c_str());
	printf("   Conversion Type: ");
	if (cli::switchMAPI.isSet())
	{
		printf("MAPI -> MIME\n");

		printf("   Save Format: %s\n", cli::switchRFC822.isSet() ? "RFC822" : "RFC1521");

		if (cli::switchWrap.isSet())
		{
			printf("   Line Wrap: ");
			if (ulWrapLines == 0)
				printf("OFF\n");
			else
				printf("%lu\n", ulWrapLines);
		}
	}
	else if (cli::switchMAPI.isSet())
	{
		printf("MIME -> MAPI\n");
		if (cli::switchUnicode.isSet())
		{
			printf("   Building Unicode MSG file\n");
		}

		if (cli::switchCharset.isSet())
		{
			printf("   CodePage: %lu\n", ulCodePage);
			printf("   CharSetType: ");
			switch (cSetType)
			{
			case CHARSET_BODY:
				printf("CHARSET_BODY");
				break;
			case CHARSET_HEADER:
				printf("CHARSET_HEADER");
				break;
			case CHARSET_WEB:
				printf("CHARSET_WEB");
				break;
			}

			printf("\n");
			printf("   CharSetApplyType: ");
			switch (cSetApplyType)
			{
			case CSET_APPLY_UNTAGGED:
				printf("CSET_APPLY_UNTAGGED");
				break;
			case CSET_APPLY_ALL:
				printf("CSET_APPLY_ALL");
				break;
			case CSET_APPLY_TAG_ALL:
				printf("CSET_APPLY_TAG_ALL");
				break;
			}

			printf("\n");
		}
	}

	if (cli::switchCCSFFlags.hasArgs())
	{
		auto szFlags = flags::InterpretFlags(flagCcsf, convertFlags);
		if (!szFlags.empty())
		{
			printf("   Conversion Flags: %ws\n", szFlags.c_str());
		}
	}

	if (cli::switchEncoding.isSet())
	{
		auto szType = flags::InterpretFlags(flagIet, ulEncodingType);
		if (!szType.empty())
		{
			printf("   Encoding Type: %ws\n", szType.c_str());
		}
	}

	if (cli::switchAddressBook.isSet())
	{
		printf("   Using Address Book\n");
	}

	auto hRes = S_OK;

	LPADRBOOK lpAdrBook = nullptr;
	if (cli::switchAddressBook.isSet() && ProgOpts.lpMAPISession)
	{
		WC_MAPI(ProgOpts.lpMAPISession->OpenAddressBook(NULL, nullptr, AB_NO_DIALOG, &lpAdrBook));
		if (FAILED(hRes)) printf("OpenAddressBook returned an error: 0x%08lx\n", hRes);
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
				printf("MimeOleGetCodePageCharset returned 0x%08lX\n", hRes);
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
		printf("File converted successfully.\n");
	}
	else if (REGDB_E_CLASSNOTREG == hRes)
	{
		printf("MAPI <-> MIME converter not found. Perhaps Outlook is not installed.\n");
	}
	else
	{
		printf("Conversion returned an error: 0x%08lx\n", hRes);
	}

	if (lpAdrBook) lpAdrBook->Release();
}