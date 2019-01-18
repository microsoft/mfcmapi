#include <StdAfx.h>
#include <MrMapi/MMMapiMime.h>
#include <MAPI/MapiMime.h>
#include <core/utility/import.h>
#include <Interpret/InterpretProp.h>
#include <core/mapi/extraPropTags.h>
#include <MrMapi/cli.h>

#define CHECKFLAG(__flag) ((ProgOpts.MAPIMIMEFlags & (__flag)) == (__flag))
void DoMAPIMIME(_In_ cli::MYOPTIONS ProgOpts)
{
	printf("Message File Converter\n");
	printf("Options specified:\n");
	printf("   Input File: %ws\n", ProgOpts.lpszInput.c_str());
	printf("   Output File: %ws\n", ProgOpts.lpszOutput.c_str());
	printf("   Conversion Type: ");
	if (CHECKFLAG(cli::MAPIMIME_TOMIME))
	{
		printf("MAPI -> MIME\n");

		printf("   Save Format: %s\n", CHECKFLAG(cli::MAPIMIME_RFC822) ? "RFC822" : "RFC1521");

		if (CHECKFLAG(cli::MAPIMIME_WRAP))
		{
			printf("   Line Wrap: ");
			if (0 == ProgOpts.ulWrapLines)
				printf("OFF\n");
			else
				printf("%lu\n", ProgOpts.ulWrapLines);
		}
	}
	else if (CHECKFLAG(cli::MAPIMIME_TOMAPI))
	{
		printf("MIME -> MAPI\n");
		if (CHECKFLAG(cli::MAPIMIME_UNICODE))
		{
			printf("   Building Unicode MSG file\n");
		}
		if (CHECKFLAG(cli::MAPIMIME_CHARSET))
		{
			printf("   CodePage: %lu\n", ProgOpts.ulCodePage);
			printf("   CharSetType: ");
			switch (ProgOpts.cSetType)
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
			switch (ProgOpts.cSetApplyType)
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

	if (0 != ProgOpts.convertFlags)
	{
		auto szFlags = interpretprop::InterpretFlags(flagCcsf, ProgOpts.convertFlags);
		if (!szFlags.empty())
		{
			printf("   Conversion Flags: %ws\n", szFlags.c_str());
		}
	}

	if (CHECKFLAG(cli::MAPIMIME_ENCODING))
	{
		auto szType = interpretprop::InterpretFlags(flagIet, ProgOpts.ulEncodingType);
		if (!szType.empty())
		{
			printf("   Encoding Type: %ws\n", szType.c_str());
		}
	}

	if (CHECKFLAG(cli::MAPIMIME_ADDRESSBOOK))
	{
		printf("   Using Address Book\n");
	}

	auto hRes = S_OK;

	LPADRBOOK lpAdrBook = nullptr;
	if (CHECKFLAG(cli::MAPIMIME_ADDRESSBOOK) && ProgOpts.lpMAPISession)
	{
		WC_MAPI(ProgOpts.lpMAPISession->OpenAddressBook(NULL, NULL, AB_NO_DIALOG, &lpAdrBook));
		if (FAILED(hRes)) printf("OpenAddressBook returned an error: 0x%08lx\n", hRes);
	}

	if (CHECKFLAG(cli::MAPIMIME_TOMIME))
	{
		// Source file is MSG, target is EML
		hRes = WC_H(mapi::mapimime::ConvertMSGToEML(
			ProgOpts.lpszInput.c_str(),
			ProgOpts.lpszOutput.c_str(),
			ProgOpts.convertFlags,
			CHECKFLAG(cli::MAPIMIME_ENCODING) ? static_cast<ENCODINGTYPE>(ProgOpts.ulEncodingType) : IET_UNKNOWN,
			CHECKFLAG(cli::MAPIMIME_RFC822) ? SAVE_RFC822 : SAVE_RFC1521,
			CHECKFLAG(cli::MAPIMIME_WRAP) ? ProgOpts.ulWrapLines : USE_DEFAULT_WRAPPING,
			lpAdrBook));
	}
	else if (CHECKFLAG(cli::MAPIMIME_TOMAPI))
	{
		// Source file is EML, target is MSG
		HCHARSET hCharSet = nullptr;
		if (CHECKFLAG(cli::MAPIMIME_CHARSET))
		{
			hRes = WC_H(import::MyMimeOleGetCodePageCharset(ProgOpts.ulCodePage, ProgOpts.cSetType, &hCharSet));
			if (FAILED(hRes))
			{
				printf("MimeOleGetCodePageCharset returned 0x%08lX\n", hRes);
			}
		}

		if (SUCCEEDED(hRes))
		{
			hRes = WC_H(mapi::mapimime::ConvertEMLToMSG(
				ProgOpts.lpszInput.c_str(),
				ProgOpts.lpszOutput.c_str(),
				ProgOpts.convertFlags,
				CHECKFLAG(cli::MAPIMIME_CHARSET),
				hCharSet,
				ProgOpts.cSetApplyType,
				lpAdrBook,
				CHECKFLAG(cli::MAPIMIME_UNICODE)));
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