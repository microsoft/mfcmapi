#include <StdAfx.h>
#include <MrMapi/MrMAPI.h>
#include <MrMapi/MMMapiMime.h>
#include <MAPI/MAPIMime.h>
#include <ImportProcs.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

#define CHECKFLAG(__flag) ((ProgOpts.ulMAPIMIMEFlags & (__flag)) == (__flag))
void DoMAPIMIME(_In_ MYOPTIONS ProgOpts)
{
	printf("Message File Converter\n");
	printf("Options specified:\n");
	printf("   Input File: %ws\n", ProgOpts.lpszInput.c_str());
	printf("   Output File: %ws\n", ProgOpts.lpszOutput.c_str());
	printf("   Conversion Type: ");
	if (CHECKFLAG(MAPIMIME_TOMIME))
	{
		printf("MAPI -> MIME\n");

		printf("   Save Format: %s\n", CHECKFLAG(MAPIMIME_RFC822) ? "RFC822" : "RFC1521");

		if (CHECKFLAG(MAPIMIME_WRAP))
		{
			printf("   Line Wrap: ");
			if (0 == ProgOpts.ulWrapLines)
				printf("OFF\n");
			else
				printf("%u\n", ProgOpts.ulWrapLines);
		}
	}
	else if (CHECKFLAG(MAPIMIME_TOMAPI))
	{
		printf("MIME -> MAPI\n");
		if (CHECKFLAG(MAPIMIME_UNICODE))
		{
			printf("   Building Unicode MSG file\n");
		}
		if (CHECKFLAG(MAPIMIME_CHARSET))
		{
			printf("   CodePage: %u\n", ProgOpts.ulCodePage);
			printf("   CharSetType: ");
			switch (ProgOpts.cSetType)
			{
			case CHARSET_BODY: printf("CHARSET_BODY"); break;
			case CHARSET_HEADER: printf("CHARSET_HEADER"); break;
			case CHARSET_WEB: printf("CHARSET_WEB"); break;
			}
			printf("\n");
			printf("   CharSetApplyType: ");
			switch (ProgOpts.cSetApplyType)
			{
			case CSET_APPLY_UNTAGGED: printf("CSET_APPLY_UNTAGGED"); break;
			case CSET_APPLY_ALL: printf("CSET_APPLY_ALL"); break;
			case CSET_APPLY_TAG_ALL: printf("CSET_APPLY_TAG_ALL"); break;
			}
			printf("\n");
		}
	}

	if (0 != ProgOpts.ulConvertFlags)
	{
		auto szFlags = interpretprop::InterpretFlags(flagCcsf, ProgOpts.ulConvertFlags);
		if (!szFlags.empty())
		{
			printf("   Conversion Flags: %ws\n", szFlags.c_str());
		}
	}

	if (CHECKFLAG(MAPIMIME_ENCODING))
	{
		auto szType = interpretprop::InterpretFlags(flagIet, ProgOpts.ulEncodingType);
		if (!szType.empty())
		{
			printf("   Encoding Type: %ws\n", szType.c_str());
		}
	}

	if (CHECKFLAG(MAPIMIME_ADDRESSBOOK))
	{
		printf("   Using Address Book\n");
	}

	auto hRes = S_OK;

	LPADRBOOK lpAdrBook = nullptr;
	if (CHECKFLAG(MAPIMIME_ADDRESSBOOK) && ProgOpts.lpMAPISession)
	{
		WC_MAPI(ProgOpts.lpMAPISession->OpenAddressBook(NULL, NULL, AB_NO_DIALOG, &lpAdrBook));
		if (FAILED(hRes)) printf("OpenAddressBook returned an error: 0x%08x\n", hRes);
	}

	if (CHECKFLAG(MAPIMIME_TOMIME))
	{
		// Source file is MSG, target is EML
		WC_H(ConvertMSGToEML(
			ProgOpts.lpszInput.c_str(),
			ProgOpts.lpszOutput.c_str(),
			ProgOpts.ulConvertFlags,
			CHECKFLAG(MAPIMIME_ENCODING) ? static_cast<ENCODINGTYPE>(ProgOpts.ulEncodingType) : IET_UNKNOWN,
			CHECKFLAG(MAPIMIME_RFC822) ? SAVE_RFC822 : SAVE_RFC1521,
			CHECKFLAG(MAPIMIME_WRAP) ? ProgOpts.ulWrapLines : USE_DEFAULT_WRAPPING,
			lpAdrBook));
	}
	else if (CHECKFLAG(MAPIMIME_TOMAPI))
	{
		// Source file is EML, target is MSG
		HCHARSET hCharSet = nullptr;
		if (CHECKFLAG(MAPIMIME_CHARSET))
		{
			WC_H(MyMimeOleGetCodePageCharset(ProgOpts.ulCodePage, ProgOpts.cSetType, &hCharSet));
			if (FAILED(hRes))
			{
				printf("MimeOleGetCodePageCharset returned 0x%08X\n", hRes);
			}
		}
		if (SUCCEEDED(hRes))
		{
			WC_H(ConvertEMLToMSG(
				ProgOpts.lpszInput.c_str(),
				ProgOpts.lpszOutput.c_str(),
				ProgOpts.ulConvertFlags,
				CHECKFLAG(MAPIMIME_CHARSET),
				hCharSet,
				ProgOpts.cSetApplyType,
				lpAdrBook,
				CHECKFLAG(MAPIMIME_UNICODE)));
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
		printf("Conversion returned an error: 0x%08x\n", hRes);
	}
	if (lpAdrBook) lpAdrBook->Release();
}