#include <StdAfx.h>
#include <MrMapi/MMPst.h>
#include <MrMapi/mmcli.h>
#include <core/utility/registry.h>
#include <core/utility/output.h>

#define NDB_CRYPT_NONE 0
#define NDB_CRYPT_PERMUTE 1
#define NDB_CRYPT_CYCLIC 2

struct PSTHEADER
{
	DWORD dwMagic;
	DWORD dwCRCPartial;
	WORD wMagicClient;
	WORD wVer;
	WORD wVerClient;
	BYTE bPlatformCreate;
	BYTE bPlatformAccess;
	DWORD dwReserved1;
	DWORD dwReserved2;
};

#define NDBANSISMALL 14
#define NDBANSILARGE 15
#define NDBUNICODE 23
#define NDBUNICODE2 36

struct BREFANSI
{
	DWORD bid;
	DWORD ib;
};

struct ROOTANSI
{
	ULONG dwReserved;
	DWORD ibFileEof;
	DWORD ibAMapLast;
	DWORD cbAMapFree;
	DWORD cbPMapFree;
	BREFANSI brefNBT;
	BREFANSI brefBBT;
	BYTE fAMapValid;
	BYTE bARVec;
	WORD cARVec;
};

struct HEADER2ANSI
{
	DWORD bidNextB;
	DWORD bidNextP;
	DWORD dwUnique;
	BYTE rgnid[128];
	ROOTANSI root;
	BYTE rgbFM[128];
	BYTE rgbFP[128];
	BYTE bSentinel;
	BYTE bCryptMethod;
};

struct BREFUNICODE
{
	ULONGLONG bid;
	ULONGLONG ib;
};

struct ROOTUNICODE
{
	ULONG dwReserved;
	ULONGLONG ibFileEof;
	ULONGLONG ibAMapLast;
	ULONGLONG cbAMapFree;
	ULONGLONG cbPMapFree;
	BREFUNICODE brefNBT;
	BREFUNICODE brefBBT;
	BYTE fAMapValid;
	BYTE bARVec;
	WORD cARVec;
};

struct HEADER2UNICODE
{
	ULONGLONG bidUnused;
	ULONGLONG bidNextP;
	DWORD dwUnique;
	BYTE rgnid[128];
	ROOTUNICODE root;
	BYTE rgbFM[128];
	BYTE rgbFP[128];
	BYTE bSentinel;
	BYTE bCryptMethod;
};

void PrintCryptType(BYTE bCryptMethod) noexcept
{
	wprintf(L"0x%02X (", bCryptMethod);
	switch (bCryptMethod)
	{
	case NDB_CRYPT_NONE:
		wprintf(L"not encoded");
		break;
	case NDB_CRYPT_PERMUTE:
		wprintf(L"permutative encoding");
		break;
	case NDB_CRYPT_CYCLIC:
		wprintf(L"cyclic encoding");
		break;
	}

	wprintf(L")");
}

void PrintAMAPValid(BYTE fAMapValid) noexcept
{
	wprintf(L"0x%02X (", fAMapValid);
	switch (fAMapValid)
	{
	case 0:
		wprintf(L"not valid");
		break;
	case 1:
	case 2:
		wprintf(L"valid");
		break;
	}

	wprintf(L")");
}

#define KB (1024)
#define MB (KB * 1024)
#define GB (MB * 1024)

void PrintFileSize(ULONGLONG ullFileSize) noexcept
{
	double scaledSize = 0;
	if (ullFileSize > GB)
	{
		scaledSize = ullFileSize / static_cast<double>(GB);
		wprintf(L"%.2f GB", scaledSize);
	}
	else if (ullFileSize > MB)
	{
		scaledSize = ullFileSize / static_cast<double>(MB);
		wprintf(L"%.2f MB", scaledSize);
	}
	else if (ullFileSize > KB)
	{
		scaledSize = ullFileSize / static_cast<double>(KB);
		wprintf(L"%.2f KB", scaledSize);
	}

	wprintf(L" (%I64u bytes)", ullFileSize);
}

void DoPST()
{
	const auto input = cli::switchInput[0];
	wprintf(L"Analyzing %ws\n", input.c_str());

	struct _stat64 stats = {};
	if (FAILED(EC_W32(_wstati64(input.c_str(), &stats))))
	{
		wprintf(L"Input file does not exist.\n");
		if (cli::switchVerbose.isSet())
		{
			wprintf(L"errno = %d\n", errno);
		}

		return;
	}

	const auto fIn = output::MyOpenFileMode(input, L"rb");
	if (!fIn)
	{
		wprintf(L"Cannot open input file. Try closing Outlook and trying again.\n");
		wprintf(L"File Size (actual) = ");
		PrintFileSize(stats.st_size);
		wprintf(L"\n");
		return;
	}

	PSTHEADER pstHeader = {};
	if (fread(&pstHeader, sizeof(PSTHEADER), 1, fIn))
	{
		ULONGLONG ibFileEof = 0;
		ULONGLONG cbAMapFree = 0;
		BYTE fAMapValid = 0;
		BYTE bCryptMethod = 0;

		if (NDBANSISMALL == pstHeader.wVer || NDBANSILARGE == pstHeader.wVer)
		{
			wprintf(L"ANSI PST (%ws)\n", NDBANSISMALL == pstHeader.wVer ? L"small" : L"large");
			HEADER2ANSI h2Ansi = {0};
			if (fread(&h2Ansi, sizeof(HEADER2ANSI), 1, fIn))
			{
				ibFileEof = h2Ansi.root.ibFileEof;
				cbAMapFree = h2Ansi.root.cbAMapFree;
				fAMapValid = h2Ansi.root.fAMapValid;
				bCryptMethod = h2Ansi.bCryptMethod;
			}
		}
		else if (NDBUNICODE == pstHeader.wVer || NDBUNICODE2 == pstHeader.wVer)
		{
			wprintf(L"Unicode PST\n");
			HEADER2UNICODE h2Unicode = {};
			if (fread(&h2Unicode, sizeof(HEADER2UNICODE), 1, fIn))
			{
				ibFileEof = h2Unicode.root.ibFileEof;
				cbAMapFree = h2Unicode.root.cbAMapFree;
				fAMapValid = h2Unicode.root.fAMapValid;
				bCryptMethod = h2Unicode.bCryptMethod;
			}
		}

		if (ibFileEof != static_cast<ULONGLONG>(stats.st_size))
		{
			wprintf(L"File Size (header) = ");
			PrintFileSize(ibFileEof);
			wprintf(L"\n");
			wprintf(L"File Size (actual) = ");
			PrintFileSize(stats.st_size);
			wprintf(L"\n");
		}
		else
		{
			wprintf(L"File Size = ");
			PrintFileSize(ibFileEof);
			wprintf(L"\n");
		}

		wprintf(L"Free Space = ");
		PrintFileSize(cbAMapFree);
		wprintf(L"\n");
		if (ibFileEof != 0)
		{
			const auto percentFree = cbAMapFree * 100.0 / ibFileEof;
			wprintf(L"Percent free = %.2f%%\n", percentFree);
		}

		if (cli::switchVerbose.isSet())
		{
			wprintf(L"\n");
			wprintf(L"fAMapValid = ");
			PrintAMAPValid(fAMapValid);
			wprintf(L"\n");
			wprintf(L"Crypt Method = ");
			PrintCryptType(bCryptMethod);
			wprintf(L"\n");

			wprintf(L"Magic Number = 0x%08X\n", pstHeader.dwMagic);
			wprintf(L"wMagicClient = 0x%04X\n", pstHeader.wMagicClient);
			wprintf(L"wVer = 0x%04X\n", pstHeader.wVer);
			wprintf(L"wVerClient = 0x%04X\n", pstHeader.wVerClient);
		}
	}
	else
	{
		wprintf(L"Could not read from file. File may be locked or empty.\n");
	}

	fclose(fIn);
}