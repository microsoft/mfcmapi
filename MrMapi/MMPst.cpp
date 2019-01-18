#include <StdAfx.h>
#include <MrMapi/cli.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>
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

void PrintCryptType(BYTE bCryptMethod)
{
	printf("0x%02X (", bCryptMethod);
	switch (bCryptMethod)
	{
	case NDB_CRYPT_NONE:
		printf("not encoded");
		break;
	case NDB_CRYPT_PERMUTE:
		printf("permutative encoding");
		break;
	case NDB_CRYPT_CYCLIC:
		printf("cyclic encoding");
		break;
	}

	printf(")");
}

void PrintAMAPValid(BYTE fAMapValid)
{
	printf("0x%02X (", fAMapValid);
	switch (fAMapValid)
	{
	case 0:
		printf("not valid");
		break;
	case 1:
	case 2:
		printf("valid");
		break;
	}

	printf(")");
}

#define KB (1024)
#define MB (KB * 1024)
#define GB (MB * 1024)

void PrintFileSize(ULONGLONG ullFileSize)
{
	double scaledSize = 0;
	if (ullFileSize > GB)
	{
		scaledSize = ullFileSize / static_cast<double>(GB);
		printf("%.2f GB", scaledSize);
	}
	else if (ullFileSize > MB)
	{
		scaledSize = ullFileSize / static_cast<double>(MB);
		printf("%.2f MB", scaledSize);
	}
	else if (ullFileSize > KB)
	{
		scaledSize = ullFileSize / static_cast<double>(KB);
		printf("%.2f KB", scaledSize);
	}

	printf(" (%I64u bytes)", ullFileSize);
}

void DoPST(_In_ cli::MYOPTIONS ProgOpts)
{
	printf("Analyzing %ws\n", ProgOpts.lpszInput.c_str());

	struct _stat64 stats = {0};
	_wstati64(ProgOpts.lpszInput.c_str(), &stats);

	const auto fIn = output::MyOpenFileMode(ProgOpts.lpszInput, L"rb");
	if (fIn)
	{
		PSTHEADER pstHeader = {0};
		if (fread(&pstHeader, sizeof(PSTHEADER), 1, fIn))
		{
			ULONGLONG ibFileEof = 0;
			ULONGLONG cbAMapFree = 0;
			BYTE fAMapValid = 0;
			BYTE bCryptMethod = 0;

			if (NDBANSISMALL == pstHeader.wVer || NDBANSILARGE == pstHeader.wVer)
			{
				printf("ANSI PST (%ws)\n", NDBANSISMALL == pstHeader.wVer ? L"small" : L"large");
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
				printf("Unicode PST\n");
				HEADER2UNICODE h2Unicode = {0};
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
				printf("File Size (header) = ");
				PrintFileSize(ibFileEof);
				printf("\n");
				printf("File Size (actual) = ");
				PrintFileSize(stats.st_size);
				printf("\n");
			}
			else
			{
				printf("File Size = ");
				PrintFileSize(ibFileEof);
				printf("\n");
			}

			printf("Free Space = ");
			PrintFileSize(cbAMapFree);
			printf("\n");
			const auto percentFree = cbAMapFree * 100.0 / ibFileEof;
			printf("Percent free = %.2f%%\n", percentFree);
			if (fIsSet(DBGGeneric))
			{
				printf("\n");
				printf("fAMapValid = ");
				PrintAMAPValid(fAMapValid);
				printf("\n");
				printf("Crypt Method = ");
				PrintCryptType(bCryptMethod);
				printf("\n");

				printf("Magic Number = 0x%08X\n", pstHeader.dwMagic);
				printf("wMagicClient = 0x%04X\n", pstHeader.wMagicClient);
				printf("wVer = 0x%04X\n", pstHeader.wVer);
				printf("wVerClient = 0x%04X\n", pstHeader.wVerClient);
			}
		}
		else
		{
			printf("Could not read from %ws. File may be locked or empty.\n", ProgOpts.lpszInput.c_str());
		}

		fclose(fIn);
	}
	else
	{
		printf("Cannot open input file %ws\n", ProgOpts.lpszInput.c_str());
	}
}