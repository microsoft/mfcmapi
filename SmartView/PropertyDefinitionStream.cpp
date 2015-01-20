#include "stdafx.h"
#include "..\stdafx.h"
#include "PropertyDefinitionStream.h"
#include "..\String.h"
#include "..\InterpretProp2.h"
#include "..\ParseProperty.h"
#include "..\ExtraPropTags.h"

PropertyDefinitionStream::PropertyDefinitionStream(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	m_wVersion = 0;
	m_dwFieldDefinitionCount = 0;
	m_pfdFieldDefinitions = 0;
}

PropertyDefinitionStream::~PropertyDefinitionStream()
{
	if (m_pfdFieldDefinitions)
	{
		DWORD iDef = 0;
		for (iDef = 0; iDef < m_dwFieldDefinitionCount; iDef++)
		{
			delete[] m_pfdFieldDefinitions[iDef].szNmidName;
			delete[] m_pfdFieldDefinitions[iDef].pasNameANSI.szCharacters;
			delete[] m_pfdFieldDefinitions[iDef].pasFormulaANSI.szCharacters;
			delete[] m_pfdFieldDefinitions[iDef].pasValidationRuleANSI.szCharacters;
			delete[] m_pfdFieldDefinitions[iDef].pasValidationTextANSI.szCharacters;
			delete[] m_pfdFieldDefinitions[iDef].pasErrorANSI.szCharacters;

			if (m_pfdFieldDefinitions[iDef].psbSkipBlocks)
			{
				DWORD iSkip = 0;
				for (iSkip = 0; iSkip < m_pfdFieldDefinitions[iDef].dwSkipBlockCount; iSkip++)
				{
					delete[] m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContent;
				}

				delete[] m_pfdFieldDefinitions[iDef].psbSkipBlocks;
			}
		}
		delete[] m_pfdFieldDefinitions;
	}
}

void ReadPackedAnsiString(_In_ CBinaryParser* pParser, _In_ PackedAnsiString* ppasString)
{
	if (!pParser || !ppasString) return;

	pParser->GetBYTE(&ppasString->cchLength);
	if (0xFF == ppasString->cchLength)
	{
		pParser->GetWORD(&ppasString->cchExtendedLength);
	}

	pParser->GetStringA(ppasString->cchExtendedLength ? ppasString->cchExtendedLength : ppasString->cchLength,
		&ppasString->szCharacters);
}

void ReadPackedUnicodeString(_In_ CBinaryParser* pParser, _In_ PackedUnicodeString* ppusString)
{
	if (!pParser || !ppusString) return;

	pParser->GetBYTE(&ppusString->cchLength);
	if (0xFF == ppusString->cchLength)
	{
		pParser->GetWORD(&ppusString->cchExtendedLength);
	}

	pParser->GetStringW(ppusString->cchExtendedLength ? ppusString->cchExtendedLength : ppusString->cchLength,
		&ppusString->szCharacters);
}

void PropertyDefinitionStream::Parse()
{
	m_Parser.GetWORD(&m_wVersion);
	m_Parser.GetDWORD(&m_dwFieldDefinitionCount);
	if (m_dwFieldDefinitionCount && m_dwFieldDefinitionCount < _MaxEntriesLarge)
	{
		m_pfdFieldDefinitions = new FieldDefinition[m_dwFieldDefinitionCount];

		if (m_pfdFieldDefinitions)
		{
			memset(m_pfdFieldDefinitions, 0, m_dwFieldDefinitionCount * sizeof(FieldDefinition));

			DWORD iDef = 0;
			for (iDef = 0; iDef < m_dwFieldDefinitionCount; iDef++)
			{
				m_Parser.GetDWORD(&m_pfdFieldDefinitions[iDef].dwFlags);
				m_Parser.GetWORD(&m_pfdFieldDefinitions[iDef].wVT);
				m_Parser.GetDWORD(&m_pfdFieldDefinitions[iDef].dwDispid);
				m_Parser.GetWORD(&m_pfdFieldDefinitions[iDef].wNmidNameLength);
				m_Parser.GetStringW(m_pfdFieldDefinitions[iDef].wNmidNameLength,
					&m_pfdFieldDefinitions[iDef].szNmidName);

				ReadPackedAnsiString(&m_Parser, &m_pfdFieldDefinitions[iDef].pasNameANSI);
				ReadPackedAnsiString(&m_Parser, &m_pfdFieldDefinitions[iDef].pasFormulaANSI);
				ReadPackedAnsiString(&m_Parser, &m_pfdFieldDefinitions[iDef].pasValidationRuleANSI);
				ReadPackedAnsiString(&m_Parser, &m_pfdFieldDefinitions[iDef].pasValidationTextANSI);
				ReadPackedAnsiString(&m_Parser, &m_pfdFieldDefinitions[iDef].pasErrorANSI);

				if (PropDefV2 == m_wVersion)
				{
					m_Parser.GetDWORD(&m_pfdFieldDefinitions[iDef].dwInternalType);

					// Have to count how many skip blocks are here.
					// The only way to do that is to parse them. So we parse once without storing, allocate, then reparse.
					size_t stBookmark = m_Parser.GetCurrentOffset();

					DWORD dwSkipBlockCount = 0;

					for (;;)
					{
						dwSkipBlockCount++;
						DWORD dwBlock = 0;
						m_Parser.GetDWORD(&dwBlock);
						if (!dwBlock) break; // we hit the last, zero byte block, or the end of the buffer
						m_Parser.Advance(dwBlock);
					}

					m_Parser.SetCurrentOffset(stBookmark); // We're done with our first pass, restore the bookmark

					m_pfdFieldDefinitions[iDef].dwSkipBlockCount = dwSkipBlockCount;
					if (m_pfdFieldDefinitions[iDef].dwSkipBlockCount &&
						m_pfdFieldDefinitions[iDef].dwSkipBlockCount < _MaxEntriesSmall)
					{
						m_pfdFieldDefinitions[iDef].psbSkipBlocks = new SkipBlock[m_pfdFieldDefinitions[iDef].dwSkipBlockCount];

						if (m_pfdFieldDefinitions[iDef].psbSkipBlocks)
						{
							memset(m_pfdFieldDefinitions[iDef].psbSkipBlocks, 0, m_pfdFieldDefinitions[iDef].dwSkipBlockCount * sizeof(SkipBlock));

							DWORD iSkip = 0;
							for (iSkip = 0; iSkip < m_pfdFieldDefinitions[iDef].dwSkipBlockCount; iSkip++)
							{
								m_Parser.GetDWORD(&m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].dwSize);
								m_Parser.GetBYTES(m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].dwSize,
									_MaxBytes,
									&m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContent);
							}
						}
					}
				}
			}
		}
	}
}

_Check_return_ wstring PackedAnsiStringToString(DWORD dwFlags, _In_ PackedAnsiString* ppasString)
{
	if (!ppasString) return L"";

	wstring szFieldName;
	wstring szPackedAnsiString;

	szFieldName = formatmessage(dwFlags);

	szPackedAnsiString = formatmessage(IDS_PROPDEFPACKEDSTRINGLEN,
		szFieldName.c_str(),
		(0xFF == ppasString->cchLength) ? ppasString->cchExtendedLength : ppasString->cchLength);
	if (ppasString->szCharacters)
	{
		szPackedAnsiString += formatmessage(IDS_PROPDEFPACKEDSTRINGCHARS, szFieldName);
		szPackedAnsiString += LPSTRToWstring(ppasString->szCharacters);
	}

	return szPackedAnsiString;
}

_Check_return_ wstring PackedUnicodeStringToString(DWORD dwFlags, _In_ PackedUnicodeString* ppusString)
{
	if (!ppusString) return L"";

	wstring szFieldName;
	wstring szPackedUnicodeString;

	szFieldName = formatmessage(dwFlags);

	szPackedUnicodeString = formatmessage(IDS_PROPDEFPACKEDSTRINGLEN,
		szFieldName.c_str(),
		(0xFF == ppusString->cchLength) ? ppusString->cchExtendedLength : ppusString->cchLength);
	if (ppusString->szCharacters)
	{
		szPackedUnicodeString += formatmessage(IDS_PROPDEFPACKEDSTRINGCHARS, szFieldName);
		szPackedUnicodeString += ppusString->szCharacters;
	}

	return szPackedUnicodeString;
}

_Check_return_ wstring PropertyDefinitionStream::ToStringInternal()
{
	wstring szPropertyDefinitionStream;
	wstring szVersion = InterpretFlags(flagPropDefVersion, m_wVersion);

	szPropertyDefinitionStream = formatmessage(IDS_PROPDEFHEADER,
		m_wVersion, szVersion.c_str(),
		m_dwFieldDefinitionCount);

	if (m_pfdFieldDefinitions)
	{
		DWORD iDef = 0;
		for (iDef = 0; iDef < m_dwFieldDefinitionCount; iDef++)
		{
			wstring szFlags = InterpretFlags(flagPDOFlag, m_pfdFieldDefinitions[iDef].dwFlags);
			wstring szVarEnum = InterpretFlags(flagVarEnum, m_pfdFieldDefinitions[iDef].wVT);
			szPropertyDefinitionStream += formatmessage(IDS_PROPDEFFIELDHEADER,
				iDef,
				m_pfdFieldDefinitions[iDef].dwFlags, szFlags.c_str(),
				m_pfdFieldDefinitions[iDef].wVT, szVarEnum.c_str(),
				m_pfdFieldDefinitions[iDef].dwDispid);

			if (m_pfdFieldDefinitions[iDef].dwDispid)
			{
				if (m_pfdFieldDefinitions[iDef].dwDispid < 0x8000)
				{
					LPTSTR szDispidName = NULL;
					PropTagToPropName(m_pfdFieldDefinitions[iDef].dwDispid, false, NULL, &szDispidName);
					if (szDispidName)
					{
						szPropertyDefinitionStream += formatmessage(IDS_PROPDEFDISPIDTAG, LPTSTRToWstring(szDispidName).c_str());
					}

					delete[] szDispidName;
				}
				else
				{
					wstring szDispidName;
					MAPINAMEID mnid = { 0 };
					mnid.lpguid = NULL;
					mnid.ulKind = MNID_ID;
					mnid.Kind.lID = m_pfdFieldDefinitions[iDef].dwDispid;
					szDispidName = NameIDToPropName(&mnid);
					if (!szDispidName.empty())
					{
						szPropertyDefinitionStream += formatmessage(IDS_PROPDEFDISPIDNAMED, szDispidName.c_str());
					}
				}
			}

			szPropertyDefinitionStream += formatmessage(IDS_PROPDEFNMIDNAME,
				m_pfdFieldDefinitions[iDef].wNmidNameLength);
			if (m_pfdFieldDefinitions[iDef].szNmidName)
			{
				szPropertyDefinitionStream += m_pfdFieldDefinitions[iDef].szNmidName;
			}

			wstring szTab1 = formatmessage(IDS_PROPDEFTAB1);

			szPropertyDefinitionStream += szTab1 + PackedAnsiStringToString(IDS_PROPDEFNAME, &m_pfdFieldDefinitions[iDef].pasNameANSI);
			szPropertyDefinitionStream += szTab1 + PackedAnsiStringToString(IDS_PROPDEFFORUMULA, &m_pfdFieldDefinitions[iDef].pasFormulaANSI);
			szPropertyDefinitionStream += szTab1 + PackedAnsiStringToString(IDS_PROPDEFVRULE, &m_pfdFieldDefinitions[iDef].pasValidationRuleANSI);
			szPropertyDefinitionStream += szTab1 + PackedAnsiStringToString(IDS_PROPDEFVTEXT, &m_pfdFieldDefinitions[iDef].pasValidationTextANSI);
			szPropertyDefinitionStream += szTab1 + PackedAnsiStringToString(IDS_PROPDEFERROR, &m_pfdFieldDefinitions[iDef].pasErrorANSI);

			if (PropDefV2 == m_wVersion)
			{
				szFlags = InterpretFlags(flagInternalType, m_pfdFieldDefinitions[iDef].dwInternalType);
				szPropertyDefinitionStream += formatmessage(IDS_PROPDEFINTERNALTYPE,
					m_pfdFieldDefinitions[iDef].dwInternalType, szFlags.c_str(),
					m_pfdFieldDefinitions[iDef].dwSkipBlockCount);

				if (m_pfdFieldDefinitions[iDef].psbSkipBlocks)
				{
					DWORD iSkip = 0;
					for (iSkip = 0; iSkip < m_pfdFieldDefinitions[iDef].dwSkipBlockCount; iSkip++)
					{
						szPropertyDefinitionStream += formatmessage(IDS_PROPDEFSKIPBLOCK,
							iSkip,
							m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].dwSize);

						if (m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].dwSize &&
							m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContent)
						{
							if (0 == iSkip)
							{
								// Parse this on the fly
								CBinaryParser ParserFirstBlock(
									m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].dwSize,
									m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContent);
								PackedUnicodeString pusString = { 0 };
								ReadPackedUnicodeString(&ParserFirstBlock, &pusString);
								szPropertyDefinitionStream += formatmessage(IDS_PROPDEFTAB2) + PackedUnicodeStringToString(IDS_PROPDEFFIELDNAME, &pusString);

								delete[] pusString.szCharacters;
							}
							else
							{
								SBinary sBin = { 0 };
								sBin.cb = (ULONG)m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].dwSize;
								sBin.lpb = m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContent;

								szPropertyDefinitionStream += formatmessage(IDS_PROPDEFCONTENT) + BinToHexString(&sBin, true);
							}
						}
					}
				}
			}
		}
	}

	return szPropertyDefinitionStream;
}