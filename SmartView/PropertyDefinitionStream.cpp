#include "stdafx.h"
#include "PropertyDefinitionStream.h"
#include "InterpretProp2.h"
#include "ExtraPropTags.h"

PropertyDefinitionStream::PropertyDefinitionStream()
{
	m_wVersion = 0;
	m_dwFieldDefinitionCount = 0;
}

PackedAnsiString ReadPackedAnsiString(_In_ CBinaryParser* pParser)
{
	PackedAnsiString packedAnsiString;
	packedAnsiString.cchLength = 0;
	packedAnsiString.cchExtendedLength = 0;
	if (pParser)
	{
		pParser->GetBYTE(&packedAnsiString.cchLength);
		if (0xFF == packedAnsiString.cchLength)
		{
			pParser->GetWORD(&packedAnsiString.cchExtendedLength);
		}

		packedAnsiString.szCharacters = pParser->GetStringA(packedAnsiString.cchExtendedLength ? packedAnsiString.cchExtendedLength : packedAnsiString.cchLength);
	}

	return packedAnsiString;
}

PackedUnicodeString ReadPackedUnicodeString(_In_ CBinaryParser* pParser)
{
	PackedUnicodeString packedUnicodeString;
	packedUnicodeString.cchLength = 0;
	packedUnicodeString.cchExtendedLength = 0;
	if (pParser)
	{
		pParser->GetBYTE(&packedUnicodeString.cchLength);
		if (0xFF == packedUnicodeString.cchLength)
		{
			pParser->GetWORD(&packedUnicodeString.cchExtendedLength);
		}

		packedUnicodeString.szCharacters = pParser->GetStringW(packedUnicodeString.cchExtendedLength ? packedUnicodeString.cchExtendedLength : packedUnicodeString.cchLength);
	}

	return packedUnicodeString;
}

void PropertyDefinitionStream::Parse()
{
	m_Parser.GetWORD(&m_wVersion);
	m_Parser.GetDWORD(&m_dwFieldDefinitionCount);
	if (m_dwFieldDefinitionCount && m_dwFieldDefinitionCount < _MaxEntriesLarge)
	{
		for (DWORD iDef = 0; iDef < m_dwFieldDefinitionCount; iDef++)
		{
			FieldDefinition fieldDefinition;
			m_Parser.GetDWORD(&fieldDefinition.dwFlags);
			m_Parser.GetWORD(&fieldDefinition.wVT);
			m_Parser.GetDWORD(&fieldDefinition.dwDispid);
			m_Parser.GetWORD(&fieldDefinition.wNmidNameLength);
			fieldDefinition.szNmidName = m_Parser.GetStringW(fieldDefinition.wNmidNameLength);

			fieldDefinition.pasNameANSI = ReadPackedAnsiString(&m_Parser);
			fieldDefinition.pasFormulaANSI = ReadPackedAnsiString(&m_Parser);
			fieldDefinition.pasValidationRuleANSI = ReadPackedAnsiString(&m_Parser);
			fieldDefinition.pasValidationTextANSI = ReadPackedAnsiString(&m_Parser);
			fieldDefinition.pasErrorANSI = ReadPackedAnsiString(&m_Parser);

			if (PropDefV2 == m_wVersion)
			{
				m_Parser.GetDWORD(&fieldDefinition.dwInternalType);

				// Have to count how many skip blocks are here.
				// The only way to do that is to parse them. So we parse once without storing, allocate, then reparse.
				auto stBookmark = m_Parser.GetCurrentOffset();

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

				fieldDefinition.dwSkipBlockCount = dwSkipBlockCount;
				if (fieldDefinition.dwSkipBlockCount &&
					fieldDefinition.dwSkipBlockCount < _MaxEntriesSmall)
				{
					for (DWORD iSkip = 0; iSkip < fieldDefinition.dwSkipBlockCount; iSkip++)
					{
						SkipBlock skipBlock;
						m_Parser.GetDWORD(&skipBlock.dwSize);
						skipBlock.lpbContent = m_Parser.GetBYTES(skipBlock.dwSize, _MaxBytes);
						fieldDefinition.psbSkipBlocks.push_back(skipBlock);
					}
				}
			}

			m_pfdFieldDefinitions.push_back(fieldDefinition);
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
		0xFF == ppasString->cchLength ? ppasString->cchExtendedLength : ppasString->cchLength);
	if (ppasString->szCharacters.length())
	{
		szPackedAnsiString += formatmessage(IDS_PROPDEFPACKEDSTRINGCHARS);
		szPackedAnsiString += stringTowstring(ppasString->szCharacters);
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
		0xFF == ppusString->cchLength ? ppusString->cchExtendedLength : ppusString->cchLength);
	if (ppusString->szCharacters.length())
	{
		szPackedUnicodeString += formatmessage(IDS_PROPDEFPACKEDSTRINGCHARS);
		szPackedUnicodeString += ppusString->szCharacters;
	}

	return szPackedUnicodeString;
}

_Check_return_ wstring PropertyDefinitionStream::ToStringInternal()
{
	wstring szPropertyDefinitionStream;
	auto szVersion = InterpretFlags(flagPropDefVersion, m_wVersion);

	szPropertyDefinitionStream = formatmessage(IDS_PROPDEFHEADER,
		m_wVersion, szVersion.c_str(),
		m_dwFieldDefinitionCount);

	for (DWORD iDef = 0; iDef < m_pfdFieldDefinitions.size(); iDef++)
	{
		auto szFlags = InterpretFlags(flagPDOFlag, m_pfdFieldDefinitions[iDef].dwFlags);
		auto szVarEnum = InterpretFlags(flagVarEnum, m_pfdFieldDefinitions[iDef].wVT);
		szPropertyDefinitionStream += formatmessage(IDS_PROPDEFFIELDHEADER,
			iDef,
			m_pfdFieldDefinitions[iDef].dwFlags, szFlags.c_str(),
			m_pfdFieldDefinitions[iDef].wVT, szVarEnum.c_str(),
			m_pfdFieldDefinitions[iDef].dwDispid);

		if (m_pfdFieldDefinitions[iDef].dwDispid)
		{
			if (m_pfdFieldDefinitions[iDef].dwDispid < 0x8000)
			{
				wstring szExactMatch;
				wstring szDispidName;
				PropTagToPropName(m_pfdFieldDefinitions[iDef].dwDispid, false, szExactMatch, szDispidName);
				if (!szDispidName.empty())
				{
					szPropertyDefinitionStream += formatmessage(IDS_PROPDEFDISPIDTAG, szDispidName.c_str());
				}
			}
			else
			{
				wstring szDispidName;
				MAPINAMEID mnid = { nullptr };
				mnid.lpguid = nullptr;
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
		szPropertyDefinitionStream += m_pfdFieldDefinitions[iDef].szNmidName;

		auto szTab1 = formatmessage(IDS_PROPDEFTAB1);

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

			for (DWORD iSkip = 0; iSkip < m_pfdFieldDefinitions[iDef].psbSkipBlocks.size(); iSkip++)
			{
				szPropertyDefinitionStream += formatmessage(IDS_PROPDEFSKIPBLOCK,
					iSkip,
					m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].dwSize);

				if (m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContent.size())
				{
					if (0 == iSkip)
					{
						// Parse this on the fly
						CBinaryParser ParserFirstBlock(
							m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContent.size(),
							m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContent.data());
						auto pusString = ReadPackedUnicodeString(&ParserFirstBlock);
						szPropertyDefinitionStream += formatmessage(IDS_PROPDEFTAB2) + PackedUnicodeStringToString(IDS_PROPDEFFIELDNAME, &pusString);
					}
					else
					{
						szPropertyDefinitionStream += formatmessage(IDS_PROPDEFCONTENT) + BinToHexString(m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContent, true);
					}
				}
			}
		}
	}

	return szPropertyDefinitionStream;
}