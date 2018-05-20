#include "StdAfx.h"
#include "PropertyDefinitionStream.h"
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
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
			packedAnsiString.cchLength = pParser->Get<BYTE>();
			if (0xFF == packedAnsiString.cchLength)
			{
				packedAnsiString.cchExtendedLength = pParser->Get<WORD>();
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
			packedUnicodeString.cchLength = pParser->Get<BYTE>();
			if (0xFF == packedUnicodeString.cchLength)
			{
				packedUnicodeString.cchExtendedLength = pParser->Get<WORD>();
			}

			packedUnicodeString.szCharacters = pParser->GetStringW(packedUnicodeString.cchExtendedLength ? packedUnicodeString.cchExtendedLength : packedUnicodeString.cchLength);
		}

		return packedUnicodeString;
	}

	void PropertyDefinitionStream::Parse()
	{
		m_wVersion = m_Parser.Get<WORD>();
		m_dwFieldDefinitionCount = m_Parser.Get<DWORD>();
		if (m_dwFieldDefinitionCount && m_dwFieldDefinitionCount < _MaxEntriesLarge)
		{
			for (DWORD iDef = 0; iDef < m_dwFieldDefinitionCount; iDef++)
			{
				FieldDefinition fieldDefinition;
				fieldDefinition.dwFlags = m_Parser.Get<DWORD>();
				fieldDefinition.wVT = m_Parser.Get<WORD>();
				fieldDefinition.dwDispid = m_Parser.Get<DWORD>();
				fieldDefinition.wNmidNameLength = m_Parser.Get<WORD>();
				fieldDefinition.szNmidName = m_Parser.GetStringW(fieldDefinition.wNmidNameLength);

				fieldDefinition.pasNameANSI = ReadPackedAnsiString(&m_Parser);
				fieldDefinition.pasFormulaANSI = ReadPackedAnsiString(&m_Parser);
				fieldDefinition.pasValidationRuleANSI = ReadPackedAnsiString(&m_Parser);
				fieldDefinition.pasValidationTextANSI = ReadPackedAnsiString(&m_Parser);
				fieldDefinition.pasErrorANSI = ReadPackedAnsiString(&m_Parser);

				if (PropDefV2 == m_wVersion)
				{
					fieldDefinition.dwInternalType = m_Parser.Get<DWORD>();

					// Have to count how many skip blocks are here.
					// The only way to do that is to parse them. So we parse once without storing, allocate, then reparse.
					const auto stBookmark = m_Parser.GetCurrentOffset();

					DWORD dwSkipBlockCount = 0;

					for (;;)
					{
						dwSkipBlockCount++;
						const auto dwBlock = m_Parser.Get<DWORD>();
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
							skipBlock.dwSize = m_Parser.Get<DWORD>();
							skipBlock.lpbContent = m_Parser.GetBYTES(skipBlock.dwSize, _MaxBytes);
							fieldDefinition.psbSkipBlocks.push_back(skipBlock);
						}
					}
				}

				m_pfdFieldDefinitions.push_back(fieldDefinition);
			}
		}
	}

	_Check_return_ std::wstring PackedAnsiStringToString(DWORD dwFlags, _In_ PackedAnsiString* ppasString)
	{
		if (!ppasString) return L"";

		std::wstring szFieldName;

		szFieldName = strings::formatmessage(dwFlags);

		auto szPackedAnsiString = strings::formatmessage(IDS_PROPDEFPACKEDSTRINGLEN,
			szFieldName.c_str(),
			0xFF == ppasString->cchLength ? ppasString->cchExtendedLength : ppasString->cchLength);
		if (ppasString->szCharacters.length())
		{
			szPackedAnsiString += strings::formatmessage(IDS_PROPDEFPACKEDSTRINGCHARS);
			szPackedAnsiString += strings::stringTowstring(ppasString->szCharacters);
		}

		return szPackedAnsiString;
	}

	_Check_return_ std::wstring PackedUnicodeStringToString(DWORD dwFlags, _In_ PackedUnicodeString* ppusString)
	{
		if (!ppusString) return L"";

		auto szFieldName = strings::formatmessage(dwFlags);

		auto szPackedUnicodeString = strings::formatmessage(IDS_PROPDEFPACKEDSTRINGLEN,
			szFieldName.c_str(),
			0xFF == ppusString->cchLength ? ppusString->cchExtendedLength : ppusString->cchLength);
		if (ppusString->szCharacters.length())
		{
			szPackedUnicodeString += strings::formatmessage(IDS_PROPDEFPACKEDSTRINGCHARS);
			szPackedUnicodeString += ppusString->szCharacters;
		}

		return szPackedUnicodeString;
	}

	_Check_return_ std::wstring PropertyDefinitionStream::ToStringInternal()
	{
		auto szVersion = InterpretFlags(flagPropDefVersion, m_wVersion);

		std::wstring szPropertyDefinitionStream = strings::formatmessage(IDS_PROPDEFHEADER,
			m_wVersion, szVersion.c_str(),
			m_dwFieldDefinitionCount);

		for (DWORD iDef = 0; iDef < m_pfdFieldDefinitions.size(); iDef++)
		{
			auto szFlags = InterpretFlags(flagPDOFlag, m_pfdFieldDefinitions[iDef].dwFlags);
			auto szVarEnum = InterpretFlags(flagVarEnum, m_pfdFieldDefinitions[iDef].wVT);
			szPropertyDefinitionStream += strings::formatmessage(IDS_PROPDEFFIELDHEADER,
				iDef,
				m_pfdFieldDefinitions[iDef].dwFlags, szFlags.c_str(),
				m_pfdFieldDefinitions[iDef].wVT, szVarEnum.c_str(),
				m_pfdFieldDefinitions[iDef].dwDispid);

			if (m_pfdFieldDefinitions[iDef].dwDispid)
			{
				if (m_pfdFieldDefinitions[iDef].dwDispid < 0x8000)
				{
					auto propTagNames = PropTagToPropName(m_pfdFieldDefinitions[iDef].dwDispid, false);
					if (!propTagNames.bestGuess.empty())
					{
						szPropertyDefinitionStream += strings::formatmessage(IDS_PROPDEFDISPIDTAG, propTagNames.bestGuess.c_str());
					}

					if (!propTagNames.otherMatches.empty())
					{
						szPropertyDefinitionStream += strings::formatmessage(IDS_PROPDEFDISPIDOTHER, propTagNames.otherMatches.c_str());
					}
				}
				else
				{
					std::wstring szDispidName;
					MAPINAMEID mnid = { nullptr };
					mnid.lpguid = nullptr;
					mnid.ulKind = MNID_ID;
					mnid.Kind.lID = m_pfdFieldDefinitions[iDef].dwDispid;
					szDispidName = strings::join(NameIDToPropNames(&mnid), L", ");
					if (!szDispidName.empty())
					{
						szPropertyDefinitionStream += strings::formatmessage(IDS_PROPDEFDISPIDNAMED, szDispidName.c_str());
					}
				}
			}

			szPropertyDefinitionStream += strings::formatmessage(IDS_PROPDEFNMIDNAME,
				m_pfdFieldDefinitions[iDef].wNmidNameLength);
			szPropertyDefinitionStream += m_pfdFieldDefinitions[iDef].szNmidName;

			const auto szTab1 = strings::formatmessage(IDS_PROPDEFTAB1);

			szPropertyDefinitionStream += szTab1 + PackedAnsiStringToString(IDS_PROPDEFNAME, &m_pfdFieldDefinitions[iDef].pasNameANSI);
			szPropertyDefinitionStream += szTab1 + PackedAnsiStringToString(IDS_PROPDEFFORUMULA, &m_pfdFieldDefinitions[iDef].pasFormulaANSI);
			szPropertyDefinitionStream += szTab1 + PackedAnsiStringToString(IDS_PROPDEFVRULE, &m_pfdFieldDefinitions[iDef].pasValidationRuleANSI);
			szPropertyDefinitionStream += szTab1 + PackedAnsiStringToString(IDS_PROPDEFVTEXT, &m_pfdFieldDefinitions[iDef].pasValidationTextANSI);
			szPropertyDefinitionStream += szTab1 + PackedAnsiStringToString(IDS_PROPDEFERROR, &m_pfdFieldDefinitions[iDef].pasErrorANSI);

			if (PropDefV2 == m_wVersion)
			{
				szFlags = InterpretFlags(flagInternalType, m_pfdFieldDefinitions[iDef].dwInternalType);
				szPropertyDefinitionStream += strings::formatmessage(IDS_PROPDEFINTERNALTYPE,
					m_pfdFieldDefinitions[iDef].dwInternalType, szFlags.c_str(),
					m_pfdFieldDefinitions[iDef].dwSkipBlockCount);

				for (DWORD iSkip = 0; iSkip < m_pfdFieldDefinitions[iDef].psbSkipBlocks.size(); iSkip++)
				{
					szPropertyDefinitionStream += strings::formatmessage(IDS_PROPDEFSKIPBLOCK,
						iSkip,
						m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].dwSize);

					if (!m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContent.empty())
					{
						if (0 == iSkip)
						{
							// Parse this on the fly
							CBinaryParser ParserFirstBlock(
								m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContent.size(),
								m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContent.data());
							auto pusString = ReadPackedUnicodeString(&ParserFirstBlock);
							szPropertyDefinitionStream += strings::formatmessage(IDS_PROPDEFTAB2) + PackedUnicodeStringToString(IDS_PROPDEFFIELDNAME, &pusString);
						}
						else
						{
							szPropertyDefinitionStream += strings::formatmessage(IDS_PROPDEFCONTENT) + strings::BinToHexString(m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContent, true);
						}
					}
				}
			}
		}

		return szPropertyDefinitionStream;
	}
}