#include <StdAfx.h>
#include <Interpret/SmartView/PropertyDefinitionStream.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
	PropertyDefinitionStream::PropertyDefinitionStream() {}

	PackedAnsiString ReadPackedAnsiString(_In_ CBinaryParser* pParser)
	{
		PackedAnsiString packedAnsiString;
		if (pParser)
		{
			packedAnsiString.cchLength = pParser->GetBlock<BYTE>();
			if (0xFF == packedAnsiString.cchLength.getData())
			{
				packedAnsiString.cchExtendedLength = pParser->GetBlock<WORD>();
			}

			packedAnsiString.szCharacters = pParser->GetBlockStringA(
				packedAnsiString.cchExtendedLength.getData() ? packedAnsiString.cchExtendedLength.getData()
															 : packedAnsiString.cchLength.getData());
		}

		return packedAnsiString;
	}

	PackedUnicodeString ReadPackedUnicodeString(_In_ CBinaryParser* pParser)
	{
		PackedUnicodeString packedUnicodeString;
		if (pParser)
		{
			packedUnicodeString.cchLength = pParser->GetBlock<BYTE>();
			if (0xFF == packedUnicodeString.cchLength.getData())
			{
				packedUnicodeString.cchExtendedLength = pParser->GetBlock<WORD>();
			}

			packedUnicodeString.szCharacters = pParser->GetBlockStringW(
				packedUnicodeString.cchExtendedLength.getData() ? packedUnicodeString.cchExtendedLength.getData()
																: packedUnicodeString.cchLength.getData());
		}

		return packedUnicodeString;
	}

	void PropertyDefinitionStream::Parse()
	{
		m_wVersion = m_Parser.GetBlock<WORD>();
		m_dwFieldDefinitionCount = m_Parser.GetBlock<DWORD>();
		if (m_dwFieldDefinitionCount.getData() && m_dwFieldDefinitionCount.getData() < _MaxEntriesLarge)
		{
			for (DWORD iDef = 0; iDef < m_dwFieldDefinitionCount.getData(); iDef++)
			{
				FieldDefinition fieldDefinition;
				fieldDefinition.dwFlags = m_Parser.GetBlock<DWORD>();
				fieldDefinition.wVT = m_Parser.GetBlock<WORD>();
				fieldDefinition.dwDispid = m_Parser.GetBlock<DWORD>();
				fieldDefinition.wNmidNameLength = m_Parser.GetBlock<WORD>();
				fieldDefinition.szNmidName = m_Parser.GetBlockStringW(fieldDefinition.wNmidNameLength.getData());

				fieldDefinition.pasNameANSI = ReadPackedAnsiString(&m_Parser);
				fieldDefinition.pasFormulaANSI = ReadPackedAnsiString(&m_Parser);
				fieldDefinition.pasValidationRuleANSI = ReadPackedAnsiString(&m_Parser);
				fieldDefinition.pasValidationTextANSI = ReadPackedAnsiString(&m_Parser);
				fieldDefinition.pasErrorANSI = ReadPackedAnsiString(&m_Parser);

				if (PropDefV2 == m_wVersion.getData())
				{
					fieldDefinition.dwInternalType = m_Parser.GetBlock<DWORD>();

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
					if (fieldDefinition.dwSkipBlockCount && fieldDefinition.dwSkipBlockCount < _MaxEntriesSmall)
					{
						for (DWORD iSkip = 0; iSkip < fieldDefinition.dwSkipBlockCount; iSkip++)
						{
							SkipBlock skipBlock;
							skipBlock.dwSize = m_Parser.GetBlock<DWORD>();
							skipBlock.lpbContent = m_Parser.GetBlockBYTES(skipBlock.dwSize.getData(), _MaxBytes);
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

		auto szPackedAnsiString = strings::formatmessage(
			IDS_PROPDEFPACKEDSTRINGLEN,
			szFieldName.c_str(),
			0xFF == ppasString->cchLength.getData() ? ppasString->cchExtendedLength.getData()
													: ppasString->cchLength.getData());
		if (ppasString->szCharacters.getData().length())
		{
			szPackedAnsiString += strings::formatmessage(IDS_PROPDEFPACKEDSTRINGCHARS);
			szPackedAnsiString += strings::stringTowstring(ppasString->szCharacters.getData());
		}

		return szPackedAnsiString;
	}

	_Check_return_ std::wstring PackedUnicodeStringToString(DWORD dwFlags, _In_ PackedUnicodeString* ppusString)
	{
		if (!ppusString) return L"";

		auto szFieldName = strings::formatmessage(dwFlags);

		auto szPackedUnicodeString = strings::formatmessage(
			IDS_PROPDEFPACKEDSTRINGLEN,
			szFieldName.c_str(),
			0xFF == ppusString->cchLength.getData() ? ppusString->cchExtendedLength.getData()
													: ppusString->cchLength.getData());
		if (ppusString->szCharacters.getData().length())
		{
			szPackedUnicodeString += strings::formatmessage(IDS_PROPDEFPACKEDSTRINGCHARS);
			szPackedUnicodeString += ppusString->szCharacters.getData();
		}

		return szPackedUnicodeString;
	}

	_Check_return_ std::wstring PropertyDefinitionStream::ToStringInternal()
	{
		auto szVersion = interpretprop::InterpretFlags(flagPropDefVersion, m_wVersion.getData());

		auto szPropertyDefinitionStream = strings::formatmessage(
			IDS_PROPDEFHEADER, m_wVersion.getData(), szVersion.c_str(), m_dwFieldDefinitionCount.getData());

		for (DWORD iDef = 0; iDef < m_pfdFieldDefinitions.size(); iDef++)
		{
			auto szFlags = interpretprop::InterpretFlags(flagPDOFlag, m_pfdFieldDefinitions[iDef].dwFlags.getData());
			auto szVarEnum = interpretprop::InterpretFlags(flagVarEnum, m_pfdFieldDefinitions[iDef].wVT.getData());
			szPropertyDefinitionStream += strings::formatmessage(
				IDS_PROPDEFFIELDHEADER,
				iDef,
				m_pfdFieldDefinitions[iDef].dwFlags.getData(),
				szFlags.c_str(),
				m_pfdFieldDefinitions[iDef].wVT.getData(),
				szVarEnum.c_str(),
				m_pfdFieldDefinitions[iDef].dwDispid.getData());

			if (m_pfdFieldDefinitions[iDef].dwDispid.getData())
			{
				if (m_pfdFieldDefinitions[iDef].dwDispid.getData() < 0x8000)
				{
					auto propTagNames =
						interpretprop::PropTagToPropName(m_pfdFieldDefinitions[iDef].dwDispid.getData(), false);
					if (!propTagNames.bestGuess.empty())
					{
						szPropertyDefinitionStream +=
							strings::formatmessage(IDS_PROPDEFDISPIDTAG, propTagNames.bestGuess.c_str());
					}

					if (!propTagNames.otherMatches.empty())
					{
						szPropertyDefinitionStream +=
							strings::formatmessage(IDS_PROPDEFDISPIDOTHER, propTagNames.otherMatches.c_str());
					}
				}
				else
				{
					std::wstring szDispidName;
					MAPINAMEID mnid = {nullptr};
					mnid.lpguid = nullptr;
					mnid.ulKind = MNID_ID;
					mnid.Kind.lID = m_pfdFieldDefinitions[iDef].dwDispid.getData();
					szDispidName = strings::join(interpretprop::NameIDToPropNames(&mnid), L", ");
					if (!szDispidName.empty())
					{
						szPropertyDefinitionStream +=
							strings::formatmessage(IDS_PROPDEFDISPIDNAMED, szDispidName.c_str());
					}
				}
			}

			szPropertyDefinitionStream +=
				strings::formatmessage(IDS_PROPDEFNMIDNAME, m_pfdFieldDefinitions[iDef].wNmidNameLength.getData());
			szPropertyDefinitionStream += m_pfdFieldDefinitions[iDef].szNmidName.getData();

			const auto szTab1 = strings::formatmessage(IDS_PROPDEFTAB1);

			szPropertyDefinitionStream +=
				szTab1 + PackedAnsiStringToString(IDS_PROPDEFNAME, &m_pfdFieldDefinitions[iDef].pasNameANSI);
			szPropertyDefinitionStream +=
				szTab1 + PackedAnsiStringToString(IDS_PROPDEFFORUMULA, &m_pfdFieldDefinitions[iDef].pasFormulaANSI);
			szPropertyDefinitionStream +=
				szTab1 + PackedAnsiStringToString(IDS_PROPDEFVRULE, &m_pfdFieldDefinitions[iDef].pasValidationRuleANSI);
			szPropertyDefinitionStream +=
				szTab1 + PackedAnsiStringToString(IDS_PROPDEFVTEXT, &m_pfdFieldDefinitions[iDef].pasValidationTextANSI);
			szPropertyDefinitionStream +=
				szTab1 + PackedAnsiStringToString(IDS_PROPDEFERROR, &m_pfdFieldDefinitions[iDef].pasErrorANSI);

			if (PropDefV2 == m_wVersion.getData())
			{
				szFlags = interpretprop::InterpretFlags(
					flagInternalType, m_pfdFieldDefinitions[iDef].dwInternalType.getData());
				szPropertyDefinitionStream += strings::formatmessage(
					IDS_PROPDEFINTERNALTYPE,
					m_pfdFieldDefinitions[iDef].dwInternalType.getData(),
					szFlags.c_str(),
					m_pfdFieldDefinitions[iDef].dwSkipBlockCount);

				for (DWORD iSkip = 0; iSkip < m_pfdFieldDefinitions[iDef].psbSkipBlocks.size(); iSkip++)
				{
					szPropertyDefinitionStream += strings::formatmessage(
						IDS_PROPDEFSKIPBLOCK, iSkip, m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].dwSize.getData());

					if (!m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContent.getData().empty())
					{
						if (0 == iSkip)
						{
							// Parse this on the fly
							CBinaryParser ParserFirstBlock(
								m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContent.getData().size(),
								m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContent.getData().data());
							auto pusString = ReadPackedUnicodeString(&ParserFirstBlock);
							szPropertyDefinitionStream += strings::formatmessage(IDS_PROPDEFTAB2) +
														  PackedUnicodeStringToString(IDS_PROPDEFFIELDNAME, &pusString);
						}
						else
						{
							szPropertyDefinitionStream +=
								strings::formatmessage(IDS_PROPDEFCONTENT) +
								strings::BinToHexString(
									m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContent.getData(), true);
						}
					}
				}
			}
		}

		return szPropertyDefinitionStream;
	}
}