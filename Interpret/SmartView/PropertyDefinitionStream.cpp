#include <StdAfx.h>
#include <Interpret/SmartView/PropertyDefinitionStream.h>
#include <Interpret/InterpretProp.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <core/utility/strings.h>
#include <core/mapi/cache/namedPropCache.h>

namespace smartview
{
	PackedAnsiString ReadPackedAnsiString(_In_ CBinaryParser& parser)
	{
		PackedAnsiString packedAnsiString;
		packedAnsiString.cchLength = parser.Get<BYTE>();
		if (0xFF == packedAnsiString.cchLength)
		{
			packedAnsiString.cchExtendedLength = parser.Get<WORD>();
		}

		packedAnsiString.szCharacters = parser.GetStringA(
			packedAnsiString.cchExtendedLength ? packedAnsiString.cchExtendedLength.getData()
											   : packedAnsiString.cchLength.getData());

		return packedAnsiString;
	}

	PackedUnicodeString ReadPackedUnicodeString(_In_ CBinaryParser& parser)
	{
		PackedUnicodeString packedUnicodeString;
		packedUnicodeString.cchLength = parser.Get<BYTE>();
		if (0xFF == packedUnicodeString.cchLength)
		{
			packedUnicodeString.cchExtendedLength = parser.Get<WORD>();
		}

		packedUnicodeString.szCharacters = parser.GetStringW(
			packedUnicodeString.cchExtendedLength ? packedUnicodeString.cchExtendedLength.getData()
												  : packedUnicodeString.cchLength.getData());

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

				fieldDefinition.pasNameANSI = ReadPackedAnsiString(m_Parser);
				fieldDefinition.pasFormulaANSI = ReadPackedAnsiString(m_Parser);
				fieldDefinition.pasValidationRuleANSI = ReadPackedAnsiString(m_Parser);
				fieldDefinition.pasValidationTextANSI = ReadPackedAnsiString(m_Parser);
				fieldDefinition.pasErrorANSI = ReadPackedAnsiString(m_Parser);

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
						m_Parser.advance(dwBlock);
					}

					m_Parser.SetCurrentOffset(stBookmark); // We're done with our first pass, restore the bookmark

					fieldDefinition.dwSkipBlockCount = dwSkipBlockCount;
					if (fieldDefinition.dwSkipBlockCount && fieldDefinition.dwSkipBlockCount < _MaxEntriesSmall)
					{
						fieldDefinition.psbSkipBlocks.reserve(fieldDefinition.dwSkipBlockCount);
						for (DWORD iSkip = 0; iSkip < fieldDefinition.dwSkipBlockCount; iSkip++)
						{
							SkipBlock skipBlock;
							skipBlock.dwSize = m_Parser.Get<DWORD>();
							if (iSkip == 0)
							{
								skipBlock.lpbContentText = ReadPackedUnicodeString(m_Parser);
							}
							else
							{
								skipBlock.lpbContent = m_Parser.GetBYTES(skipBlock.dwSize, _MaxBytes);
							}

							fieldDefinition.psbSkipBlocks.push_back(skipBlock);
						}
					}
				}

				m_pfdFieldDefinitions.push_back(fieldDefinition);
			}
		}
	}

	_Check_return_ block PackedAnsiStringToBlock(_In_ const std::wstring& szFieldName, _In_ PackedAnsiString& pasString)
	{
		auto data = pasString.cchLength;

		if (0xFF == pasString.cchLength)
		{
			data.setText(L"\t%1!ws!: Length = 0x%2!04X!", szFieldName.c_str(), pasString.cchExtendedLength.getData());
		}
		else
		{
			data.setText(L"\t%1!ws!: Length = 0x%2!04X!", szFieldName.c_str(), pasString.cchLength.getData());
		}

		if (pasString.szCharacters.length())
		{
			data.addHeader(L" Characters = ");
			data.addBlock(pasString.szCharacters, strings::stringTowstring(pasString.szCharacters));
		}

		data.terminateBlock();
		return data;
	}

	_Check_return_ block
	PackedUnicodeStringToBlock(_In_ const std::wstring& szFieldName, _In_ PackedUnicodeString& pusString)
	{
		auto data = pusString.cchLength;

		if (0xFF == pusString.cchLength)
		{
			data.setText(L"\t%1!ws!: Length = 0x%2!04X!", szFieldName.c_str(), pusString.cchExtendedLength.getData());
		}
		else
		{
			data.setText(L"\t%1!ws!: Length = 0x%2!04X!", szFieldName.c_str(), pusString.cchLength.getData());
		}

		if (pusString.szCharacters.length())
		{
			data.addHeader(L" Characters = ");
			data.addBlock(pusString.szCharacters, pusString.szCharacters);
		}

		data.terminateBlock();
		return data;
	}

	void PropertyDefinitionStream::ParseBlocks()
	{
		setRoot(L"Property Definition Stream\r\n");
		auto szVersion = flags::InterpretFlags(flagPropDefVersion, m_wVersion);
		addBlock(m_wVersion, L"Version = 0x%1!04X! = %2!ws!\r\n", m_wVersion.getData(), szVersion.c_str());
		addBlock(m_dwFieldDefinitionCount, L"FieldDefinitionCount = 0x%1!08X!", m_dwFieldDefinitionCount.getData());

		for (DWORD iDef = 0; iDef < m_pfdFieldDefinitions.size(); iDef++)
		{
			terminateBlock();
			addHeader(L"Definition: %1!d!\r\n", iDef);

			auto szFlags = flags::InterpretFlags(flagPDOFlag, m_pfdFieldDefinitions[iDef].dwFlags);
			addBlock(
				m_pfdFieldDefinitions[iDef].dwFlags,
				L"\tFlags = 0x%1!08X! = %2!ws!\r\n",
				m_pfdFieldDefinitions[iDef].dwFlags.getData(),
				szFlags.c_str());
			auto szVarEnum = flags::InterpretFlags(flagVarEnum, m_pfdFieldDefinitions[iDef].wVT);
			addBlock(
				m_pfdFieldDefinitions[iDef].wVT,
				L"\tVT = 0x%1!04X! = %2!ws!\r\n",
				m_pfdFieldDefinitions[iDef].wVT.getData(),
				szVarEnum.c_str());
			addBlock(
				m_pfdFieldDefinitions[iDef].dwDispid,
				L"\tDispID = 0x%1!08X!",
				m_pfdFieldDefinitions[iDef].dwDispid.getData());

			if (m_pfdFieldDefinitions[iDef].dwDispid)
			{
				if (m_pfdFieldDefinitions[iDef].dwDispid < 0x8000)
				{
					auto propTagNames = interpretprop::PropTagToPropName(m_pfdFieldDefinitions[iDef].dwDispid, false);
					if (!propTagNames.bestGuess.empty())
					{
						addBlock(m_pfdFieldDefinitions[iDef].dwDispid, L" = %1!ws!", propTagNames.bestGuess.c_str());
					}

					if (!propTagNames.otherMatches.empty())
					{
						addBlock(
							m_pfdFieldDefinitions[iDef].dwDispid, L": (%1!ws!)", propTagNames.otherMatches.c_str());
					}
				}
				else
				{
					std::wstring szDispidName;
					MAPINAMEID mnid = {};
					mnid.lpguid = nullptr;
					mnid.ulKind = MNID_ID;
					mnid.Kind.lID = m_pfdFieldDefinitions[iDef].dwDispid;
					szDispidName = strings::join(cache::NameIDToPropNames(&mnid), L", ");
					if (!szDispidName.empty())
					{
						addBlock(m_pfdFieldDefinitions[iDef].dwDispid, L" = %1!ws!", szDispidName.c_str());
					}
				}
			}

			terminateBlock();
			addBlock(
				m_pfdFieldDefinitions[iDef].wNmidNameLength,
				L"\tNmidNameLength = 0x%1!04X!\r\n",
				m_pfdFieldDefinitions[iDef].wNmidNameLength.getData());
			addBlock(
				m_pfdFieldDefinitions[iDef].szNmidName,
				L"\tNmidName = %1!ws!\r\n",
				m_pfdFieldDefinitions[iDef].szNmidName.c_str());

			addBlock(PackedAnsiStringToBlock(L"NameAnsi", m_pfdFieldDefinitions[iDef].pasNameANSI));
			addBlock(PackedAnsiStringToBlock(L"FormulaANSI", m_pfdFieldDefinitions[iDef].pasFormulaANSI));
			addBlock(PackedAnsiStringToBlock(L"ValidationRuleANSI", m_pfdFieldDefinitions[iDef].pasValidationRuleANSI));
			addBlock(PackedAnsiStringToBlock(L"ValidationTextANSI", m_pfdFieldDefinitions[iDef].pasValidationTextANSI));
			addBlock(PackedAnsiStringToBlock(L"ErrorANSI", m_pfdFieldDefinitions[iDef].pasErrorANSI));

			if (PropDefV2 == m_wVersion)
			{
				szFlags = flags::InterpretFlags(flagInternalType, m_pfdFieldDefinitions[iDef].dwInternalType);
				addBlock(
					m_pfdFieldDefinitions[iDef].dwInternalType,
					L"\tInternalType = 0x%1!08X! = %2!ws!\r\n",
					m_pfdFieldDefinitions[iDef].dwInternalType.getData(),
					szFlags.c_str());
				addHeader(L"\tSkipBlockCount = %1!d!", m_pfdFieldDefinitions[iDef].dwSkipBlockCount);

				for (DWORD iSkip = 0; iSkip < m_pfdFieldDefinitions[iDef].psbSkipBlocks.size(); iSkip++)
				{
					terminateBlock();
					addHeader(L"\tSkipBlock: %1!d!\r\n", iSkip);
					addBlock(
						m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].dwSize,
						L"\t\tSize = 0x%1!08X!",
						m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].dwSize.getData());

					if (0 == iSkip)
					{
						terminateBlock();
						addBlock(PackedUnicodeStringToBlock(
							L"\tFieldName", m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContentText));
					}
					else if (!m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContent.empty())

					{
						terminateBlock();
						addHeader(L"\t\tContent = ");
						addBlock(m_pfdFieldDefinitions[iDef].psbSkipBlocks[iSkip].lpbContent);
					}
				}
			}
		}
	}
} // namespace smartview