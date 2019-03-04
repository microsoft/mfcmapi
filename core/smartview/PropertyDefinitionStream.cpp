#include <core/stdafx.h>
#include <core/smartview/PropertyDefinitionStream.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <core/utility/strings.h>
#include <core/mapi/cache/namedPropCache.h>
#include <core/interpret/proptags.h>

namespace smartview
{
	PackedAnsiString ReadPackedAnsiString(_In_ binaryParser& parser)
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

	PackedUnicodeString ReadPackedUnicodeString(_In_ binaryParser& parser)
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

	_Check_return_ block
	PackedAnsiStringToBlock(_In_ const std::wstring& szFieldName, _In_ const PackedAnsiString& pasString)
	{
		auto data = pasString.cchLength;

		if (0xFF == pasString.cchLength)
		{
			data.setText(L"\t%1!ws!: Length = 0x%2!04X!", szFieldName.c_str(), pasString.cchExtendedLength.getData());
			data.setSize(pasString.cchLength.getSize() + pasString.cchExtendedLength.getSize());
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
	PackedUnicodeStringToBlock(_In_ const std::wstring& szFieldName, _In_ const PackedUnicodeString& pusString)
	{
		auto data = pusString.cchLength;

		if (0xFF == pusString.cchLength)
		{
			data.setText(L"\t%1!ws!: Length = 0x%2!04X!", szFieldName.c_str(), pusString.cchExtendedLength.getData());
			data.setSize(pusString.cchLength.getSize() + pusString.cchExtendedLength.getSize());
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

		auto iDef = DWORD{};
		for (const auto& def : m_pfdFieldDefinitions)
		{
			terminateBlock();
			auto fieldDef = block{};
			fieldDef.setText(L"Definition: %1!d!\r\n", iDef);

			auto szFlags = flags::InterpretFlags(flagPDOFlag, def.dwFlags);
			fieldDef.addBlock(def.dwFlags, L"\tFlags = 0x%1!08X! = %2!ws!\r\n", def.dwFlags.getData(), szFlags.c_str());
			auto szVarEnum = flags::InterpretFlags(flagVarEnum, def.wVT);
			fieldDef.addBlock(def.wVT, L"\tVT = 0x%1!04X! = %2!ws!\r\n", def.wVT.getData(), szVarEnum.c_str());
			fieldDef.addBlock(def.dwDispid, L"\tDispID = 0x%1!08X!", def.dwDispid.getData());

			if (def.dwDispid)
			{
				if (def.dwDispid < 0x8000)
				{
					auto propTagNames = proptags::PropTagToPropName(def.dwDispid, false);
					if (!propTagNames.bestGuess.empty())
					{
						fieldDef.addBlock(def.dwDispid, L" = %1!ws!", propTagNames.bestGuess.c_str());
					}

					if (!propTagNames.otherMatches.empty())
					{
						fieldDef.addBlock(def.dwDispid, L": (%1!ws!)", propTagNames.otherMatches.c_str());
					}
				}
				else
				{
					std::wstring szDispidName;
					MAPINAMEID mnid = {};
					mnid.lpguid = nullptr;
					mnid.ulKind = MNID_ID;
					mnid.Kind.lID = def.dwDispid;
					szDispidName = strings::join(cache::NameIDToPropNames(&mnid), L", ");
					if (!szDispidName.empty())
					{
						fieldDef.addBlock(def.dwDispid, L" = %1!ws!", szDispidName.c_str());
					}
				}
			}

			fieldDef.terminateBlock();
			fieldDef.addBlock(def.wNmidNameLength, L"\tNmidNameLength = 0x%1!04X!\r\n", def.wNmidNameLength.getData());
			fieldDef.addBlock(def.szNmidName, L"\tNmidName = %1!ws!\r\n", def.szNmidName.c_str());

			fieldDef.addBlock(PackedAnsiStringToBlock(L"NameAnsi", def.pasNameANSI));
			fieldDef.addBlock(PackedAnsiStringToBlock(L"FormulaANSI", def.pasFormulaANSI));
			fieldDef.addBlock(PackedAnsiStringToBlock(L"ValidationRuleANSI", def.pasValidationRuleANSI));
			fieldDef.addBlock(PackedAnsiStringToBlock(L"ValidationTextANSI", def.pasValidationTextANSI));
			fieldDef.addBlock(PackedAnsiStringToBlock(L"ErrorANSI", def.pasErrorANSI));

			if (PropDefV2 == m_wVersion)
			{
				szFlags = flags::InterpretFlags(flagInternalType, def.dwInternalType);
				fieldDef.addBlock(
					def.dwInternalType,
					L"\tInternalType = 0x%1!08X! = %2!ws!\r\n",
					def.dwInternalType.getData(),
					szFlags.c_str());
				fieldDef.addHeader(L"\tSkipBlockCount = %1!d!", def.dwSkipBlockCount);

				auto iSkip = DWORD{};
				for (const auto& sb : def.psbSkipBlocks)
				{
					fieldDef.terminateBlock();
					auto skipBlock = block{};
					skipBlock.setText(L"\tSkipBlock: %1!d!\r\n", iSkip);
					skipBlock.addBlock(sb.dwSize, L"\t\tSize = 0x%1!08X!", sb.dwSize.getData());

					if (0 == iSkip)
					{
						skipBlock.terminateBlock();
						skipBlock.addBlock(PackedUnicodeStringToBlock(L"\tFieldName", sb.lpbContentText));
					}
					else if (!sb.lpbContent.empty())
					{
						skipBlock.terminateBlock();
						skipBlock.addHeader(L"\t\tContent = ");
						skipBlock.addBlock(sb.lpbContent);
					}

					fieldDef.addBlock(skipBlock);
					iSkip++;
				}
			}

			addBlock(fieldDef);
			iDef++;
		}
	}
} // namespace smartview