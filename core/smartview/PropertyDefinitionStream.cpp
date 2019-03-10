#include <core/stdafx.h>
#include <core/smartview/PropertyDefinitionStream.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <core/utility/strings.h>
#include <core/mapi/cache/namedPropCache.h>
#include <core/interpret/proptags.h>

namespace smartview
{
	void PackedAnsiString::parse(std::shared_ptr<binaryParser>& parser)
	{
		cchLength = parser->Get<BYTE>();
		if (0xFF == cchLength)
		{
			cchExtendedLength = parser->Get<WORD>();
		}

		szCharacters.parse(parser, cchExtendedLength ? cchExtendedLength.getData() : cchLength.getData());
	}

	void PackedUnicodeString::parse(std::shared_ptr<binaryParser>& parser)
	{
		cchLength = parser->Get<BYTE>();
		if (0xFF == cchLength)
		{
			cchExtendedLength = parser->Get<WORD>();
		}

		szCharacters.parse(parser, cchExtendedLength ? cchExtendedLength.getData() : cchLength.getData());
	}

	SkipBlock::SkipBlock(std::shared_ptr<binaryParser>& parser, DWORD iSkip)
	{
		dwSize = parser->Get<DWORD>();
		if (iSkip == 0)
		{
			lpbContentText.parse(parser);
		}
		else
		{
			lpbContent = parser->GetBYTES(dwSize, _MaxBytes);
		}
	}

	FieldDefinition::FieldDefinition(std::shared_ptr<binaryParser>& parser, WORD version)
	{
		dwFlags = parser->Get<DWORD>();
		wVT = parser->Get<WORD>();
		dwDispid = parser->Get<DWORD>();
		wNmidNameLength = parser->Get<WORD>();
		szNmidName.parse(parser, wNmidNameLength);

		pasNameANSI.parse(parser);
		pasFormulaANSI.parse(parser);
		pasValidationRuleANSI.parse(parser);
		pasValidationTextANSI.parse(parser);
		pasErrorANSI.parse(parser);

		if (version == PropDefV2)
		{
			dwInternalType = parser->Get<DWORD>();

			// Have to count how many skip blocks are here.
			// The only way to do that is to parse them. So we parse once without storing, allocate, then reparse.
			const auto stBookmark = parser->GetCurrentOffset();

			DWORD skipBlockCount = 0;

			for (;;)
			{
				skipBlockCount++;
				const auto dwBlock = parser->Get<DWORD>();
				if (!dwBlock) break; // we hit the last, zero byte block, or the end of the buffer
				parser->advance(dwBlock);
			}

			parser->SetCurrentOffset(stBookmark); // We're done with our first pass, restore the bookmark

			dwSkipBlockCount = skipBlockCount;
			if (dwSkipBlockCount && dwSkipBlockCount < _MaxEntriesSmall)
			{
				psbSkipBlocks.reserve(dwSkipBlockCount);
				for (DWORD iSkip = 0; iSkip < dwSkipBlockCount; iSkip++)
				{
					psbSkipBlocks.emplace_back(std::make_shared<SkipBlock>(parser, iSkip));
				}
			}
		}
	}

	void PropertyDefinitionStream::Parse()
	{
		m_wVersion = m_Parser->Get<WORD>();
		m_dwFieldDefinitionCount = m_Parser->Get<DWORD>();
		if (m_dwFieldDefinitionCount)
		{
			if (m_dwFieldDefinitionCount < _MaxEntriesLarge)
			{
				for (DWORD iDef = 0; iDef < m_dwFieldDefinitionCount; iDef++)
				{
					m_pfdFieldDefinitions.emplace_back(std::make_shared<FieldDefinition>(m_Parser, m_wVersion));
				}
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

			auto szFlags = flags::InterpretFlags(flagPDOFlag, def->dwFlags);
			fieldDef.addBlock(
				def->dwFlags, L"\tFlags = 0x%1!08X! = %2!ws!\r\n", def->dwFlags.getData(), szFlags.c_str());
			auto szVarEnum = flags::InterpretFlags(flagVarEnum, def->wVT);
			fieldDef.addBlock(def->wVT, L"\tVT = 0x%1!04X! = %2!ws!\r\n", def->wVT.getData(), szVarEnum.c_str());
			fieldDef.addBlock(def->dwDispid, L"\tDispID = 0x%1!08X!", def->dwDispid.getData());

			if (def->dwDispid)
			{
				if (def->dwDispid < 0x8000)
				{
					auto propTagNames = proptags::PropTagToPropName(def->dwDispid, false);
					if (!propTagNames.bestGuess.empty())
					{
						fieldDef.addBlock(def->dwDispid, L" = %1!ws!", propTagNames.bestGuess.c_str());
					}

					if (!propTagNames.otherMatches.empty())
					{
						fieldDef.addBlock(def->dwDispid, L": (%1!ws!)", propTagNames.otherMatches.c_str());
					}
				}
				else
				{
					std::wstring szDispidName;
					MAPINAMEID mnid = {};
					mnid.lpguid = nullptr;
					mnid.ulKind = MNID_ID;
					mnid.Kind.lID = def->dwDispid;
					szDispidName = strings::join(cache::NameIDToPropNames(&mnid), L", ");
					if (!szDispidName.empty())
					{
						fieldDef.addBlock(def->dwDispid, L" = %1!ws!", szDispidName.c_str());
					}
				}
			}

			fieldDef.terminateBlock();
			fieldDef.addBlock(
				def->wNmidNameLength, L"\tNmidNameLength = 0x%1!04X!\r\n", def->wNmidNameLength.getData());
			fieldDef.addBlock(def->szNmidName, L"\tNmidName = %1!ws!\r\n", def->szNmidName.c_str());

			fieldDef.addBlock(PackedAnsiStringToBlock(L"NameAnsi", def->pasNameANSI));
			fieldDef.addBlock(PackedAnsiStringToBlock(L"FormulaANSI", def->pasFormulaANSI));
			fieldDef.addBlock(PackedAnsiStringToBlock(L"ValidationRuleANSI", def->pasValidationRuleANSI));
			fieldDef.addBlock(PackedAnsiStringToBlock(L"ValidationTextANSI", def->pasValidationTextANSI));
			fieldDef.addBlock(PackedAnsiStringToBlock(L"ErrorANSI", def->pasErrorANSI));

			if (m_wVersion == PropDefV2)
			{
				szFlags = flags::InterpretFlags(flagInternalType, def->dwInternalType);
				fieldDef.addBlock(
					def->dwInternalType,
					L"\tInternalType = 0x%1!08X! = %2!ws!\r\n",
					def->dwInternalType.getData(),
					szFlags.c_str());
				fieldDef.addHeader(L"\tSkipBlockCount = %1!d!", def->dwSkipBlockCount);

				auto iSkip = DWORD{};
				for (const auto& sb : def->psbSkipBlocks)
				{
					fieldDef.terminateBlock();
					auto skipBlock = block{};
					skipBlock.setText(L"\tSkipBlock: %1!d!\r\n", iSkip);
					skipBlock.addBlock(sb->dwSize, L"\t\tSize = 0x%1!08X!", sb->dwSize.getData());

					if (0 == iSkip)
					{
						skipBlock.terminateBlock();
						skipBlock.addBlock(PackedUnicodeStringToBlock(L"\tFieldName", sb->lpbContentText));
					}
					else if (!sb->lpbContent.empty())
					{
						skipBlock.terminateBlock();
						skipBlock.addHeader(L"\t\tContent = ");
						skipBlock.addBlock(sb->lpbContent);
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