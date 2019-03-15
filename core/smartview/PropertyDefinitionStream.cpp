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
		cchLength.parse<BYTE>(parser);
		if (0xFF == cchLength)
		{
			cchExtendedLength.parse<WORD>(parser);
		}

		szCharacters.parse(parser, cchExtendedLength ? cchExtendedLength.getData() : cchLength.getData());
	}

	void PackedUnicodeString::parse(std::shared_ptr<binaryParser>& parser)
	{
		cchLength.parse<BYTE>(parser);
		if (0xFF == cchLength)
		{
			cchExtendedLength.parse<WORD>(parser);
		}

		szCharacters =
			blockStringW::parse(parser, cchExtendedLength ? cchExtendedLength.getData() : cchLength.getData());
	}

	SkipBlock::SkipBlock(std::shared_ptr<binaryParser>& parser, DWORD iSkip)
	{
		dwSize.parse<DWORD>(parser);
		if (iSkip == 0)
		{
			lpbContentText.parse(parser);
		}
		else
		{
			lpbContent.parse(parser, dwSize, _MaxBytes);
		}
	}

	FieldDefinition::FieldDefinition(std::shared_ptr<binaryParser>& parser, WORD version)
	{
		dwFlags.parse<DWORD>(parser);
		wVT.parse<WORD>(parser);
		dwDispid.parse<DWORD>(parser);
		wNmidNameLength.parse<WORD>(parser);
		szNmidName = blockStringW::parse(parser, wNmidNameLength);

		pasNameANSI.parse(parser);
		pasFormulaANSI.parse(parser);
		pasValidationRuleANSI.parse(parser);
		pasValidationTextANSI.parse(parser);
		pasErrorANSI.parse(parser);

		if (version == PropDefV2)
		{
			dwInternalType.parse<DWORD>(parser);

			// Have to count how many skip blocks are here.
			// The only way to do that is to parse them. So we parse once without storing, allocate, then reparse.
			const auto stBookmark = parser->GetCurrentOffset();

			DWORD skipBlockCount = 0;

			for (;;)
			{
				skipBlockCount++;
				const auto& dwBlock = blockT<DWORD>(parser);
				if (dwBlock == 0) break; // we hit the last, zero byte block, or the end of the buffer
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
		m_wVersion.parse<WORD>(m_Parser);
		m_dwFieldDefinitionCount.parse<DWORD>(m_Parser);
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

	_Check_return_ block& PackedAnsiString::toBlock(_In_ const std::wstring& szFieldName)
	{
		auto& data = cchLength;

		if (0xFF == cchLength)
		{
			data.setText(L"\t%1!ws!: Length = 0x%2!04X!", szFieldName.c_str(), cchExtendedLength.getData());
			data.setSize(cchLength.getSize() + cchExtendedLength.getSize());
		}
		else
		{
			data.setText(L"\t%1!ws!: Length = 0x%2!04X!", szFieldName.c_str(), cchLength.getData());
		}

		if (szCharacters.length())
		{
			data.addHeader(L" Characters = ");
			data.addChild(szCharacters, strings::stringTowstring(szCharacters));
		}

		data.terminateBlock();
		return data;
	}

	_Check_return_ block& PackedUnicodeString::toBlock(_In_ const std::wstring& szFieldName)
	{
		auto& data = cchLength;

		if (0xFF == cchLength)
		{
			data.setText(L"\t%1!ws!: Length = 0x%2!04X!", szFieldName.c_str(), cchExtendedLength.getData());
			data.setSize(cchLength.getSize() + cchExtendedLength.getSize());
		}
		else
		{
			data.setText(L"\t%1!ws!: Length = 0x%2!04X!", szFieldName.c_str(), cchLength.getData());
		}

		if (szCharacters->length())
		{
			data.addHeader(L" Characters = ");
			data.addChild(szCharacters, szCharacters->c_str());
		}

		data.terminateBlock();
		return data;
	}

	void PropertyDefinitionStream::ParseBlocks()
	{
		setRoot(L"Property Definition Stream\r\n");
		auto szVersion = flags::InterpretFlags(flagPropDefVersion, m_wVersion);
		addChild(m_wVersion, L"Version = 0x%1!04X! = %2!ws!\r\n", m_wVersion.getData(), szVersion.c_str());
		addChild(m_dwFieldDefinitionCount, L"FieldDefinitionCount = 0x%1!08X!", m_dwFieldDefinitionCount.getData());

		auto iDef = 0;
		for (const auto& def : m_pfdFieldDefinitions)
		{
			terminateBlock();
			auto fieldDef = std::make_shared<block>();
			fieldDef->setText(L"Definition: %1!d!\r\n", iDef);

			auto szFlags = flags::InterpretFlags(flagPDOFlag, def->dwFlags);
			fieldDef->addChild(
				def->dwFlags, L"\tFlags = 0x%1!08X! = %2!ws!\r\n", def->dwFlags.getData(), szFlags.c_str());
			auto szVarEnum = flags::InterpretFlags(flagVarEnum, def->wVT);
			fieldDef->addChild(def->wVT, L"\tVT = 0x%1!04X! = %2!ws!\r\n", def->wVT.getData(), szVarEnum.c_str());
			fieldDef->addChild(def->dwDispid, L"\tDispID = 0x%1!08X!", def->dwDispid.getData());

			if (def->dwDispid)
			{
				if (def->dwDispid < 0x8000)
				{
					auto propTagNames = proptags::PropTagToPropName(def->dwDispid, false);
					if (!propTagNames.bestGuess.empty())
					{
						fieldDef->addChild(def->dwDispid, L" = %1!ws!", propTagNames.bestGuess.c_str());
					}

					if (!propTagNames.otherMatches.empty())
					{
						fieldDef->addChild(def->dwDispid, L": (%1!ws!)", propTagNames.otherMatches.c_str());
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
						fieldDef->addChild(def->dwDispid, L" = %1!ws!", szDispidName.c_str());
					}
				}
			}

			fieldDef->terminateBlock();
			fieldDef->addChild(
				def->wNmidNameLength, L"\tNmidNameLength = 0x%1!04X!\r\n", def->wNmidNameLength.getData());
			fieldDef->addChild(def->szNmidName, L"\tNmidName = %1!ws!\r\n", def->szNmidName->c_str());

			fieldDef->addChild(def->pasNameANSI.toBlock(L"NameAnsi"));
			fieldDef->addChild(def->pasFormulaANSI.toBlock(L"FormulaANSI"));
			fieldDef->addChild(def->pasValidationRuleANSI.toBlock(L"ValidationRuleANSI"));
			fieldDef->addChild(def->pasValidationTextANSI.toBlock(L"ValidationTextANSI"));
			fieldDef->addChild(def->pasErrorANSI.toBlock(L"ErrorANSI"));

			if (m_wVersion == PropDefV2)
			{
				szFlags = flags::InterpretFlags(flagInternalType, def->dwInternalType);
				fieldDef->addChild(
					def->dwInternalType,
					L"\tInternalType = 0x%1!08X! = %2!ws!\r\n",
					def->dwInternalType.getData(),
					szFlags.c_str());
				fieldDef->addHeader(L"\tSkipBlockCount = %1!d!", def->dwSkipBlockCount);

				auto iSkip = 0;
				for (const auto& sb : def->psbSkipBlocks)
				{
					fieldDef->terminateBlock();
					auto skipBlock = std::make_shared<block>();
					skipBlock->setText(L"\tSkipBlock: %1!d!\r\n", iSkip);
					skipBlock->addChild(sb->dwSize, L"\t\tSize = 0x%1!08X!", sb->dwSize.getData());

					if (0 == iSkip)
					{
						skipBlock->terminateBlock();
						skipBlock->addChild(sb->lpbContentText.toBlock(L"\tFieldName"));
					}
					else if (!sb->lpbContent.empty())
					{
						skipBlock->terminateBlock();
						skipBlock->addHeader(L"\t\tContent = ");
						skipBlock->addChild(sb->lpbContent);
					}

					fieldDef->addChild(skipBlock);
					iSkip++;
				}
			}

			addChild(fieldDef);
			iDef++;
		}
	}
} // namespace smartview