#include <core/stdafx.h>
#include <core/smartview/PropertyDefinitionStream.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <core/utility/strings.h>
#include <core/mapi/cache/namedProps.h>
#include <core/interpret/proptags.h>
#include <core/smartview/block/scratchBlock.h>

namespace smartview
{
	void PackedAnsiString::parse(const std::shared_ptr<binaryParser>& parser)
	{
		cchLength = blockT<BYTE>::parse(parser);
		if (*cchLength == 0xFF)
		{
			cchExtendedLength = blockT<WORD>::parse(parser);
		}

		szCharacters =
			blockStringA::parse(parser, *cchExtendedLength ? cchExtendedLength->getData() : cchLength->getData());
	}

	void PackedUnicodeString::parse(const std::shared_ptr<binaryParser>& parser)
	{
		cchLength = blockT<BYTE>::parse(parser);
		if (*cchLength == 0xFF)
		{
			cchExtendedLength = blockT<WORD>::parse(parser);
		}

		szCharacters =
			blockStringW::parse(parser, *cchExtendedLength ? cchExtendedLength->getData() : cchLength->getData());
	}

	SkipBlock::SkipBlock(const std::shared_ptr<binaryParser>& parser, DWORD iSkip)
	{
		dwSize = blockT<DWORD>::parse(parser);
		if (iSkip == 0)
		{
			lpbContentText.parse(parser);
		}
		else
		{
			lpbContent = blockBytes::parse(parser, *dwSize, _MaxBytes);
		}
	}

	FieldDefinition::FieldDefinition(const std::shared_ptr<binaryParser>& parser, WORD version)
	{
		dwFlags = blockT<DWORD>::parse(parser);
		wVT = blockT<WORD>::parse(parser);
		dwDispid = blockT<DWORD>::parse(parser);
		wNmidNameLength = blockT<WORD>::parse(parser);
		szNmidName = blockStringW::parse(parser, *wNmidNameLength);

		pasNameANSI.parse(parser);
		pasFormulaANSI.parse(parser);
		pasValidationRuleANSI.parse(parser);
		pasValidationTextANSI.parse(parser);
		pasErrorANSI.parse(parser);

		if (version == PropDefV2)
		{
			dwInternalType = blockT<DWORD>::parse(parser);

			// Have to count how many skip blocks are here.
			// The only way to do that is to parse them. So we parse once without storing, allocate, then reparse.
			const auto stBookmark = parser->getOffset();

			DWORD skipBlockCount = 0;

			for (;;)
			{
				skipBlockCount++;
				const auto dwBlock = blockT<DWORD>::parse(parser);
				if (*dwBlock == 0) break; // we hit the last, zero byte block, or the end of the buffer
				parser->advance(*dwBlock);
			}

			parser->setOffset(stBookmark); // We're done with our first pass, restore the bookmark

			if (skipBlockCount && skipBlockCount < _MaxEntriesSmall)
			{
				psbSkipBlocks.reserve(skipBlockCount);
				for (DWORD i = 0; i < skipBlockCount; i++)
				{
					psbSkipBlocks.emplace_back(std::make_shared<SkipBlock>(parser, i));
				}
			}
		}
	}

	void PropertyDefinitionStream::parse()
	{
		m_wVersion = blockT<WORD>::parse(m_Parser);
		m_dwFieldDefinitionCount = blockT<DWORD>::parse(m_Parser);
		if (*m_dwFieldDefinitionCount)
		{
			if (*m_dwFieldDefinitionCount < _MaxEntriesLarge)
			{
				for (DWORD i = 0; i < *m_dwFieldDefinitionCount; i++)
				{
					m_pfdFieldDefinitions.emplace_back(std::make_shared<FieldDefinition>(m_Parser, *m_wVersion));
				}
			}
		}
	}

	std::shared_ptr<block> PackedAnsiString::toBlock(_In_ const std::wstring& szFieldName)
	{
		const auto& data = cchLength;

		if (*cchLength == 0xFF)
		{
			data->setText(L"\t%1!ws!: Length = 0x%2!04X!", szFieldName.c_str(), cchExtendedLength->getData());
			data->setSize(cchLength->getSize() + cchExtendedLength->getSize());
		}
		else
		{
			data->setText(L"\t%1!ws!: Length = 0x%2!04X!", szFieldName.c_str(), cchLength->getData());
		}

		if (szCharacters->length())
		{
			data->addHeader(L" Characters = ");
			data->addChild(szCharacters);
		}

		data->terminateBlock();
		return data;
	}

	std::shared_ptr<block> PackedUnicodeString::toBlock(_In_ const std::wstring& szFieldName)
	{
		const auto& data = cchLength;

		if (*cchLength == 0xFF)
		{
			data->setText(L"\t%1!ws!: Length = 0x%2!04X!", szFieldName.c_str(), cchExtendedLength->getData());
			data->setSize(cchLength->getSize() + cchExtendedLength->getSize());
		}
		else
		{
			data->setText(L"\t%1!ws!: Length = 0x%2!04X!", szFieldName.c_str(), cchLength->getData());
		}

		if (szCharacters->length())
		{
			data->addHeader(L" Characters = ");
			data->addChild(szCharacters, szCharacters->c_str());
		}

		data->terminateBlock();
		return data;
	}

	void PropertyDefinitionStream::parseBlocks()
	{
		setText(L"Property Definition Stream\r\n");
		auto szVersion = flags::InterpretFlags(flagPropDefVersion, *m_wVersion);
		addChild(m_wVersion, L"Version = 0x%1!04X! = %2!ws!\r\n", m_wVersion->getData(), szVersion.c_str());
		addChild(m_dwFieldDefinitionCount, L"FieldDefinitionCount = 0x%1!08X!", m_dwFieldDefinitionCount->getData());

		auto iDef = 0;
		for (const auto& def : m_pfdFieldDefinitions)
		{
			terminateBlock();
			auto fieldDef = create(L"Definition: %1!d!\r\n", iDef);
			addChild(fieldDef);

			auto szFlags = flags::InterpretFlags(flagPDOFlag, *def->dwFlags);
			fieldDef->addChild(
				def->dwFlags, L"\tFlags = 0x%1!08X! = %2!ws!\r\n", def->dwFlags->getData(), szFlags.c_str());
			auto szVarEnum = flags::InterpretFlags(flagVarEnum, *def->wVT);
			fieldDef->addChild(def->wVT, L"\tVT = 0x%1!04X! = %2!ws!\r\n", def->wVT->getData(), szVarEnum.c_str());
			fieldDef->addChild(def->dwDispid, L"\tDispID = 0x%1!08X!", def->dwDispid->getData());

			if (*def->dwDispid)
			{
				if (*def->dwDispid < 0x8000)
				{
					auto propTagNames = proptags::PropTagToPropName(*def->dwDispid, false);
					if (!propTagNames.bestGuess.empty())
					{
						def->dwDispid->addHeader(L" = %1!ws!", propTagNames.bestGuess.c_str());
					}

					if (!propTagNames.otherMatches.empty())
					{
						def->dwDispid->addHeader(L": (%1!ws!)", propTagNames.otherMatches.c_str());
					}
				}
				else
				{
					std::wstring szDispidName;
					auto mnid = MAPINAMEID{};
					mnid.lpguid = nullptr;
					mnid.ulKind = MNID_ID;
					mnid.Kind.lID = *def->dwDispid;
					szDispidName = strings::join(cache::NameIDToPropNames(&mnid), L", ");
					if (!szDispidName.empty())
					{
						def->dwDispid->addHeader(L" = %1!ws!", szDispidName.c_str());
					}
				}
			}

			fieldDef->terminateBlock();
			fieldDef->addChild(
				def->wNmidNameLength, L"\tNmidNameLength = 0x%1!04X!\r\n", def->wNmidNameLength->getData());
			fieldDef->addChild(def->szNmidName, L"\tNmidName = %1!ws!\r\n", def->szNmidName->c_str());

			fieldDef->addChild(def->pasNameANSI.toBlock(L"NameAnsi"));
			fieldDef->addChild(def->pasFormulaANSI.toBlock(L"FormulaANSI"));
			fieldDef->addChild(def->pasValidationRuleANSI.toBlock(L"ValidationRuleANSI"));
			fieldDef->addChild(def->pasValidationTextANSI.toBlock(L"ValidationTextANSI"));
			fieldDef->addChild(def->pasErrorANSI.toBlock(L"ErrorANSI"));

			if (*m_wVersion == PropDefV2)
			{
				szFlags = flags::InterpretFlags(flagInternalType, *def->dwInternalType);
				fieldDef->addChild(
					def->dwInternalType,
					L"\tInternalType = 0x%1!08X! = %2!ws!\r\n",
					def->dwInternalType->getData(),
					szFlags.c_str());
				fieldDef->addHeader(L"\tSkipBlockCount = %1!d!", def->psbSkipBlocks.size());

				auto iSkipBlock = 0;
				for (const auto& sb : def->psbSkipBlocks)
				{
					fieldDef->terminateBlock();
					auto skipBlock = create(L"\tSkipBlock: %1!d!\r\n", iSkipBlock);
					fieldDef->addChild(skipBlock);
					skipBlock->addChild(sb->dwSize, L"\t\tSize = 0x%1!08X!", sb->dwSize->getData());

					if (iSkipBlock == 0)
					{
						skipBlock->terminateBlock();
						skipBlock->addChild(sb->lpbContentText.toBlock(L"\tFieldName"));
					}
					else if (!sb->lpbContent->empty())
					{
						skipBlock->terminateBlock();
						skipBlock->addLabeledChild(L"\t\tContent = ", sb->lpbContent);
					}

					iSkipBlock++;
				}
			}

			iDef++;
		}
	}
} // namespace smartview