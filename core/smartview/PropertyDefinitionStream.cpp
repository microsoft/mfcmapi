#include <core/stdafx.h>
#include <core/smartview/PropertyDefinitionStream.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <core/utility/strings.h>
#include <core/mapi/cache/namedProps.h>
#include <core/interpret/proptags.h>

namespace smartview
{
	void PackedAnsiString::parse()
	{
		cchLength = blockT<BYTE>::parse(parser);
		if (*cchLength == 0xFF)
		{
			cchExtendedLength = blockT<WORD>::parse(parser);
		}

		szCharacters =
			blockStringA::parse(parser, *cchExtendedLength ? cchExtendedLength->getData() : cchLength->getData());
	}

	void PackedAnsiString::parseBlocks()
	{
		addChild(cchLength);

		if (*cchLength == 0xFF)
		{
			cchLength->setText(L"Length = 0x%1!04X!", cchExtendedLength->getData());
			cchLength->setSize(cchLength->getSize() + cchExtendedLength->getSize());
		}
		else
		{
			cchLength->setText(L"Length = 0x%1!04X!", cchLength->getData());
		}

		if (szCharacters->length())
		{
			addLabeledChild(L"Characters", szCharacters);
		}
	}

	void PackedUnicodeString::parse()
	{
		cchLength = blockT<BYTE>::parse(parser);
		if (*cchLength == 0xFF)
		{
			cchExtendedLength = blockT<WORD>::parse(parser);
		}

		szCharacters =
			blockStringW::parse(parser, *cchExtendedLength ? cchExtendedLength->getData() : cchLength->getData());
	}

	void PackedUnicodeString::parseBlocks()
	{
		addChild(cchLength);

		if (*cchLength == 0xFF)
		{
			cchLength->setText(L"Length = 0x%1!04X!", cchExtendedLength->getData());
			cchLength->setSize(cchLength->getSize() + cchExtendedLength->getSize());
		}
		else
		{
			cchLength->setText(L"Length = 0x%1!04X!", cchLength->getData());
		}

		if (szCharacters->length())
		{
			addLabeledChild(L"Characters", szCharacters);
		}
	}

	void SkipBlock::parse()
	{
		dwSize = blockT<DWORD>::parse(parser);
		if (iSkip == 0)
		{
			lpbContentText = block::parse<PackedUnicodeString>(parser, false);
		}
		else
		{
			lpbContent = blockBytes::parse(parser, *dwSize, _MaxBytes);
		}
	}

	void SkipBlock::parseBlocks()
	{
		addChild(dwSize, L"Size = 0x%1!08X!", dwSize->getData());

		if (iSkip == 0)
		{
			addChild(lpbContentText, L"FieldName");
		}
		else if (!lpbContent->empty())
		{
			addLabeledChild(L"Content", lpbContent);
		}
	}

	void FieldDefinition::parse()
	{
		dwFlags = blockT<DWORD>::parse(parser);
		wVT = blockT<WORD>::parse(parser);
		dwDispid = blockT<DWORD>::parse(parser);
		wNmidNameLength = blockT<WORD>::parse(parser);
		szNmidName = blockStringW::parse(parser, *wNmidNameLength);

		pasNameANSI = block::parse<PackedAnsiString>(parser, false);
		pasFormulaANSI = block::parse<PackedAnsiString>(parser, false);
		pasValidationRuleANSI = block::parse<PackedAnsiString>(parser, false);
		pasValidationTextANSI = block::parse<PackedAnsiString>(parser, false);
		pasErrorANSI = block::parse<PackedAnsiString>(parser, false);

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
					auto skipBlock = std::make_shared<SkipBlock>(i);
					skipBlock->block::parse(parser, false);
					psbSkipBlocks.emplace_back(skipBlock);
				}
			}
		}
	}

	void FieldDefinition::parseBlocks()
	{
		auto szFlags = flags::InterpretFlags(flagPDOFlag, *dwFlags);
		addChild(dwFlags, L"Flags = 0x%1!08X! = %2!ws!", dwFlags->getData(), szFlags.c_str());
		auto szVarEnum = flags::InterpretFlags(flagVarEnum, *wVT);
		addChild(wVT, L"VT = 0x%1!04X! = %2!ws!", wVT->getData(), szVarEnum.c_str());
		addChild(dwDispid, L"DispID = 0x%1!08X!", dwDispid->getData());

		if (*dwDispid)
		{
			if (*dwDispid < 0x8000)
			{
				auto propTagNames = proptags::PropTagToPropName(*dwDispid, false);
				if (!propTagNames.bestGuess.empty())
				{
					dwDispid->addSubHeader(L" = %1!ws!", propTagNames.bestGuess.c_str());
				}

				if (!propTagNames.otherMatches.empty())
				{
					dwDispid->addSubHeader(L": (%1!ws!)", propTagNames.otherMatches.c_str());
				}
			}
			else
			{
				std::wstring szDispidName;
				auto mnid = MAPINAMEID{};
				mnid.lpguid = nullptr;
				mnid.ulKind = MNID_ID;
				mnid.Kind.lID = *dwDispid;
				szDispidName = strings::join(cache::NameIDToPropNames(&mnid), L", ");
				if (!szDispidName.empty())
				{
					dwDispid->addSubHeader(L" = %1!ws!", szDispidName.c_str());
				}
			}
		}

		addChild(wNmidNameLength, L"NmidNameLength = 0x%1!04X!", wNmidNameLength->getData());
		addChild(szNmidName, L"NmidName = %1!ws!", szNmidName->c_str());

		addChild(pasNameANSI, L"NameAnsi");
		addChild(pasFormulaANSI, L"FormulaANSI");
		addChild(pasValidationRuleANSI, L"ValidationRuleANSI");
		addChild(pasValidationTextANSI, L"ValidationTextANSI");
		addChild(pasErrorANSI, L"ErrorANSI");

		if (version == PropDefV2)
		{
			szFlags = flags::InterpretFlags(flagInternalType, *dwInternalType);
			addChild(dwInternalType, L"InternalType = 0x%1!08X! = %2!ws!", dwInternalType->getData(), szFlags.c_str());
			addHeader(L"SkipBlockCount = %1!d!", psbSkipBlocks.size());

			auto iSkipBlock = 0;
			for (const auto& sb : psbSkipBlocks)
			{
				addChild(sb, L"SkipBlock[%1!d!]", iSkipBlock);

				iSkipBlock++;
			}
		}
	}

	void PropertyDefinitionStream::parse()
	{
		m_wVersion = blockT<WORD>::parse(parser);
		m_dwFieldDefinitionCount = blockT<DWORD>::parse(parser);
		if (*m_dwFieldDefinitionCount)
		{
			if (*m_dwFieldDefinitionCount < _MaxEntriesLarge)
			{
				for (DWORD i = 0; i < *m_dwFieldDefinitionCount; i++)
				{
					auto def = std::make_shared<FieldDefinition>(*m_wVersion);
					def->block::parse(parser, false);
					m_pfdFieldDefinitions.emplace_back(def);
				}
			}
		}
	}

	void PropertyDefinitionStream::parseBlocks()
	{
		setText(L"Property Definition Stream");
		auto szVersion = flags::InterpretFlags(flagPropDefVersion, *m_wVersion);
		addChild(m_wVersion, L"Version = 0x%1!04X! = %2!ws!", m_wVersion->getData(), szVersion.c_str());
		addChild(m_dwFieldDefinitionCount, L"FieldDefinitionCount = 0x%1!08X!", m_dwFieldDefinitionCount->getData());

		auto i = 0;
		for (const auto& def : m_pfdFieldDefinitions)
		{
			addChild(def, L"Definition: %1!d!", i++);
		}
	}
} // namespace smartview