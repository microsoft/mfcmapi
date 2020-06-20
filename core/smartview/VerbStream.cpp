#include <core/stdafx.h>
#include <core/smartview/VerbStream.h>
#include <core/smartview/SmartView.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	VerbData::VerbData(const std::shared_ptr<binaryParser>& parser)
	{
		VerbType = blockT<DWORD>::parse(parser);
		DisplayNameCount = blockT<BYTE>::parse(parser);
		DisplayName = blockStringA::parse(parser, *DisplayNameCount);
		MsgClsNameCount = blockT<BYTE>::parse(parser);
		MsgClsName = blockStringA::parse(parser, *MsgClsNameCount);
		Internal1StringCount = blockT<BYTE>::parse(parser);
		Internal1String = blockStringA::parse(parser, *Internal1StringCount);
		DisplayNameCountRepeat = blockT<BYTE>::parse(parser);
		DisplayNameRepeat = blockStringA::parse(parser, *DisplayNameCountRepeat);
		Internal2 = blockT<DWORD>::parse(parser);
		Internal3 = blockT<BYTE>::parse(parser);
		fUseUSHeaders = blockT<DWORD>::parse(parser);
		Internal4 = blockT<DWORD>::parse(parser);
		SendBehavior = blockT<DWORD>::parse(parser);
		Internal5 = blockT<DWORD>::parse(parser);
		ID = blockT<DWORD>::parse(parser);
		Internal6 = blockT<DWORD>::parse(parser);
	}

	VerbExtraData::VerbExtraData(const std::shared_ptr<binaryParser>& parser)
	{
		DisplayNameCount = blockT<BYTE>::parse(parser);
		DisplayName = blockStringW::parse(parser, *DisplayNameCount);
		DisplayNameCountRepeat = blockT<BYTE>::parse(parser);
		DisplayNameRepeat = blockStringW::parse(parser, *DisplayNameCountRepeat);
	}

	void VerbStream::parse()
	{
		m_Version = blockT<WORD>::parse(parser);
		m_Count = blockT<DWORD>::parse(parser);

		if (*m_Count && *m_Count < _MaxEntriesSmall)
		{
			m_lpVerbData.reserve(*m_Count);
			for (DWORD i = 0; i < *m_Count; i++)
			{
				m_lpVerbData.emplace_back(std::make_shared<VerbData>(parser));
			}
		}

		m_Version2 = blockT<WORD>::parse(parser);

		if (*m_Count && *m_Count < _MaxEntriesSmall)
		{
			m_lpVerbExtraData.reserve(*m_Count);
			for (DWORD i = 0; i < *m_Count; i++)
			{
				m_lpVerbExtraData.emplace_back(std::make_shared<VerbExtraData>(parser));
			}
		}
	}

	void VerbStream::parseBlocks()
	{
		setText(L"Verb Stream");
		addChild(m_Version, L"Version = 0x%1!04X!", m_Version->getData());
		addChild(m_Count, L"Count = 0x%1!08X!", m_Count->getData());

		auto i = 0;
		for (const auto& verbData : m_lpVerbData)
		{
			terminateBlock();
			addBlankLine();
			auto dataBlock = create(L"VerbData[%1!d!]", i);
			addChild(dataBlock);
			dataBlock->addChild(verbData->VerbType, L"VerbType = 0x%1!08X!", verbData->VerbType->getData());
			dataBlock->addChild(
				verbData->DisplayNameCount, L"DisplayNameCount = 0x%1!02X!", verbData->DisplayNameCount->getData());
			dataBlock->addChild(verbData->DisplayName, L"DisplayName = \"%1!hs!\"", verbData->DisplayName->c_str());
			dataBlock->addChild(
				verbData->MsgClsNameCount, L"MsgClsNameCount = 0x%1!02X!", verbData->MsgClsNameCount->getData());
			dataBlock->addChild(verbData->MsgClsName, L"MsgClsName = \"%1!hs!\"", verbData->MsgClsName->c_str());
			dataBlock->addChild(
				verbData->Internal1StringCount,
				L"Internal1StringCount = 0x%1!02X!",
				verbData->Internal1StringCount->getData());
			dataBlock->addChild(
				verbData->Internal1String, L"Internal1String = \"%1!hs!\"", verbData->Internal1String->c_str());
			dataBlock->addChild(
				verbData->DisplayNameCountRepeat,
				L"DisplayNameCountRepeat = 0x%1!02X!",
				verbData->DisplayNameCountRepeat->getData());
			dataBlock->addChild(
				verbData->DisplayNameRepeat, L"DisplayNameRepeat = \"%1!hs!\"", verbData->DisplayNameRepeat->c_str());
			dataBlock->addChild(verbData->Internal2, L"Internal2 = 0x%1!08X!", verbData->Internal2->getData());
			dataBlock->addChild(verbData->Internal3, L"Internal3 = 0x%1!08X!", verbData->Internal3->getData());
			dataBlock->addChild(
				verbData->fUseUSHeaders, L"fUseUSHeaders = 0x%1!02X!", verbData->fUseUSHeaders->getData());
			dataBlock->addChild(verbData->Internal4, L"Internal4 = 0x%1!08X!", verbData->Internal4->getData());
			dataBlock->addChild(verbData->SendBehavior, L"SendBehavior = 0x%1!08X!", verbData->SendBehavior->getData());
			dataBlock->addChild(verbData->Internal5, L"Internal5 = 0x%1!08X!", verbData->Internal5->getData());
			dataBlock->addChild(
				verbData->ID,
				L"ID = 0x%1!08X! = %2!ws!",
				verbData->ID->getData(),
				InterpretNumberAsStringProp(*verbData->ID, PR_LAST_VERB_EXECUTED).c_str());
			dataBlock->addChild(verbData->Internal6, L"Internal6 = 0x%1!08X!", verbData->Internal6->getData());

			i++;
		}

		terminateBlock();
		addBlankLine();
		addChild(m_Version2, L"Version2 = 0x%1!04X!", m_Version2->getData());

		i = 0;
		for (const auto& ved : m_lpVerbExtraData)
		{
			auto dataBlock = create(L"VerbExtraData[%1!d!]", i);
			addChild(dataBlock);
			dataBlock->addChild(ved->DisplayNameCount, L"DisplayNameCount = 0x%1!02X!", ved->DisplayNameCount->getData());
			dataBlock->addChild(ved->DisplayName, L"DisplayName = \"%1!ws!\"", ved->DisplayName->c_str());
			dataBlock->addChild(
				ved->DisplayNameCountRepeat,
				L"DisplayNameCountRepeat = 0x%1!02X!",
				ved->DisplayNameCountRepeat->getData());
			dataBlock->addChild(
				ved->DisplayNameRepeat, L"DisplayNameRepeat = \"%1!ws!\"", ved->DisplayNameRepeat->c_str());

			i++;
		}
	}
} // namespace smartview