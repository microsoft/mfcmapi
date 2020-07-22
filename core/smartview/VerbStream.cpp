#include <core/stdafx.h>
#include <core/smartview/VerbStream.h>
#include <core/smartview/SmartView.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	void VerbData::parse()
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

	void VerbData::parseBlocks()
	{
		addChild(VerbType, L"VerbType = 0x%1!08X!", VerbType->getData());
		addChild(DisplayNameCount, L"DisplayNameCount = 0x%1!02X!", DisplayNameCount->getData());
		addChild(DisplayName, L"DisplayName = \"%1!hs!\"", DisplayName->c_str());
		addChild(MsgClsNameCount, L"MsgClsNameCount = 0x%1!02X!", MsgClsNameCount->getData());
		addChild(MsgClsName, L"MsgClsName = \"%1!hs!\"", MsgClsName->c_str());
		addChild(Internal1StringCount, L"Internal1StringCount = 0x%1!02X!", Internal1StringCount->getData());
		addChild(Internal1String, L"Internal1String = \"%1!hs!\"", Internal1String->c_str());
		addChild(DisplayNameCountRepeat, L"DisplayNameCountRepeat = 0x%1!02X!", DisplayNameCountRepeat->getData());
		addChild(DisplayNameRepeat, L"DisplayNameRepeat = \"%1!hs!\"", DisplayNameRepeat->c_str());
		addChild(Internal2, L"Internal2 = 0x%1!08X!", Internal2->getData());
		addChild(Internal3, L"Internal3 = 0x%1!08X!", Internal3->getData());
		addChild(fUseUSHeaders, L"fUseUSHeaders = 0x%1!02X!", fUseUSHeaders->getData());
		addChild(Internal4, L"Internal4 = 0x%1!08X!", Internal4->getData());
		addChild(SendBehavior, L"SendBehavior = 0x%1!08X!", SendBehavior->getData());
		addChild(Internal5, L"Internal5 = 0x%1!08X!", Internal5->getData());
		addChild(
			ID,
			L"ID = 0x%1!08X! = %2!ws!",
			ID->getData(),
			InterpretNumberAsStringProp(*ID, PR_LAST_VERB_EXECUTED).c_str());
		addChild(Internal6, L"Internal6 = 0x%1!08X!", Internal6->getData());
	}

	void VerbExtraData::parse()
	{
		DisplayNameCount = blockT<BYTE>::parse(parser);
		DisplayName = blockStringW::parse(parser, *DisplayNameCount);
		DisplayNameCountRepeat = blockT<BYTE>::parse(parser);
		DisplayNameRepeat = blockStringW::parse(parser, *DisplayNameCountRepeat);
	}

	void VerbExtraData::parseBlocks()
	{
		addChild(DisplayNameCount, L"DisplayNameCount = 0x%1!02X!", DisplayNameCount->getData());
		addChild(DisplayName, L"DisplayName = \"%1!ws!\"", DisplayName->c_str());
		addChild(DisplayNameCountRepeat, L"DisplayNameCountRepeat = 0x%1!02X!", DisplayNameCountRepeat->getData());
		addChild(DisplayNameRepeat, L"DisplayNameRepeat = \"%1!ws!\"", DisplayNameRepeat->c_str());
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
				m_lpVerbData.emplace_back(block::parse<VerbData>(parser, false));
			}
		}

		m_Version2 = blockT<WORD>::parse(parser);

		if (*m_Count && *m_Count < _MaxEntriesSmall)
		{
			m_lpVerbExtraData.reserve(*m_Count);
			for (DWORD i = 0; i < *m_Count; i++)
			{
				m_lpVerbExtraData.emplace_back(block::parse<VerbExtraData>(parser, false));
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
			addChild(verbData, L"VerbData[%1!d!]", i++);
		}

		addChild(m_Version2, L"Version2 = 0x%1!04X!", m_Version2->getData());

		i = 0;
		for (const auto& ved : m_lpVerbExtraData)
		{
			auto dataBlock = create(L"VerbExtraData[%1!d!]", i);
			addChild(ved, L"VerbExtraData[%1!d!]", i++);
		}
	}
} // namespace smartview