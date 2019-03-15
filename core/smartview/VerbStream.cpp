#include <core/stdafx.h>
#include <core/smartview/VerbStream.h>
#include <core/smartview/SmartView.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	VerbData::VerbData(std::shared_ptr<binaryParser>& parser)
	{
		VerbType.parse<DWORD>(parser);
		DisplayNameCount.parse<BYTE>(parser);
		DisplayName.parse(parser, DisplayNameCount);
		MsgClsNameCount.parse<BYTE>(parser);
		MsgClsName.parse(parser, MsgClsNameCount);
		Internal1StringCount.parse<BYTE>(parser);
		Internal1String.parse(parser, Internal1StringCount);
		DisplayNameCountRepeat.parse<BYTE>(parser);
		DisplayNameRepeat.parse(parser, DisplayNameCountRepeat);
		Internal2.parse<DWORD>(parser);
		Internal3.parse<BYTE>(parser);
		fUseUSHeaders.parse<DWORD>(parser);
		Internal4.parse<DWORD>(parser);
		SendBehavior.parse<DWORD>(parser);
		Internal5.parse<DWORD>(parser);
		ID.parse<DWORD>(parser);
		Internal6.parse<DWORD>(parser);
	}

	VerbExtraData::VerbExtraData(std::shared_ptr<binaryParser>& parser)
	{
		DisplayNameCount.parse<BYTE>(parser);
		DisplayName = blockStringW::parse(parser, DisplayNameCount);
		DisplayNameCountRepeat.parse<BYTE>(parser);
		DisplayNameRepeat = blockStringW::parse(parser, DisplayNameCountRepeat);
	}

	void VerbStream::Parse()
	{
		m_Version.parse<WORD>(m_Parser);
		m_Count.parse<DWORD>(m_Parser);

		if (m_Count && m_Count < _MaxEntriesSmall)
		{
			m_lpVerbData.reserve(m_Count);
			for (ULONG i = 0; i < m_Count; i++)
			{
				m_lpVerbData.emplace_back(std::make_shared<VerbData>(m_Parser));
			}
		}

		m_Version2.parse<WORD>(m_Parser);

		if (m_Count && m_Count < _MaxEntriesSmall)
		{
			m_lpVerbExtraData.reserve(m_Count);
			for (ULONG i = 0; i < m_Count; i++)
			{
				m_lpVerbExtraData.emplace_back(std::make_shared<VerbExtraData>(m_Parser));
			}
		}
	}

	_Check_return_ void VerbStream::ParseBlocks()
	{
		setRoot(L"Verb Stream\r\n");
		addChild(m_Version, L"Version = 0x%1!04X!\r\n", m_Version.getData());
		addChild(m_Count, L"Count = 0x%1!08X!", m_Count.getData());

		auto i = 0;
		for (const auto& verbData : m_lpVerbData)
		{
			terminateBlock();
			addBlankLine();
			addHeader(L"VerbData[%1!d!]\r\n", i);
			addChild(verbData->VerbType, L"VerbType = 0x%1!08X!\r\n", verbData->VerbType.getData());
			addChild(
				verbData->DisplayNameCount, L"DisplayNameCount = 0x%1!02X!\r\n", verbData->DisplayNameCount.getData());
			addChild(verbData->DisplayName, L"DisplayName = \"%1!hs!\"\r\n", verbData->DisplayName.c_str());
			addChild(
				verbData->MsgClsNameCount, L"MsgClsNameCount = 0x%1!02X!\r\n", verbData->MsgClsNameCount.getData());
			addChild(verbData->MsgClsName, L"MsgClsName = \"%1!hs!\"\r\n", verbData->MsgClsName.c_str());
			addChild(
				verbData->Internal1StringCount,
				L"Internal1StringCount = 0x%1!02X!\r\n",
				verbData->Internal1StringCount.getData());
			addChild(verbData->Internal1String, L"Internal1String = \"%1!hs!\"\r\n", verbData->Internal1String.c_str());
			addChild(
				verbData->DisplayNameCountRepeat,
				L"DisplayNameCountRepeat = 0x%1!02X!\r\n",
				verbData->DisplayNameCountRepeat.getData());
			addChild(
				verbData->DisplayNameRepeat,
				L"DisplayNameRepeat = \"%1!hs!\"\r\n",
				verbData->DisplayNameRepeat.c_str());
			addChild(verbData->Internal2, L"Internal2 = 0x%1!08X!\r\n", verbData->Internal2.getData());
			addChild(verbData->Internal3, L"Internal3 = 0x%1!08X!\r\n", verbData->Internal3.getData());
			addChild(verbData->fUseUSHeaders, L"fUseUSHeaders = 0x%1!02X!\r\n", verbData->fUseUSHeaders.getData());
			addChild(verbData->Internal4, L"Internal4 = 0x%1!08X!\r\n", verbData->Internal4.getData());
			addChild(verbData->SendBehavior, L"SendBehavior = 0x%1!08X!\r\n", verbData->SendBehavior.getData());
			addChild(verbData->Internal5, L"Internal5 = 0x%1!08X!\r\n", verbData->Internal5.getData());
			addChild(
				verbData->ID,
				L"ID = 0x%1!08X! = %2!ws!\r\n",
				verbData->ID.getData(),
				InterpretNumberAsStringProp(verbData->ID, PR_LAST_VERB_EXECUTED).c_str());
			addChild(verbData->Internal6, L"Internal6 = 0x%1!08X!", verbData->Internal6.getData());
			i++;
		}

		terminateBlock();
		addBlankLine();
		addChild(m_Version2, L"Version2 = 0x%1!04X!", m_Version2.getData());

		i = 0;
		for (const auto& ved : m_lpVerbExtraData)
		{
			terminateBlock();
			addBlankLine();
			addHeader(L"VerbExtraData[%1!d!]\r\n", i);
			addChild(ved->DisplayNameCount, L"DisplayNameCount = 0x%1!02X!\r\n", ved->DisplayNameCount.getData());
			addChild(ved->DisplayName, L"DisplayName = \"%1!ws!\"\r\n", ved->DisplayName->c_str());
			addChild(
				ved->DisplayNameCountRepeat,
				L"DisplayNameCountRepeat = 0x%1!02X!\r\n",
				ved->DisplayNameCountRepeat.getData());
			addChild(ved->DisplayNameRepeat, L"DisplayNameRepeat = \"%1!ws!\"", ved->DisplayNameRepeat->c_str());
			i++;
		}
	}
} // namespace smartview