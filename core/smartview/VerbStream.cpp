#include <core/stdafx.h>
#include <core/smartview/VerbStream.h>
#include <core/smartview/SmartView.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	VerbData::VerbData(std::shared_ptr<binaryParser>& parser)
	{
		VerbType = parser->Get<DWORD>();
		DisplayNameCount = parser->Get<BYTE>();
		DisplayName.parse(parser, DisplayNameCount);
		MsgClsNameCount = parser->Get<BYTE>();
		MsgClsName.parse(parser, MsgClsNameCount);
		Internal1StringCount = parser->Get<BYTE>();
		Internal1String.parse(parser, Internal1StringCount);
		DisplayNameCountRepeat = parser->Get<BYTE>();
		DisplayNameRepeat.parse(parser, DisplayNameCountRepeat);
		Internal2 = parser->Get<DWORD>();
		Internal3 = parser->Get<BYTE>();
		fUseUSHeaders = parser->Get<DWORD>();
		Internal4 = parser->Get<DWORD>();
		SendBehavior = parser->Get<DWORD>();
		Internal5 = parser->Get<DWORD>();
		ID = parser->Get<DWORD>();
		Internal6 = parser->Get<DWORD>();
	}

	VerbExtraData::VerbExtraData(std::shared_ptr<binaryParser>& parser)
	{
		DisplayNameCount = parser->Get<BYTE>();
		DisplayName.parse(parser, DisplayNameCount);
		DisplayNameCountRepeat = parser->Get<BYTE>();
		DisplayNameRepeat.parse(parser, DisplayNameCountRepeat);
	}

	void VerbStream::Parse()
	{
		m_Version = m_Parser->Get<WORD>();
		m_Count = m_Parser->Get<DWORD>();

		if (m_Count && m_Count < _MaxEntriesSmall)
		{
			m_lpVerbData.reserve(m_Count);
			for (ULONG i = 0; i < m_Count; i++)
			{
				m_lpVerbData.emplace_back(std::make_shared<VerbData>(m_Parser));
			}
		}

		m_Version2 = m_Parser->Get<WORD>();

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
		addBlock(m_Version, L"Version = 0x%1!04X!\r\n", m_Version.getData());
		addBlock(m_Count, L"Count = 0x%1!08X!", m_Count.getData());

		auto i = 0;
		for (const auto& verbData : m_lpVerbData)
		{
			terminateBlock();
			addBlankLine();
			addHeader(L"VerbData[%1!d!]\r\n", i);
			addBlock(verbData->VerbType, L"VerbType = 0x%1!08X!\r\n", verbData->VerbType.getData());
			addBlock(
				verbData->DisplayNameCount, L"DisplayNameCount = 0x%1!02X!\r\n", verbData->DisplayNameCount.getData());
			addBlock(verbData->DisplayName, L"DisplayName = \"%1!hs!\"\r\n", verbData->DisplayName.c_str());
			addBlock(
				verbData->MsgClsNameCount, L"MsgClsNameCount = 0x%1!02X!\r\n", verbData->MsgClsNameCount.getData());
			addBlock(verbData->MsgClsName, L"MsgClsName = \"%1!hs!\"\r\n", verbData->MsgClsName.c_str());
			addBlock(
				verbData->Internal1StringCount,
				L"Internal1StringCount = 0x%1!02X!\r\n",
				verbData->Internal1StringCount.getData());
			addBlock(verbData->Internal1String, L"Internal1String = \"%1!hs!\"\r\n", verbData->Internal1String.c_str());
			addBlock(
				verbData->DisplayNameCountRepeat,
				L"DisplayNameCountRepeat = 0x%1!02X!\r\n",
				verbData->DisplayNameCountRepeat.getData());
			addBlock(
				verbData->DisplayNameRepeat,
				L"DisplayNameRepeat = \"%1!hs!\"\r\n",
				verbData->DisplayNameRepeat.c_str());
			addBlock(verbData->Internal2, L"Internal2 = 0x%1!08X!\r\n", verbData->Internal2.getData());
			addBlock(verbData->Internal3, L"Internal3 = 0x%1!08X!\r\n", verbData->Internal3.getData());
			addBlock(verbData->fUseUSHeaders, L"fUseUSHeaders = 0x%1!02X!\r\n", verbData->fUseUSHeaders.getData());
			addBlock(verbData->Internal4, L"Internal4 = 0x%1!08X!\r\n", verbData->Internal4.getData());
			addBlock(verbData->SendBehavior, L"SendBehavior = 0x%1!08X!\r\n", verbData->SendBehavior.getData());
			addBlock(verbData->Internal5, L"Internal5 = 0x%1!08X!\r\n", verbData->Internal5.getData());
			addBlock(
				verbData->ID,
				L"ID = 0x%1!08X! = %2!ws!\r\n",
				verbData->ID.getData(),
				InterpretNumberAsStringProp(verbData->ID, PR_LAST_VERB_EXECUTED).c_str());
			addBlock(verbData->Internal6, L"Internal6 = 0x%1!08X!", verbData->Internal6.getData());
			i++;
		}

		terminateBlock();
		addBlankLine();
		addBlock(m_Version2, L"Version2 = 0x%1!04X!", m_Version2.getData());

		i = 0;
		for (const auto& ved : m_lpVerbExtraData)
		{
			terminateBlock();
			addBlankLine();
			addHeader(L"VerbExtraData[%1!d!]\r\n", i);
			addBlock(ved->DisplayNameCount, L"DisplayNameCount = 0x%1!02X!\r\n", ved->DisplayNameCount.getData());
			addBlock(ved->DisplayName, L"DisplayName = \"%1!ws!\"\r\n", ved->DisplayName.c_str());
			addBlock(
				ved->DisplayNameCountRepeat,
				L"DisplayNameCountRepeat = 0x%1!02X!\r\n",
				ved->DisplayNameCountRepeat.getData());
			addBlock(ved->DisplayNameRepeat, L"DisplayNameRepeat = \"%1!ws!\"", ved->DisplayNameRepeat.c_str());
			i++;
		}
	}
} // namespace smartview