#include <core/stdafx.h>
#include <core/smartview/VerbStream.h>
#include <core/smartview/SmartView.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	void VerbStream::Parse()
	{
		m_Version = m_Parser->Get<WORD>();
		m_Count = m_Parser->Get<DWORD>();

		if (m_Count && m_Count < _MaxEntriesSmall)
		{
			m_lpVerbData.reserve(m_Count);
			for (ULONG i = 0; i < m_Count; i++)
			{
				VerbData verbData;
				verbData.VerbType = m_Parser->Get<DWORD>();
				verbData.DisplayNameCount = m_Parser->Get<BYTE>();
				verbData.DisplayName = m_Parser->GetStringA(verbData.DisplayNameCount);
				verbData.MsgClsNameCount = m_Parser->Get<BYTE>();
				verbData.MsgClsName = m_Parser->GetStringA(verbData.MsgClsNameCount);
				verbData.Internal1StringCount = m_Parser->Get<BYTE>();
				verbData.Internal1String = m_Parser->GetStringA(verbData.Internal1StringCount);
				verbData.DisplayNameCountRepeat = m_Parser->Get<BYTE>();
				verbData.DisplayNameRepeat = m_Parser->GetStringA(verbData.DisplayNameCountRepeat);
				verbData.Internal2 = m_Parser->Get<DWORD>();
				verbData.Internal3 = m_Parser->Get<BYTE>();
				verbData.fUseUSHeaders = m_Parser->Get<DWORD>();
				verbData.Internal4 = m_Parser->Get<DWORD>();
				verbData.SendBehavior = m_Parser->Get<DWORD>();
				verbData.Internal5 = m_Parser->Get<DWORD>();
				verbData.ID = m_Parser->Get<DWORD>();
				verbData.Internal6 = m_Parser->Get<DWORD>();
				m_lpVerbData.push_back(verbData);
			}
		}

		m_Version2 = m_Parser->Get<WORD>();

		if (m_Count && m_Count < _MaxEntriesSmall)
		{
			m_lpVerbExtraData.reserve(m_Count);
			for (ULONG i = 0; i < m_Count; i++)
			{
				VerbExtraData verbExtraData;
				verbExtraData.DisplayNameCount = m_Parser->Get<BYTE>();
				verbExtraData.DisplayName = m_Parser->GetStringW(verbExtraData.DisplayNameCount);
				verbExtraData.DisplayNameCountRepeat = m_Parser->Get<BYTE>();
				verbExtraData.DisplayNameRepeat = m_Parser->GetStringW(verbExtraData.DisplayNameCountRepeat);
				m_lpVerbExtraData.push_back(verbExtraData);
			}
		}
	}

	_Check_return_ void VerbStream::ParseBlocks()
	{
		setRoot(L"Verb Stream\r\n");
		addBlock(m_Version, L"Version = 0x%1!04X!\r\n", m_Version.getData());
		addBlock(m_Count, L"Count = 0x%1!08X!", m_Count.getData());

		for (ULONG i = 0; i < m_lpVerbData.size(); i++)
		{
			terminateBlock();
			addBlankLine();
			addHeader(L"VerbData[%1!d!]\r\n", i);
			addBlock(m_lpVerbData[i].VerbType, L"VerbType = 0x%1!08X!\r\n", m_lpVerbData[i].VerbType.getData());
			addBlock(
				m_lpVerbData[i].DisplayNameCount,
				L"DisplayNameCount = 0x%1!02X!\r\n",
				m_lpVerbData[i].DisplayNameCount.getData());
			addBlock(m_lpVerbData[i].DisplayName, L"DisplayName = \"%1!hs!\"\r\n", m_lpVerbData[i].DisplayName.c_str());
			addBlock(
				m_lpVerbData[i].MsgClsNameCount,
				L"MsgClsNameCount = 0x%1!02X!\r\n",
				m_lpVerbData[i].MsgClsNameCount.getData());
			addBlock(m_lpVerbData[i].MsgClsName, L"MsgClsName = \"%1!hs!\"\r\n", m_lpVerbData[i].MsgClsName.c_str());
			addBlock(
				m_lpVerbData[i].Internal1StringCount,
				L"Internal1StringCount = 0x%1!02X!\r\n",
				m_lpVerbData[i].Internal1StringCount.getData());
			addBlock(
				m_lpVerbData[i].Internal1String,
				L"Internal1String = \"%1!hs!\"\r\n",
				m_lpVerbData[i].Internal1String.c_str());
			addBlock(
				m_lpVerbData[i].DisplayNameCountRepeat,
				L"DisplayNameCountRepeat = 0x%1!02X!\r\n",
				m_lpVerbData[i].DisplayNameCountRepeat.getData());
			addBlock(
				m_lpVerbData[i].DisplayNameRepeat,
				L"DisplayNameRepeat = \"%1!hs!\"\r\n",
				m_lpVerbData[i].DisplayNameRepeat.c_str());
			addBlock(m_lpVerbData[i].Internal2, L"Internal2 = 0x%1!08X!\r\n", m_lpVerbData[i].Internal2.getData());
			addBlock(m_lpVerbData[i].Internal3, L"Internal3 = 0x%1!08X!\r\n", m_lpVerbData[i].Internal3.getData());
			addBlock(
				m_lpVerbData[i].fUseUSHeaders,
				L"fUseUSHeaders = 0x%1!02X!\r\n",
				m_lpVerbData[i].fUseUSHeaders.getData());
			addBlock(m_lpVerbData[i].Internal4, L"Internal4 = 0x%1!08X!\r\n", m_lpVerbData[i].Internal4.getData());
			addBlock(
				m_lpVerbData[i].SendBehavior, L"SendBehavior = 0x%1!08X!\r\n", m_lpVerbData[i].SendBehavior.getData());
			addBlock(m_lpVerbData[i].Internal5, L"Internal5 = 0x%1!08X!\r\n", m_lpVerbData[i].Internal5.getData());
			addBlock(
				m_lpVerbData[i].ID,
				L"ID = 0x%1!08X! = %2!ws!\r\n",
				m_lpVerbData[i].ID.getData(),
				InterpretNumberAsStringProp(m_lpVerbData[i].ID, PR_LAST_VERB_EXECUTED).c_str());
			addBlock(m_lpVerbData[i].Internal6, L"Internal6 = 0x%1!08X!", m_lpVerbData[i].Internal6.getData());
		}

		terminateBlock();
		addBlankLine();
		addBlock(m_Version2, L"Version2 = 0x%1!04X!", m_Version2.getData());

		for (ULONG i = 0; i < m_lpVerbExtraData.size(); i++)
		{
			terminateBlock();
			addBlankLine();
			addHeader(L"VerbExtraData[%1!d!]\r\n", i);
			addBlock(
				m_lpVerbExtraData[i].DisplayNameCount,
				L"DisplayNameCount = 0x%1!02X!\r\n",
				m_lpVerbExtraData[i].DisplayNameCount.getData());
			addBlock(
				m_lpVerbExtraData[i].DisplayName,
				L"DisplayName = \"%1!ws!\"\r\n",
				m_lpVerbExtraData[i].DisplayName.c_str());
			addBlock(
				m_lpVerbExtraData[i].DisplayNameCountRepeat,
				L"DisplayNameCountRepeat = 0x%1!02X!\r\n",
				m_lpVerbExtraData[i].DisplayNameCountRepeat.getData());
			addBlock(
				m_lpVerbExtraData[i].DisplayNameRepeat,
				L"DisplayNameRepeat = \"%1!ws!\"",
				m_lpVerbExtraData[i].DisplayNameRepeat.c_str());
		}
	}
} // namespace smartview