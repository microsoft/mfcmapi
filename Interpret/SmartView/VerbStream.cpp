#include <StdAfx.h>
#include <Interpret/SmartView/VerbStream.h>
#include <Interpret/SmartView/SmartView.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
	VerbStream::VerbStream() {}

	void VerbStream::Parse()
	{
		m_Version = m_Parser.GetBlock<WORD>();
		m_Count = m_Parser.GetBlock<DWORD>();

		if (m_Count.getData() && m_Count.getData() < _MaxEntriesSmall)
		{
			for (ULONG i = 0; i < m_Count.getData(); i++)
			{
				VerbData verbData;
				verbData.VerbType = m_Parser.GetBlock<DWORD>();
				verbData.DisplayNameCount = m_Parser.GetBlock<BYTE>();
				verbData.DisplayName = m_Parser.GetBlockStringA(verbData.DisplayNameCount.getData());
				verbData.MsgClsNameCount = m_Parser.GetBlock<BYTE>();
				verbData.MsgClsName = m_Parser.GetBlockStringA(verbData.MsgClsNameCount.getData());
				verbData.Internal1StringCount = m_Parser.GetBlock<BYTE>();
				verbData.Internal1String = m_Parser.GetBlockStringA(verbData.Internal1StringCount.getData());
				verbData.DisplayNameCountRepeat = m_Parser.GetBlock<BYTE>();
				verbData.DisplayNameRepeat = m_Parser.GetBlockStringA(verbData.DisplayNameCountRepeat.getData());
				verbData.Internal2 = m_Parser.GetBlock<DWORD>();
				verbData.Internal3 = m_Parser.GetBlock<BYTE>();
				verbData.fUseUSHeaders = m_Parser.GetBlock<DWORD>();
				verbData.Internal4 = m_Parser.GetBlock<DWORD>();
				verbData.SendBehavior = m_Parser.GetBlock<DWORD>();
				verbData.Internal5 = m_Parser.GetBlock<DWORD>();
				verbData.ID = m_Parser.GetBlock<DWORD>();
				verbData.Internal6 = m_Parser.GetBlock<DWORD>();
				m_lpVerbData.push_back(verbData);
			}
		}

		m_Version2 = m_Parser.GetBlock<WORD>();

		if (m_Count.getData() && m_Count.getData() < _MaxEntriesSmall)
		{
			for (ULONG i = 0; i < m_Count.getData(); i++)
			{
				VerbExtraData verbExtraData;
				verbExtraData.DisplayNameCount = m_Parser.GetBlock<BYTE>();
				verbExtraData.DisplayName = m_Parser.GetBlockStringW(verbExtraData.DisplayNameCount.getData());
				verbExtraData.DisplayNameCountRepeat = m_Parser.GetBlock<BYTE>();
				verbExtraData.DisplayNameRepeat =
					m_Parser.GetBlockStringW(verbExtraData.DisplayNameCountRepeat.getData());
				m_lpVerbExtraData.push_back(verbExtraData);
			}
		}
	}

	_Check_return_ std::wstring VerbStream::ToStringInternal()
	{
		std::vector<std::wstring> verbStream;
		verbStream.push_back(strings::formatmessage(IDS_VERBHEADER, m_Version.getData(), m_Count.getData()));

		for (ULONG i = 0; i < m_lpVerbData.size(); i++)
		{
			verbStream.push_back(strings::formatmessage(
				IDS_VERBDATA,
				i,
				m_lpVerbData[i].VerbType.getData(),
				m_lpVerbData[i].DisplayNameCount.getData(),
				m_lpVerbData[i].DisplayName.getData().c_str(),
				m_lpVerbData[i].MsgClsNameCount.getData(),
				m_lpVerbData[i].MsgClsName.getData().c_str(),
				m_lpVerbData[i].Internal1StringCount.getData(),
				m_lpVerbData[i].Internal1String.getData().c_str(),
				m_lpVerbData[i].DisplayNameCountRepeat.getData(),
				m_lpVerbData[i].DisplayNameRepeat.getData().c_str(),
				m_lpVerbData[i].Internal2.getData(),
				m_lpVerbData[i].Internal3.getData(),
				m_lpVerbData[i].fUseUSHeaders.getData(),
				m_lpVerbData[i].Internal4.getData(),
				m_lpVerbData[i].SendBehavior.getData(),
				m_lpVerbData[i].Internal5.getData(),
				m_lpVerbData[i].ID.getData(),
				smartview::InterpretNumberAsStringProp(m_lpVerbData[i].ID.getData(), PR_LAST_VERB_EXECUTED).c_str(),
				m_lpVerbData[i].Internal6.getData()));
		}

		verbStream.push_back(strings::formatmessage(IDS_VERBVERSION2, m_Version2.getData()));

		for (ULONG i = 0; i < m_lpVerbExtraData.size(); i++)
		{
			verbStream.push_back(strings::formatmessage(
				IDS_VERBEXTRADATA,
				i,
				m_lpVerbExtraData[i].DisplayNameCount.getData(),
				m_lpVerbExtraData[i].DisplayName.getData().c_str(),
				m_lpVerbExtraData[i].DisplayNameCountRepeat.getData(),
				m_lpVerbExtraData[i].DisplayNameRepeat.getData().c_str()));
		}

		return strings::join(verbStream, L"\r\n\r\n");
	}
}