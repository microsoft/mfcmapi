#include <StdAfx.h>
#include <Interpret/SmartView/GlobalObjectId.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
	// clang-format off
	static const BYTE s_rgbSPlus[] =
	{
		0x04, 0x00, 0x00, 0x00,
		0x82, 0x00, 0xE0, 0x00,
		0x74, 0xC5, 0xB7, 0x10,
		0x1A, 0x82, 0xE0, 0x08,
	};
	// clang-format on

	void GlobalObjectId::Parse()
	{
		m_Id = m_Parser.GetBYTES(16);

		const auto b1 = m_Parser.Get<BYTE>();
		const auto b2 = m_Parser.Get<BYTE>();
		m_Year.setData(static_cast<WORD>(b1 << 8 | b2));
		m_Year.setOffset(b1.getOffset());
		m_Year.setSize(b1.getSize() + b2.getSize());

		m_Month = m_Parser.Get<BYTE>();
		const auto szFlags = interpretprop::InterpretFlags(flagGlobalObjectIdMonth, m_Month);

		m_Day = m_Parser.Get<BYTE>();

		m_CreationTime = m_Parser.Get<FILETIME>();
		m_X = m_Parser.Get<LARGE_INTEGER>();
		m_dwSize = m_Parser.Get<DWORD>();
		m_lpData = m_Parser.GetBYTES(m_dwSize, _MaxBytes);
	}

	void GlobalObjectId::ParseBlocks()
	{
		setRoot(L"Global Object ID:\r\n");
		addHeader(L"Byte Array ID = ");

		auto id = m_Id.getData();
		addBlock(m_Id);

		if (equal(id.begin(), id.end(), s_rgbSPlus))
		{
			addHeader(L" = s_rgbSPlus\r\n");
		}
		else
		{
			addHeader(L" = Unknown GUID\r\n");
		}

		addBlock(m_Year, L"Year: 0x%1!04X! = %1!d!\r\n", m_Year.getData());

		addBlock(
			m_Month,
			L"Month: 0x%1!02X! = %1!d! = %2!ws!\r\n",
			m_Month.getData(),
			interpretprop::InterpretFlags(flagGlobalObjectIdMonth, m_Month).c_str());

		addBlock(m_Day, L"Day: 0x%1!02X! = %1!d!\r\n", m_Day.getData());

		std::wstring propString;
		std::wstring altPropString;
		strings::FileTimeToString(m_CreationTime, propString, altPropString);
		addBlock(
			m_CreationTime,
			L"Creation Time = 0x%1!08X!:0x%2!08X! = %3!ws!\r\n",
			m_CreationTime.getData().dwHighDateTime,
			m_CreationTime.getData().dwLowDateTime,
			propString.c_str());

		addBlock(m_X, L"X: 0x%1!08X!:0x%2!08X!\r\n", m_X.getData().HighPart, m_X.getData().LowPart);
		addBlock(m_dwSize, L"Size: 0x%1!02X! = %1!d!\r\n", m_dwSize.getData());

		if (m_lpData.size())
		{
			addHeader(L"Data = ");
			addBlock(m_lpData);
		}
	}
} // namespace smartview