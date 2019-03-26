#include <core/stdafx.h>
#include <core/smartview/GlobalObjectId.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>
#include <core/utility/strings.h>

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
		m_Id = blockBytes::parse(m_Parser, 16);

		const auto b1 = blockT<BYTE>::parse(m_Parser);
		const auto b2 = blockT<BYTE>::parse(m_Parser);
		m_Year->setData(static_cast<WORD>(*b1 << 8 | *b2));
		m_Year->setOffset(b1->getOffset());
		m_Year->setSize(b1->getSize() + b2->getSize());

		m_Month = blockT<BYTE>::parse(m_Parser);
		const auto szFlags = flags::InterpretFlags(flagGlobalObjectIdMonth, *m_Month);

		m_Day = blockT<BYTE>::parse(m_Parser);

		m_CreationTime = blockT<FILETIME>::parse(m_Parser);
		m_X = blockT<LARGE_INTEGER>::parse(m_Parser);
		m_dwSize = blockT<DWORD>::parse(m_Parser);
		m_lpData = blockBytes::parse(m_Parser, *m_dwSize, _MaxBytes);
	}

	void GlobalObjectId::ParseBlocks()
	{
		setRoot(L"Global Object ID:\r\n");
		addHeader(L"Byte Array ID = ");

		auto id = m_Id->getData();
		addChild(m_Id);

		if (equal(id.begin(), id.end(), s_rgbSPlus))
		{
			addHeader(L" = s_rgbSPlus\r\n");
		}
		else
		{
			addHeader(L" = Unknown GUID\r\n");
		}

		addChild(m_Year, L"Year: 0x%1!04X! = %1!d!\r\n", m_Year->getData());

		addChild(
			m_Month,
			L"Month: 0x%1!02X! = %1!d! = %2!ws!\r\n",
			m_Month->getData(),
			flags::InterpretFlags(flagGlobalObjectIdMonth, *m_Month).c_str());

		addChild(m_Day, L"Day: 0x%1!02X! = %1!d!\r\n", m_Day->getData());

		std::wstring propString;
		std::wstring altPropString;
		strings::FileTimeToString(*m_CreationTime, propString, altPropString);
		addChild(
			m_CreationTime,
			L"Creation Time = 0x%1!08X!:0x%2!08X! = %3!ws!\r\n",
			m_CreationTime->getData().dwHighDateTime,
			m_CreationTime->getData().dwLowDateTime,
			propString.c_str());

		addChild(m_X, L"X: 0x%1!08X!:0x%2!08X!\r\n", m_X->getData().HighPart, m_X->getData().LowPart);
		addChild(m_dwSize, L"Size: 0x%1!02X! = %1!d!\r\n", m_dwSize->getData());

		if (m_lpData->size())
		{
			addHeader(L"Data = ");
			addChild(m_lpData);
		}
	}
} // namespace smartview