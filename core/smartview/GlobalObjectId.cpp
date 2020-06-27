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

	void GlobalObjectId::parse()
	{
		m_Id = blockBytes::parse(parser, 16);

		const auto b1 = blockT<BYTE>::parse(parser);
		const auto b2 = blockT<BYTE>::parse(parser);
		m_Year =
			blockT<WORD>::create(static_cast<WORD>(*b1 << 8 | *b2), b1->getSize() + b2->getSize(), b1->getOffset());

		m_Month = blockT<BYTE>::parse(parser);
		const auto szFlags = flags::InterpretFlags(flagGlobalObjectIdMonth, *m_Month);

		m_Day = blockT<BYTE>::parse(parser);

		m_CreationTime = blockT<FILETIME>::parse(parser);
		m_X = blockT<LARGE_INTEGER>::parse(parser);
		m_dwSize = blockT<DWORD>::parse(parser);
		m_lpData = blockBytes::parse(parser, *m_dwSize, _MaxBytes);
	}

	void GlobalObjectId::parseBlocks()
	{
		setText(L"Global Object ID");
		if (m_Id->isSet())
		{
			addLabeledChild(L"Byte Array ID =", m_Id);

			if (m_Id->equal(sizeof s_rgbSPlus, s_rgbSPlus))
			{
				m_Id->addHeader(L"= s_rgbSPlus");
			}
			else
			{
				m_Id->addHeader(L"= Unknown GUID");
			}
		}

		addChild(m_Year, L"Year: 0x%1!04X! = %1!d!", m_Year->getData());

		addChild(
			m_Month,
			L"Month: 0x%1!02X! = %1!d! = %2!ws!",
			m_Month->getData(),
			flags::InterpretFlags(flagGlobalObjectIdMonth, *m_Month).c_str());

		addChild(m_Day, L"Day: 0x%1!02X! = %1!d!", m_Day->getData());

		std::wstring propString;
		std::wstring altPropString;
		strings::FileTimeToString(*m_CreationTime, propString, altPropString);
		addChild(
			m_CreationTime,
			L"Creation Time = 0x%1!08X!:0x%2!08X! = %3!ws!",
			m_CreationTime->getData().dwHighDateTime,
			m_CreationTime->getData().dwLowDateTime,
			propString.c_str());

		addChild(m_X, L"X: 0x%1!08X!:0x%2!08X!", m_X->getData().HighPart, m_X->getData().LowPart);
		addChild(m_dwSize, L"Size: 0x%1!02X! = %1!d!", m_dwSize->getData());

		if (m_lpData->size())
		{
			addLabeledChild(L"Data =", m_lpData);
		}
	}
} // namespace smartview