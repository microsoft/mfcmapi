#pragma once
#include <Interpret/SmartView/SmartViewParser.h>

namespace smartview
{
	class GlobalObjectId : public SmartViewParser
	{
	public:
		GlobalObjectId();

	private:
		void Parse() override;

		blockBytes m_Id; // 16 bytes
		blockT<WORD> m_Year;
		blockT<BYTE> m_Month;
		blockT<BYTE> m_Day;
		blockT<FILETIME> m_CreationTime;
		blockT<LARGE_INTEGER> m_X;
		blockT<DWORD> m_dwSize;
		blockBytes m_lpData;
	};
}