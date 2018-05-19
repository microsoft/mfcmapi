#pragma once
#include "SmartViewParser.h"

namespace smartview
{
	class RecipientRowStream : public SmartViewParser
	{
	public:
		RecipientRowStream();

	private:
		void Parse() override;
		_Check_return_ std::wstring ToStringInternal() override;

		DWORD m_cVersion;
		DWORD m_cRowCount;
		LPADRENTRY m_lpAdrEntry;
	};
}