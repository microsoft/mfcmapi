#pragma once
#include <Interpret/SmartView/SmartViewParser.h>

namespace smartview
{
	class SDBin : public SmartViewParser
	{
	public:
		SDBin();
		~SDBin();

		void Init(_In_opt_ LPMAPIPROP lpMAPIProp, bool bFB);
	private:
		void Parse() override;
		_Check_return_ std::wstring ToStringInternal() override;

		LPMAPIPROP m_lpMAPIProp;
		bool m_bFB;
	};
}