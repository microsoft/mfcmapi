#pragma once
#include <Interpret/SmartView/SmartViewParser.h>

namespace smartview
{
	class SDBin : public SmartViewParser
	{
	public:
		~SDBin();

		void Init(_In_opt_ LPMAPIPROP lpMAPIProp, bool bFB);

	private:
		void Parse() override;
		void ParseBlocks() override;

		LPMAPIPROP m_lpMAPIProp{};
		bool m_bFB{};
		blockBytes m_SDbin;
	};
} // namespace smartview