#pragma once
#include "SmartViewParser.h"

class SDBin : public SmartViewParser
{
public:
	SDBin();
	~SDBin();

	void Init(_In_opt_ LPMAPIPROP lpMAPIProp, bool bFB);
private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	LPMAPIPROP m_lpMAPIProp;
	bool m_bFB;
};