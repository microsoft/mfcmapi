#pragma once
#include "Data.h"

class PropListData :public IData
{
public:
	PropListData(_In_ ULONG ulPropTag);
	ULONG m_ulPropTag;
};
