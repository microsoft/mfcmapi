#pragma once
#include "Data.h"

class ResData :public IData
{
public:
	ResData(_In_ const _SRestriction* lpOldRes);
	const _SRestriction* m_lpOldRes; // not allocated - just a pointer
	LPSRestriction m_lpNewRes; // Owned by an alloc parent - do not free
};
