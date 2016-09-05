#pragma once

class ResData
{
public:
	ResData(_In_ LPSRestriction lpOldRes);
	LPSRestriction m_lpOldRes; // not allocated - just a pointer
	LPSRestriction m_lpNewRes; // Owned by an alloc parent - do not free
};
