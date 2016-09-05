#pragma once
class CommentData
{
public:
	CommentData(_In_ LPSPropValue lpOldProp);
	LPSPropValue m_lpOldProp; // not allocated - just a pointer
	LPSPropValue m_lpNewProp; // Owned by an alloc parent - do not free
};
