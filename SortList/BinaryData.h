#pragma once

class BinaryData
{
public:
	BinaryData(_In_ LPSBinary lpOldBin);
	~BinaryData();

	SBinary m_OldBin; // not allocated - just a pointer
	SBinary m_NewBin; // MAPIAllocateMore from m_lpNewEntryList
};
