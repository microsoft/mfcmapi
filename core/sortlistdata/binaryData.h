#pragma once
#include <core/sortlistdata/data.h>

namespace sortlistdata
{
	class sortListData;

	void InitBinary(sortListData* data, _In_opt_ LPSBinary lpOldBin);

	class binaryData : public IData
	{
	public:
		binaryData(_In_opt_ LPSBinary lpOldBin);

		SBinary m_OldBin{}; // not allocated - just a pointer
		SBinary m_NewBin{}; // MAPIAllocateMore from m_lpNewEntryList
	};
} // namespace sortlistdata