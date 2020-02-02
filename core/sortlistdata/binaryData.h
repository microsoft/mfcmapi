#pragma once
#include <core/sortlistdata/data.h>

namespace sortlistdata
{
	class sortListData;

	class binaryData : public IData
	{
	public:
		static void init(sortListData* data, _In_opt_ LPSBinary lpOldBin);

		binaryData(_In_opt_ LPSBinary lpOldBin) noexcept;

		SBinary m_OldBin{}; // not allocated - just a pointer
		SBinary m_NewBin{}; // MAPIAllocateMore from m_lpNewEntryList
	};
} // namespace sortlistdata