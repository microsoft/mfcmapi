#pragma once
#include <core/sortList/data.h>

namespace controls
{
	namespace sortlistdata
	{
		class binaryData : public IData
		{
		public:
			binaryData(_In_opt_ LPSBinary lpOldBin);

			SBinary m_OldBin{}; // not allocated - just a pointer
			SBinary m_NewBin{}; // MAPIAllocateMore from m_lpNewEntryList
		};
	} // namespace sortlistdata
} // namespace controls