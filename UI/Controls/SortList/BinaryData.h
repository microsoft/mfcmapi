#pragma once
#include <UI/Controls/SortList/Data.h>

namespace controls
{
	namespace sortlistdata
	{
		class BinaryData : public IData
		{
		public:
			BinaryData(_In_ LPSBinary lpOldBin);

			SBinary m_OldBin{}; // not allocated - just a pointer
			SBinary m_NewBin{}; // MAPIAllocateMore from m_lpNewEntryList
		};
	} // namespace sortlistdata
} // namespace controls