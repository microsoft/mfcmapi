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

		_Check_return_ inline LPSBinary getCurrentBin()
		{
			if (m_NewBin.lpb)
			{
				return &m_OldBin;
			}
			else
			{
				return &m_NewBin;
			}
		}
		void setCurrentBin(_In_ const SBinary& bin) { m_NewBin = bin; }
		_Check_return_ SBinary detachBin(_In_opt_ const VOID* parent);

	private:
		SBinary m_OldBin{}; // not allocated - just a pointer
		SBinary m_NewBin{}; // MAPIAllocateMore from m_lpNewEntryList
	};
} // namespace sortlistdata