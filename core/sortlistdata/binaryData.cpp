#include <core/stdafx.h>
#include <core/sortlistdata/binaryData.h>
#include <core/sortlistdata/sortListData.h>

namespace sortlistdata
{
	void binaryData::init(sortListData* data, _In_opt_ LPSBinary lpOldBin)
	{
		if (!data) return;
		data->init(std::make_shared<binaryData>(lpOldBin));
	}

	binaryData::binaryData(_In_opt_ LPSBinary lpOldBin) noexcept
	{
		if (lpOldBin)
		{
			m_OldBin = *lpOldBin;
		}
	}

	_Check_return_ SBinary binaryData::detachBin(_In_opt_ const VOID* parent)
	{
		if (m_NewBin.lpb)
		{
			return m_NewBin;
		}
		else
		{
			return mapi::CopySBinary(m_OldBin, parent);
		}
	}
} // namespace sortlistdata