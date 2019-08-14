#include <core/stdafx.h>
#include <core/sortlistdata/binaryData.h>
#include <core/sortlistdata/sortListData.h>

namespace sortlistdata
{
	void InitBinary(sortListData* data, _In_opt_ LPSBinary lpOldBin)
	{
		if (!data) return;
		data->Init(new (std::nothrow) binaryData(lpOldBin));
	}

	binaryData::binaryData(_In_opt_ LPSBinary lpOldBin)
	{
		m_OldBin = {0};
		if (lpOldBin)
		{
			m_OldBin.cb = lpOldBin->cb;
			m_OldBin.lpb = lpOldBin->lpb;
		}

		m_NewBin = {0};
	}
} // namespace sortlistdata