#include <core/stdafx.h>
#include <core/sortlistdata/binaryData.h>

namespace controls
{
	namespace sortlistdata
	{
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
} // namespace controls