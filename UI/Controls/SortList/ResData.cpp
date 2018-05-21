#include "stdafx.h"
#include "ResData.h"

ResData::ResData(_In_ const _SRestriction* lpOldRes)
{
	m_lpOldRes = lpOldRes;
	m_lpNewRes = nullptr;
}