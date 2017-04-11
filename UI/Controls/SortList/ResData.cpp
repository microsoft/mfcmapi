#include "stdafx.h"
#include "ResData.h"

ResData::ResData(_In_ LPSRestriction lpOldRes)
{
	m_lpOldRes = lpOldRes;
	m_lpNewRes = nullptr;
}