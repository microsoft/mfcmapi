#include "stdafx.h"
#include "BinaryData.h"

BinaryData::BinaryData(_In_ LPSBinary lpOldBin)
{
	m_OldBin = { 0 };
	if (lpOldBin)
	{
		m_OldBin.cb = lpOldBin->cb;
		m_OldBin.lpb = lpOldBin->lpb;
	}

	m_NewBin = { 0 };
}