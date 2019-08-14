#include <core/stdafx.h>
#include <core/sortlistdata/mvPropData.h>
#include <core/sortlistdata/sortListData.h>

namespace sortlistdata
{
	void mvPropData::init(sortListData* data, _In_ const _SPropValue* lpProp, ULONG iProp)
	{
		if (!data) return;
		data->Init(new (std::nothrow) mvPropData(lpProp, iProp));
	}

	void mvPropData::init(sortListData* data, _In_opt_ const _SPropValue* lpProp)
	{
		if (!data) return;
		data->Init(new (std::nothrow) mvPropData(lpProp));
	}

	mvPropData::mvPropData(_In_opt_ const _SPropValue* lpProp, ULONG iProp)
	{
		if (!lpProp) return;

		switch (PROP_TYPE(lpProp->ulPropTag))
		{
		case PT_MV_I2:
			m_val.i = lpProp->Value.MVi.lpi[iProp];
			break;
		case PT_MV_LONG:
			m_val.l = lpProp->Value.MVl.lpl[iProp];
			break;
		case PT_MV_DOUBLE:
			m_val.dbl = lpProp->Value.MVdbl.lpdbl[iProp];
			break;
		case PT_MV_CURRENCY:
			m_val.cur = lpProp->Value.MVcur.lpcur[iProp];
			break;
		case PT_MV_APPTIME:
			m_val.at = lpProp->Value.MVat.lpat[iProp];
			break;
		case PT_MV_SYSTIME:
			m_val.ft = lpProp->Value.MVft.lpft[iProp];
			break;
		case PT_MV_I8:
			m_val.li = lpProp->Value.MVli.lpli[iProp];
			break;
		case PT_MV_R4:
			m_val.flt = lpProp->Value.MVflt.lpflt[iProp];
			break;
		case PT_MV_STRING8:
			m_lpszA = lpProp->Value.MVszA.lppszA[iProp];
			m_val.lpszA = const_cast<LPSTR>(m_lpszA.c_str());
			break;
		case PT_MV_UNICODE:
			m_lpszW = lpProp->Value.MVszW.lppszW[iProp];
			m_val.lpszW = const_cast<LPWSTR>(m_lpszW.c_str());
			break;
		case PT_MV_BINARY:
			m_lpBin.assign(
				lpProp->Value.MVbin.lpbin[iProp].lpb,
				lpProp->Value.MVbin.lpbin[iProp].lpb + lpProp->Value.MVbin.lpbin[iProp].cb);
			m_val.bin.cb = static_cast<ULONG>(m_lpBin.size());
			m_val.bin.lpb = m_lpBin.data();
			break;
		case PT_MV_CLSID:
			m_lpBin.assign(
				reinterpret_cast<LPBYTE>(&lpProp->Value.MVguid.lpguid[iProp]),
				reinterpret_cast<LPBYTE>(&lpProp->Value.MVguid.lpguid[iProp]) + sizeof(GUID));
			m_val.lpguid = reinterpret_cast<LPGUID>(m_lpBin.data());
			break;
		default:
			break;
		}
	}

	mvPropData::mvPropData(_In_opt_ const _SPropValue* lpProp)
	{
		if (!lpProp) return;

		// This handles most cases by default - cases needing a buffer copied are handled below
		m_val = lpProp->Value;

		switch (PROP_TYPE(lpProp->ulPropTag))
		{
		case PT_STRING8:
			m_lpszA = lpProp->Value.lpszA;
			m_val.lpszA = const_cast<LPSTR>(m_lpszA.c_str());
			break;
		case PT_UNICODE:
			m_lpszW = lpProp->Value.lpszW;
			m_val.lpszW = const_cast<LPWSTR>(m_lpszW.c_str());
			break;
		case PT_BINARY:
			m_lpBin.assign(m_val.bin.lpb, m_val.bin.lpb + lpProp->Value.bin.cb);
			m_val.bin.cb = static_cast<ULONG>(m_lpBin.size());
			m_val.bin.lpb = m_lpBin.data();
			break;
		case PT_CLSID:
			m_lpBin.assign(
				reinterpret_cast<LPBYTE>(lpProp->Value.lpguid),
				reinterpret_cast<LPBYTE>(lpProp->Value.lpguid) + sizeof(GUID));
			m_val.lpguid = reinterpret_cast<LPGUID>(m_lpBin.data());
			break;
		default:
			break;
		}
	}
} // namespace sortlistdata