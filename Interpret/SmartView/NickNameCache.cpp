#include <StdAfx.h>
#include <Interpret/SmartView/NickNameCache.h>
#include <Interpret/String.h>
#include <Interpret/SmartView/PropertyStruct.h>

namespace smartview
{
	NickNameCache::NickNameCache()
	{
		m_ulMajorVersion = 0;
		m_ulMinorVersion = 0;
		m_cRowCount = 0;
		m_lpRows = nullptr;
		m_cbEI = 0;
	}

	void NickNameCache::Parse()
	{
		m_Metadata1 = m_Parser.GetBYTES(4);
		m_ulMajorVersion = m_Parser.Get<DWORD>();
		m_ulMinorVersion = m_Parser.Get<DWORD>();
		m_cRowCount = m_Parser.Get<DWORD>();

		if (m_cRowCount && m_cRowCount < _MaxEntriesEnormous)
			m_lpRows = reinterpret_cast<LPSRow>(AllocateArray(m_cRowCount, sizeof SRow));

		if (m_lpRows)
		{
			for (DWORD i = 0; i < m_cRowCount; i++)
			{
				m_lpRows[i].cValues = m_Parser.Get<DWORD>();

				if (m_lpRows[i].cValues && m_lpRows[i].cValues < _MaxEntriesSmall)
				{
					const size_t cbBytesRead = 0;
					m_lpRows[i].lpProps = NickNameBinToSPropValue(m_lpRows[i].cValues);
					m_Parser.Advance(cbBytesRead);
				}
			}
		}

		m_cbEI = m_Parser.Get<DWORD>();
		m_lpbEI = m_Parser.GetBYTES(m_cbEI, _MaxBytes);
		m_Metadata2 = m_Parser.GetBYTES(8);
	}

	_Check_return_ LPSPropValue NickNameCache::NickNameBinToSPropValue(DWORD dwPropCount)
	{
		if (!dwPropCount || dwPropCount > _MaxEntriesSmall) return nullptr;

		const auto pspvProperty = reinterpret_cast<LPSPropValue>(AllocateArray(dwPropCount, sizeof SPropValue));
		if (!pspvProperty) return nullptr;

		for (DWORD i = 0; i < dwPropCount; i++)
		{
			const auto PropType = m_Parser.Get<WORD>();
			const auto PropID = m_Parser.Get<WORD>();

			pspvProperty[i].ulPropTag = PROP_TAG(PropType, PropID);
			pspvProperty[i].dwAlignPad = 0;

			auto dwTemp = m_Parser.Get<DWORD>(); // reserved
			const auto liTemp = m_Parser.Get<LARGE_INTEGER>(); // union

			switch (PropType)
			{
			case PT_I2:
				pspvProperty[i].Value.i = static_cast<short int>(liTemp.LowPart);
				break;
			case PT_LONG:
				pspvProperty[i].Value.l = liTemp.LowPart;
				break;
			case PT_ERROR:
				pspvProperty[i].Value.err = liTemp.LowPart;
				break;
			case PT_R4:
				pspvProperty[i].Value.flt = static_cast<float>(liTemp.QuadPart);
				break;
			case PT_DOUBLE:
				pspvProperty[i].Value.dbl = liTemp.LowPart;
				break;
			case PT_BOOLEAN:
				pspvProperty[i].Value.b = liTemp.LowPart ? true : false;
				break;
			case PT_SYSTIME:
				pspvProperty[i].Value.ft.dwHighDateTime = liTemp.HighPart;
				pspvProperty[i].Value.ft.dwLowDateTime = liTemp.LowPart;
				break;
			case PT_I8:
				pspvProperty[i].Value.li = liTemp;
				break;
			case PT_STRING8:
				dwTemp = m_Parser.Get<DWORD>();
				pspvProperty[i].Value.lpszA = GetStringA(dwTemp);
				break;
			case PT_UNICODE:
				dwTemp = m_Parser.Get<DWORD>();
				pspvProperty[i].Value.lpszW = GetStringW(dwTemp / sizeof(WCHAR));
				break;
			case PT_CLSID:
				pspvProperty[i].Value.lpguid = reinterpret_cast<LPGUID>(GetBYTES(sizeof GUID));
				break;
			case PT_BINARY:
				dwTemp = m_Parser.Get<DWORD>();
				pspvProperty[i].Value.bin.cb = dwTemp;
				// Note that we're not placing a restriction on how large a binary property we can parse. May need to revisit this.
				pspvProperty[i].Value.bin.lpb = GetBYTES(pspvProperty[i].Value.bin.cb);
				break;
			case PT_MV_BINARY:
				dwTemp = m_Parser.Get<DWORD>();
				pspvProperty[i].Value.MVbin.cValues = dwTemp;
				if (pspvProperty[i].Value.MVbin.cValues && pspvProperty[i].Value.MVbin.cValues < _MaxEntriesLarge)
				{
					pspvProperty[i].Value.MVbin.lpbin =
						reinterpret_cast<LPSBinary>(AllocateArray(dwTemp, sizeof SBinary));
					if (pspvProperty[i].Value.MVbin.lpbin)
					{
						for (ULONG j = 0; j < pspvProperty[i].Value.MVbin.cValues; j++)
						{
							dwTemp = m_Parser.Get<DWORD>();
							pspvProperty[i].Value.MVbin.lpbin[j].cb = dwTemp;
							// Note that we're not placing a restriction on how large a multivalued binary property we can parse. May need to revisit this.
							pspvProperty[i].Value.MVbin.lpbin[j].lpb =
								GetBYTES(pspvProperty[i].Value.MVbin.lpbin[j].cb);
						}
					}
				}
				break;
			case PT_MV_STRING8:
				dwTemp = m_Parser.Get<DWORD>();
				pspvProperty[i].Value.MVszA.cValues = dwTemp;
				if (pspvProperty[i].Value.MVszA.cValues && pspvProperty[i].Value.MVszA.cValues < _MaxEntriesLarge)
				{
					pspvProperty[i].Value.MVszA.lppszA = reinterpret_cast<LPSTR*>(AllocateArray(dwTemp, sizeof LPVOID));
					if (pspvProperty[i].Value.MVszA.lppszA)
					{
						for (ULONG j = 0; j < pspvProperty[i].Value.MVszA.cValues; j++)
						{
							pspvProperty[i].Value.MVszA.lppszA[j] = GetStringA();
						}
					}
				}
				break;
			case PT_MV_UNICODE:
				dwTemp = m_Parser.Get<DWORD>();
				pspvProperty[i].Value.MVszW.cValues = dwTemp;
				if (pspvProperty[i].Value.MVszW.cValues && pspvProperty[i].Value.MVszW.cValues < _MaxEntriesLarge)
				{
					pspvProperty[i].Value.MVszW.lppszW =
						reinterpret_cast<LPWSTR*>(AllocateArray(dwTemp, sizeof LPVOID));
					if (pspvProperty[i].Value.MVszW.lppszW)
					{
						for (ULONG j = 0; j < pspvProperty[i].Value.MVszW.cValues; j++)
						{
							pspvProperty[i].Value.MVszW.lppszW[j] = GetStringW();
						}
					}
				}
				break;
			default:
				break;
			}
		}

		return pspvProperty;
	}

	_Check_return_ std::wstring NickNameCache::ToStringInternal()
	{
		auto szNickNameCache = strings::formatmessage(IDS_NICKNAMEHEADER);
		szNickNameCache += strings::BinToHexString(m_Metadata1, true);

		szNickNameCache +=
			strings::formatmessage(IDS_NICKNAMEROWCOUNT, m_ulMajorVersion, m_ulMinorVersion, m_cRowCount);

		if (m_cRowCount && m_lpRows)
		{
			for (DWORD i = 0; i < m_cRowCount; i++)
			{
				szNickNameCache += strings::formatmessage(IDS_NICKNAMEROWS, i, m_lpRows[i].cValues);

				szNickNameCache += PropsToString(m_lpRows[i].cValues, m_lpRows[i].lpProps);
			}
		}

		szNickNameCache += strings::formatmessage(IDS_NICKNAMEEXTRAINFO);
		szNickNameCache += strings::BinToHexString(m_lpbEI, true);

		szNickNameCache += strings::formatmessage(IDS_NICKNAMEFOOTER);
		szNickNameCache += strings::BinToHexString(m_Metadata2, true);

		return szNickNameCache;
	}
}