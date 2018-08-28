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
					m_lpRows[i].lpProps = NickNameBinToSPropValueArray(m_lpRows[i].cValues);
					m_Parser.Advance(cbBytesRead);
				}
			}
		}

		m_cbEI = m_Parser.Get<DWORD>();
		m_lpbEI = m_Parser.GetBYTES(m_cbEI, _MaxBytes);
		m_Metadata2 = m_Parser.GetBYTES(8);
	}

	_Check_return_ LPSPropValue NickNameCache::NickNameBinToSPropValueArray(DWORD dwPropCount)
	{
		if (!dwPropCount || dwPropCount > _MaxEntriesSmall) return nullptr;

		const auto pspvProperty = reinterpret_cast<LPSPropValue>(AllocateArray(dwPropCount, sizeof SPropValue));
		if (!pspvProperty) return nullptr;

		for (DWORD i = 0; i < dwPropCount; i++)
		{
			pspvProperty[i] = NickNameBinToSPropValue();
		}

		return pspvProperty;
	}

	_Check_return_ SPropValue NickNameCache::NickNameBinToSPropValue()
	{
		auto prop = SPropValue{};
		const auto PropType = m_Parser.Get<WORD>();
		const auto PropID = m_Parser.Get<WORD>();

		prop.ulPropTag = PROP_TAG(PropType, PropID);
		prop.dwAlignPad = 0;

		(void) m_Parser.Get<DWORD>(); // reserved

		switch (PropType)
		{
		case PT_I2:
		{
			prop.Value.i = m_Parser.Get<short int>();
			m_Parser.Get<short int>();
			m_Parser.Get<DWORD>();
			break;
		}
		case PT_LONG:
		{
			prop.Value.l = m_Parser.Get<DWORD>();
			m_Parser.Get<DWORD>();
			break;
		}
		case PT_ERROR:
		{
			prop.Value.err = m_Parser.Get<DWORD>();
			m_Parser.Get<DWORD>();
			break;
		}
		case PT_R4:
		{
			prop.Value.flt = m_Parser.Get<float>();
			break;
		}
		case PT_DOUBLE:
		{
			prop.Value.dbl = m_Parser.Get<double>();
			m_Parser.Get<DWORD>();
			break;
		}
		case PT_BOOLEAN:
		{
			prop.Value.b = m_Parser.Get<short int>();
			m_Parser.Get<short int>();
			m_Parser.Get<DWORD>();
			break;
		}
		case PT_SYSTIME:
		{
			prop.Value.ft = m_Parser.Get<FILETIME>();
			break;
		}
		case PT_I8:
		{
			prop.Value.li = m_Parser.Get<LARGE_INTEGER>();
			break;
		}
		case PT_STRING8:
		{
			(void) m_Parser.Get<LARGE_INTEGER>(); // union
			const auto cb = m_Parser.Get<DWORD>();
			prop.Value.lpszA = GetStringA(cb);
			break;
		}
		case PT_UNICODE:
		{
			(void) m_Parser.Get<LARGE_INTEGER>(); // union
			const auto cb = m_Parser.Get<DWORD>();
			prop.Value.lpszW = GetStringW(cb / sizeof(WCHAR));
			break;
		}
		case PT_CLSID:
		{
			(void) m_Parser.Get<LARGE_INTEGER>(); // union
			prop.Value.lpguid = reinterpret_cast<LPGUID>(GetBYTES(sizeof GUID));
			break;
		}
		case PT_BINARY:
		{
			(void) m_Parser.Get<LARGE_INTEGER>(); // union
			prop.Value.bin.cb = m_Parser.Get<DWORD>();
			// Note that we're not placing a restriction on how large a binary property we can parse. May need to revisit this.
			prop.Value.bin.lpb = GetBYTES(prop.Value.bin.cb);
			break;
		}
		case PT_MV_BINARY:
		{
			(void) m_Parser.Get<LARGE_INTEGER>(); // union
			prop.Value.MVbin.cValues = m_Parser.Get<DWORD>();
			if (prop.Value.MVbin.cValues && prop.Value.MVbin.cValues < _MaxEntriesLarge)
			{
				prop.Value.MVbin.lpbin =
					reinterpret_cast<LPSBinary>(AllocateArray(prop.Value.MVbin.cValues, sizeof SBinary));
				if (prop.Value.MVbin.lpbin)
				{
					for (ULONG j = 0; j < prop.Value.MVbin.cValues; j++)
					{
						prop.Value.MVbin.lpbin[j].cb = m_Parser.Get<DWORD>();
						// Note that we're not placing a restriction on how large a multivalued binary property we can parse. May need to revisit this.
						prop.Value.MVbin.lpbin[j].lpb = GetBYTES(prop.Value.MVbin.lpbin[j].cb);
					}
				}
			}
			break;
		}
		case PT_MV_STRING8:
		{
			(void) m_Parser.Get<LARGE_INTEGER>(); // union
			prop.Value.MVszA.cValues = m_Parser.Get<DWORD>();
			if (prop.Value.MVszA.cValues && prop.Value.MVszA.cValues < _MaxEntriesLarge)
			{
				prop.Value.MVszA.lppszA =
					reinterpret_cast<LPSTR*>(AllocateArray(prop.Value.MVszA.cValues, sizeof LPVOID));
				if (prop.Value.MVszA.lppszA)
				{
					for (ULONG j = 0; j < prop.Value.MVszA.cValues; j++)
					{
						prop.Value.MVszA.lppszA[j] = GetStringA();
					}
				}
			}
			break;
		}
		case PT_MV_UNICODE:
		{
			(void) m_Parser.Get<LARGE_INTEGER>(); // union
			prop.Value.MVszW.cValues = m_Parser.Get<DWORD>();
			if (prop.Value.MVszW.cValues && prop.Value.MVszW.cValues < _MaxEntriesLarge)
			{
				prop.Value.MVszW.lppszW =
					reinterpret_cast<LPWSTR*>(AllocateArray(prop.Value.MVszW.cValues, sizeof LPVOID));
				if (prop.Value.MVszW.lppszW)
				{
					for (ULONG j = 0; j < prop.Value.MVszW.cValues; j++)
					{
						prop.Value.MVszW.lppszW[j] = GetStringW();
					}
				}
			}
			break;
		}
		default:
			break;
		}

		return prop;
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
				if (i > 0) szNickNameCache += L"\r\n";
				szNickNameCache += strings::formatmessage(IDS_NICKNAMEROWS, i, m_lpRows[i].cValues);

				szNickNameCache += PropsToString(m_lpRows[i].cValues, m_lpRows[i].lpProps);
			}
		}

		szNickNameCache += L"\r\n";
		szNickNameCache += strings::formatmessage(IDS_NICKNAMEEXTRAINFO);
		szNickNameCache += strings::BinToHexString(m_lpbEI, true);

		szNickNameCache += strings::formatmessage(IDS_NICKNAMEFOOTER);
		szNickNameCache += strings::BinToHexString(m_Metadata2, true);

		return szNickNameCache;
	}
}