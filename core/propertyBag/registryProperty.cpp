#include <core/stdafx.h>
#include <core/propertyBag/registryProperty.h>
#include <core/mapi/mapiFunctions.h>
#include <core/utility/strings.h>
#include <core/utility/memory.h>
#include <core/utility/registry.h>
#include <core/utility/error.h>
#include <core/interpret/proptags.h>
#include <core/property/parseProperty.h>
#include <core/smartview/SmartView.h>
#include <core/smartview/block/binaryParser.h>
#include <core/smartview/block/blockT.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockStringA.h>

namespace propertybag
{
	registryProperty::registryProperty(HKEY hKey, _In_ const std::wstring& name, DWORD dwType)
		: m_hKey(hKey), m_name(name), m_dwType(dwType)
	{
		m_secure = false;
		// Get our prop tag and determine if we're a secure prop
		if (name.size() == 8 || name.size() == 9)
		{
			ULONG num{};
			// If we're not a simple prop tag, perhaps we have a prefix
			auto str = name;
			if (strings::stripPrefix(str, L"S") && strings::tryWstringToUlong(num, str, 16, false))
			{
				m_secure = true;
				// abuse some macros to swap the order of the tag
				m_ulPropTag = PROP_TAG(PROP_ID(num), PROP_TYPE(num));
			}
			else if (strings::tryWstringToUlong(num, name, 16, false))
			{
				// abuse some macros to swap the order of the tag
				m_ulPropTag = PROP_TAG(PROP_ID(num), PROP_TYPE(num));
			}
		}
	}

	void registryProperty::ensureSPropValue()
	{
		if (m_prop.ulPropTag != 0) return;
		m_prop = {m_ulPropTag};
		m_canParseMAPI = false;

		if (m_dwType == REG_BINARY)
		{
			m_binVal = registry::ReadBinFromRegistry(m_hKey, m_name, m_secure);
			if (m_ulPropTag)
			{
				m_canParseMAPI = true;
				switch (PROP_TYPE(m_ulPropTag))
				{
				case PT_CLSID:
					if (m_binVal.size() == 16)
					{
						m_prop.Value.lpguid = reinterpret_cast<LPGUID>(const_cast<LPBYTE>(m_binVal.data()));
					}
					break;
				case PT_SYSTIME:
					if (m_binVal.size() == 8)
					{
						m_prop.Value.ft = *reinterpret_cast<LPFILETIME>(const_cast<LPBYTE>(m_binVal.data()));
					}
					break;
				case PT_I8:
					if (m_binVal.size() == 8)
					{
						m_prop.Value.li.QuadPart = static_cast<LONGLONG>(*m_binVal.data());
					}
					break;
				case PT_LONG:
					if (m_binVal.size() == 4)
					{
						m_prop.Value.l = static_cast<DWORD>(*m_binVal.data());
					}
					break;
				case PT_BOOLEAN:
					if (m_binVal.size() == 2)
					{
						m_prop.Value.b = static_cast<WORD>(*m_binVal.data());
					}
					break;
				case PT_BINARY:
					m_prop.Value.bin.cb = m_binVal.size();
					m_prop.Value.bin.lpb = const_cast<LPBYTE>(m_binVal.data());
					break;
				case PT_UNICODE:
					m_unicodeVal = m_binVal;
					m_unicodeVal.push_back(0); // Add some null terminators just in case
					m_unicodeVal.push_back(0);
					m_prop.Value.lpszW = reinterpret_cast<LPWSTR>(const_cast<LPBYTE>(m_unicodeVal.data()));
					break;
				case PT_MV_LONG:
				{
					const auto parser = std::make_shared<smartview::binaryParser>(m_binVal);
					const auto count = parser->getSize() / sizeof(LONG);
					m_bin = std::vector<BYTE>(sizeof(LONG) * count);
					m_prop.Value.MVl.lpl = reinterpret_cast<LONG*>(m_bin.data());
					m_prop.Value.MVl.cValues = count;

					for (ULONG iMVCount = 0; iMVCount < count; iMVCount++)
					{
						m_prop.Value.MVl.lpl[iMVCount] = smartview::blockT<LONG>::parse(parser)->getData();
					}

					break;
				}
				case PT_MV_BINARY:
				{
					const auto parser = std::make_shared<smartview::binaryParser>(m_binVal);
					const auto count = smartview::blockT<ULONG>::parse(parser)->getData();
					m_bin = std::vector<BYTE>(sizeof(SBinary) * count);
					m_mvBin = std::vector<std::vector<BYTE>>(count);
					m_prop.Value.MVbin.lpbin = reinterpret_cast<SBinary*>(m_bin.data());
					m_prop.Value.MVbin.cValues = count;

					// Read lengths, offsets, and binary in a single loop
					for (ULONG iMVCount = 0; iMVCount < count; iMVCount++)
					{
						const auto length = smartview::blockT<ULONG>::parse(parser)->getData();
						const auto offset = smartview::blockT<LONG>::parse(parser)->getData();
						const auto originalOffset = parser->getOffset();
						parser->setOffset(offset);
						m_mvBin[iMVCount] = *smartview::blockBytes::parse(parser, length);
						m_prop.Value.MVbin.lpbin[iMVCount] = {length, m_mvBin[iMVCount].data()};
						parser->setOffset(originalOffset);
					}

					break;
				}
				case PT_MV_UNICODE:
				{
					const auto parser = std::make_shared<smartview::binaryParser>(m_binVal);
					const auto count = smartview::blockT<ULONG>::parse(parser)->getData();
					m_bin = std::vector<BYTE>(sizeof(LPWSTR) * count);
					m_mvW = std::vector<std::wstring>(count);
					m_prop.Value.MVszW.lppszW = reinterpret_cast<LPWSTR*>(m_bin.data());
					m_prop.Value.MVszW.cValues = count;

					// Read offsets and strings in a single loop
					for (ULONG iMVCount = 0; iMVCount < count; iMVCount++)
					{
						const auto offset = smartview::blockT<LONG>::parse(parser)->getData();
						const auto originalOffset = parser->getOffset();
						parser->setOffset(offset);
						m_mvW[iMVCount] = *smartview::blockStringW::parse(parser);
						m_prop.Value.MVszW.lppszW[iMVCount] = m_mvW[iMVCount].data();
						parser->setOffset(originalOffset);
					}

					break;
				}
				case PT_MV_STRING8:
				{
					const auto parser = std::make_shared<smartview::binaryParser>(m_binVal);
					const auto count = smartview::blockT<ULONG>::parse(parser)->getData();
					m_bin = std::vector<BYTE>(sizeof(LPSTR) * count);
					m_mvA = std::vector<std::string>(count);
					m_prop.Value.MVszA.lppszA = reinterpret_cast<LPSTR*>(m_bin.data());
					m_prop.Value.MVszA.cValues = count;

					// Read offsets and strings in a single loop
					for (ULONG iMVCount = 0; iMVCount < count; iMVCount++)
					{
						const auto offset = smartview::blockT<LONG>::parse(parser)->getData();
						const auto originalOffset = parser->getOffset();
						parser->setOffset(offset);
						m_mvA[iMVCount] = *smartview::blockStringA::parse(parser);
						m_prop.Value.MVszA.lppszA[iMVCount] = m_mvA[iMVCount].data();
						parser->setOffset(originalOffset);
					}

					break;
				}
				default:
					// Haven't found any other property types out there so this ought never be used
					m_prop.ulPropTag = PROP_TAG(PT_BINARY, PROP_ID(m_ulPropTag));
					m_prop.Value.bin.cb = m_binVal.size();
					m_prop.Value.bin.lpb = const_cast<LPBYTE>(m_binVal.data());
					m_canParseMAPI = false;
					break;
				}
			}
			else
			{
				m_prop.ulPropTag = PROP_TAG(PT_BINARY, PROP_ID_NULL);
				m_prop.Value.bin.cb = m_binVal.size();
				m_prop.Value.bin.lpb = const_cast<LPBYTE>(m_binVal.data());
			}
		}
		else if (m_dwType == REG_DWORD)
		{
			m_dwVal = registry::ReadDWORDFromRegistry(m_hKey, m_name);
			m_prop.Value.l = m_dwVal;
			if (m_ulPropTag)
			{
				m_canParseMAPI = true;
			}
			else
			{
				m_prop.ulPropTag = PROP_TAG(PT_LONG, PROP_ID_NULL);
			}
		}
		else if (m_dwType == REG_SZ)
		{
			m_szVal = registry::ReadStringFromRegistry(m_hKey, m_name);
			if (!m_prop.ulPropTag) m_prop.ulPropTag = PROP_TAG(PT_UNICODE, PROP_ID_NULL);
			switch (PROP_TYPE(m_prop.ulPropTag))
			{
			case PT_UNICODE:
				m_prop.Value.lpszW = m_szVal.data();
				break;
			case PT_STRING8:
				m_ansiVal = strings::wstringTostring(m_szVal);
				m_prop.Value.lpszA = m_ansiVal.data();
				break;
			}
		}
	}

	void registryProperty::ensureModel()
	{
		if (m_model) return;
		ensureSPropValue();
		m_model = std::make_shared<model::mapiRowModel>();
		if (m_ulPropTag != 0)
		{
			m_model->ulPropTag(m_ulPropTag);
			const auto propTagNames = proptags::PropTagToPropName(m_ulPropTag, false);
			if (!propTagNames.bestGuess.empty())
			{
				m_model->name(propTagNames.bestGuess);
			}

			if (!propTagNames.otherMatches.empty())
			{
				m_model->otherName(propTagNames.otherMatches);
			}
		}
		else
		{
			m_model->name(m_name);
		}

		if (m_secure) m_model->name(m_model->name() + L" (secure)");

		if (m_dwType == REG_BINARY)
		{
			if (!m_canParseMAPI)
			{
				if (!m_model->ulPropTag()) m_model->ulPropTag(PROP_TAG(PT_BINARY, PROP_ID_NULL));
				m_model->value(strings::BinToHexString(m_binVal, true));
				m_model->altValue(strings::BinToTextString(m_binVal, true));
			}
		}
		else if (m_dwType == REG_DWORD)
		{
			if (!m_ulPropTag)
			{
				m_model->ulPropTag(PROP_TAG(PT_LONG, PROP_ID_NULL));
				m_model->value(strings::format(L"%d", m_dwVal));
				m_model->altValue(strings::format(L"0x%08X", m_dwVal));
			}
		}
		else if (m_dwType == REG_SZ)
		{
			if (!m_model->ulPropTag()) m_model->ulPropTag(PROP_TAG(PT_UNICODE, PROP_ID_NULL));
			m_model->value(m_szVal);
		}

		if (m_canParseMAPI)
		{
			std::wstring PropString;
			std::wstring AltPropString;
			property::parseProperty(&m_prop, &PropString, &AltPropString);
			m_model->value(PropString);
			m_model->altValue(AltPropString);

			const auto szSmartView =
				smartview::parsePropertySmartView(&m_prop, nullptr, nullptr, nullptr, false, false);
			if (!szSmartView.empty()) m_model->smartView(szSmartView);
		}

		// For debugging purposes right now
		//m_model->namedPropName(strings::BinToHexString(m_binVal, true));
		//m_model->namedPropGuid(strings::BinToTextString(m_binVal, true));
	}

	void registryProperty::set(_In_opt_ LPSPropValue newValue)
	{
		if (!newValue) return;

		if (m_dwType == REG_BINARY)
		{
			if (newValue->ulPropTag)
			{
				ULONG cb = 0;
				LPBYTE lpb = nullptr;
				auto write = true;
				// TODO: Implement MV props
				switch (PROP_TYPE(newValue->ulPropTag))
				{
				case PT_CLSID:
					cb = 16;
					lpb = LPBYTE(newValue->Value.lpguid);
					break;
				case PT_SYSTIME:
					cb = 8;
					lpb = LPBYTE(&newValue->Value.ft);
					break;
				case PT_I8:
					cb = 8;
					lpb = LPBYTE(&newValue->Value.li.QuadPart);
					break;
				case PT_LONG:
					cb = 4;
					lpb = LPBYTE(&newValue->Value.l);
					break;
				case PT_BOOLEAN:
					cb = 2;
					lpb = LPBYTE(&newValue->Value.b);
					break;
				case PT_BINARY:
					cb = newValue->Value.bin.cb;
					lpb = newValue->Value.bin.lpb;
					break;
				case PT_UNICODE:
					// Include null terminator
					cb = (std::wstring(newValue->Value.lpszW).length() + 1) * sizeof(wchar_t);
					lpb = LPBYTE(newValue->Value.lpszW);
					break;
				case PT_MV_LONG:
				{
					cb = sizeof(LONG) * (newValue->Value.MVl.cValues);
					m_binVal = std::vector<BYTE>(cb); // Ok to use this since we're gonna wipe it anyway
					auto offset = 0;
					for (ULONG iMVCount = 0; iMVCount < newValue->Value.MVl.cValues; iMVCount++)
					{
						memory::assign(m_binVal, offset, newValue->Value.MVl.lpl[iMVCount]);
						offset += sizeof(LONG);
					}

					lpb = m_binVal.data();
					break;
				}
				case PT_MV_BINARY:
				{
					// Allocate the header portion, which is a fixed size based on the number of binaries to pack
					cb = sizeof(LONG) * (1 + 2 * newValue->Value.MVbin.cValues); //  count and array of size/offset
					m_binVal = std::vector<BYTE>(cb); // Ok to use this since we're gonna wipe it anyway

					// Count up our additional needs for reserve
					for (ULONG iMVCount = 0; iMVCount < newValue->Value.MVbin.cValues; iMVCount++)
					{
						cb += memory::align(newValue->Value.MVbin.lpbin[iMVCount].cb);
					}

					// Then reserve additional space for the binary data to avoid realloc
					m_binVal.reserve(cb);
					auto offset = 0;
					memory::assign(m_binVal, offset, newValue->Value.MVbin.cValues);
					offset += sizeof(ULONG);
					// As we loop through and populate the header, we append the binary to the end of our array, aligning as we go
					for (ULONG iMVCount = 0; iMVCount < newValue->Value.MVbin.cValues; iMVCount++)
					{
						auto cbBin = newValue->Value.MVbin.lpbin[iMVCount].cb;
						const auto lpbBin = newValue->Value.MVbin.lpbin[iMVCount].lpb;
						memory::assign(m_binVal, offset, cbBin);
						offset += sizeof(ULONG);
						memory::assign(m_binVal, offset, m_binVal.size());
						offset += sizeof(ULONG);
						std::copy(lpbBin, lpbBin + cbBin, std::back_inserter(m_binVal));
						while (memory::align(cbBin) > cbBin)
						{
							m_binVal.push_back(0);
							cbBin++;
						}
					}

					lpb = m_binVal.data();
					break;
				}
				case PT_MV_UNICODE:
				{
					// Allocate the header portion, which is a fixed size based on the number of strings to pack
					cb = sizeof(LONG) * (1 + newValue->Value.MVszW.cValues); //  count and array of offset
					m_binVal = std::vector<BYTE>(cb); // Ok to use this since we're gonna wipe it anyway

					// Count up our additional needs for reserve
					auto strings = std::vector<std::wstring>(newValue->Value.MVszW.cValues);
					for (ULONG iMVCount = 0; iMVCount < newValue->Value.MVszW.cValues; iMVCount++)
					{
						strings[iMVCount] = newValue->Value.MVszW.lppszW[iMVCount]; // cache for convenience
						cb += (strings[iMVCount].length() + 1) * sizeof(wchar_t);
					}

					// Then reserve additional space for the string data to avoid realloc
					m_binVal.reserve(cb);
					auto offset = 0;
					memory::assign(m_binVal, offset, newValue->Value.MVszW.cValues);
					offset += sizeof(ULONG);
					// As we loop through and populate the header, we append the null terminated strings to the end of our array
					for (ULONG iMVCount = 0; iMVCount < newValue->Value.MVszW.cValues; iMVCount++)
					{
						auto cbBin = (strings[iMVCount].length() + 1) * sizeof(wchar_t);
						const auto lpbBin = reinterpret_cast<LPBYTE>(strings[iMVCount].data());
						memory::assign(m_binVal, offset, m_binVal.size());
						offset += sizeof(ULONG);
						std::copy(lpbBin, lpbBin + cbBin, std::back_inserter(m_binVal));
					}

					lpb = m_binVal.data();
					break;
				}
				case PT_MV_STRING8:
				{
					// Allocate the header portion, which is a fixed size based on the number of strings to pack
					cb = sizeof(LONG) * (1 + newValue->Value.MVszA.cValues); //  count and array of offset
					m_binVal = std::vector<BYTE>(cb); // Ok to use this since we're gonna wipe it anyway

					// Count up our additional needs for reserve
					auto strings = std::vector<std::string>(newValue->Value.MVszA.cValues);
					for (ULONG iMVCount = 0; iMVCount < newValue->Value.MVszA.cValues; iMVCount++)
					{
						strings[iMVCount] = newValue->Value.MVszA.lppszA[iMVCount]; // cache for convenience
						cb += strings[iMVCount].length() + 1;
					}

					// Then reserve additional space for the string data to avoid realloc
					m_binVal.reserve(cb);
					auto offset = 0;
					memory::assign(m_binVal, offset, newValue->Value.MVszA.cValues);
					offset += sizeof(ULONG);
					// As we loop through and populate the header, we append the null terminated strings to the end of our array
					for (ULONG iMVCount = 0; iMVCount < newValue->Value.MVszA.cValues; iMVCount++)
					{
						auto cbBin = strings[iMVCount].length() + 1;
						const auto lpbBin = reinterpret_cast<LPBYTE>(strings[iMVCount].data());
						memory::assign(m_binVal, offset, m_binVal.size());
						offset += sizeof(ULONG);
						std::copy(lpbBin, lpbBin + cbBin, std::back_inserter(m_binVal));
					}

					lpb = m_binVal.data();
					break;
				}
				default:
					write = false;
					break;
				}

				if (write)
				{
					registry::WriteBinToRegistry(m_hKey, m_name, cb, lpb, m_secure);
				}
			}
			else
			{
				registry::WriteBinToRegistry(m_hKey, m_name, newValue->Value.bin.cb, newValue->Value.bin.lpb);
			}
		}
		else if (m_dwType == REG_DWORD)
		{
			if (PROP_TYPE(newValue->ulPropTag) == PT_LONG)
			{
				registry::WriteDWORDToRegistry(m_hKey, m_name, newValue->Value.l);
			}
		}
		else if (m_dwType == REG_SZ)
		{
			switch (PROP_TYPE(newValue->ulPropTag))
			{
			case PT_UNICODE:
				registry::WriteStringToRegistry(m_hKey, m_name, newValue->Value.lpszW);
				break;
			case PT_STRING8:
				registry::WriteStringToRegistry(m_hKey, m_name, strings::stringTowstring(newValue->Value.lpszA));
				break;
			}
		}

		// After writing, clear everything out - we'll reload as needed.
		m_model = nullptr;
		m_prop = {};
		m_binVal = {};
		ensureModel();
	}
} // namespace propertybag