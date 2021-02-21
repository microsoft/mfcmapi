#include <core/stdafx.h>
#include <core/propertyBag/registryProperty.h>
#include <core/mapi/mapiFunctions.h>
#include <core/utility/strings.h>
#include <core/utility/registry.h>
#include <core/utility/error.h>
#include <core/interpret/proptags.h>
#include <core/property/parseProperty.h>
#include <core/smartview/SmartView.h>

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
			m_binVal = registry::ReadBinFromRegistry(m_hKey, m_name);
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
					if (m_secure)
					{
						// TODO - not showing right prop type in UI
						m_prop.ulPropTag = PROP_TAG(PT_BINARY, PROP_ID(m_ulPropTag));
						m_prop.Value.bin.cb = m_binVal.size();
						m_prop.Value.bin.lpb = const_cast<LPBYTE>(m_binVal.data());
					}
					else
					{
						m_unicodeVal = m_binVal;
						m_unicodeVal.push_back(0); // Add some null terminators just in case
						m_unicodeVal.push_back(0);
						m_prop.Value.lpszW = reinterpret_cast<LPWSTR>(const_cast<LPBYTE>(m_unicodeVal.data()));
					}
					break;
				default:
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
			m_model->otherName(strings::format(L"%d", m_dwType)); // Just shoving in model to see it
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
		m_model->namedPropName(strings::BinToHexString(m_binVal, true));
		m_model->namedPropGuid(strings::BinToTextString(m_binVal, true));
	}

	void registryProperty::set(_In_opt_ LPSPropValue newValue)
	{
		if (!newValue) return;

		if (m_dwType == REG_BINARY)
		{
			if (m_ulPropTag)
			{
				switch (PROP_TYPE(m_ulPropTag))
				{
				case PT_CLSID:
					registry::WriteBinToRegistry(m_hKey, m_name, 16, LPBYTE(newValue->Value.lpguid));
					break;
				case PT_SYSTIME:
					registry::WriteBinToRegistry(m_hKey, m_name, 8, LPBYTE(&newValue->Value.ft));
					break;
				case PT_I8:
					registry::WriteBinToRegistry(m_hKey, m_name, 8, LPBYTE(&newValue->Value.li.QuadPart));
					break;
				case PT_LONG:
					registry::WriteBinToRegistry(m_hKey, m_name, 4, LPBYTE(&newValue->Value.l));
					break;
				case PT_BOOLEAN:
					registry::WriteBinToRegistry(m_hKey, m_name, 2, LPBYTE(&newValue->Value.b));
					break;
				case PT_BINARY:
					registry::WriteBinToRegistry(m_hKey, m_name, newValue->Value.bin.cb, newValue->Value.bin.lpb);
					break;
				case PT_UNICODE:
					// Include null terminator
					registry::WriteBinToRegistry(
						m_hKey,
						m_name,
						(std::wstring(newValue->Value.lpszW).length() + 1) * sizeof(wchar_t),
						LPBYTE(newValue->Value.lpszW));
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

		m_model = nullptr;
		m_prop = {};
		ensureModel();
	}
} // namespace propertybag