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
	ULONG nameToPropTag(_In_ const std::wstring& name, _Inout_ bool& secure)
	{
		secure = false;
		if (name.size() != 8 && name.size() != 9) return 0;
		ULONG num{};
		// If we're not a simple prop tag, perhaps we have a prefix
		auto str = name;
		if (strings::stripPrefix(str, L"S") && strings::tryWstringToUlong(num, str, 16, false))
		{
			secure = true;
			// abuse some macros to swap the order of the tag
			return PROP_TAG(PROP_ID(num), PROP_TYPE(num));
		}

		if (strings::tryWstringToUlong(num, name, 16, false))
		{
			// abuse some macros to swap the order of the tag
			return PROP_TAG(PROP_ID(num), PROP_TYPE(num));
		}

		return 0;
	}

	_Check_return_ std::shared_ptr<model::mapiRowModel> registryProperty::toModel()
	{
		auto secure{false};
		const auto ulPropTag = nameToPropTag(m_name, secure);
		auto ret = std::make_shared<model::mapiRowModel>();
		if (ulPropTag != 0)
		{
			ret->ulPropTag(ulPropTag);
			const auto propTagNames = proptags::PropTagToPropName(ulPropTag, false);
			if (!propTagNames.bestGuess.empty())
			{
				ret->name(propTagNames.bestGuess);
			}

			if (!propTagNames.otherMatches.empty())
			{
				ret->otherName(propTagNames.otherMatches);
			}
		}
		else
		{
			ret->name(m_name);
			ret->otherName(strings::format(L"%d", m_dwType)); // Just shoving in model to see it
		}

		if (secure) ret->name(ret->name() + L" (secure)");

		auto bParseMAPI = false;
		auto prop = SPropValue{ulPropTag};
		auto unicodeVal = std::vector<BYTE>{}; // in case we need to modify the bin to aid parsing
		if (m_dwType == REG_BINARY)
		{
			if (ulPropTag)
			{
				bParseMAPI = true;
				switch (PROP_TYPE(ulPropTag))
				{
				case PT_CLSID:
					if (m_binVal.size() == 16)
					{
						prop.Value.lpguid = reinterpret_cast<LPGUID>(const_cast<LPBYTE>(m_binVal.data()));
					}
					break;
				case PT_SYSTIME:
					if (m_binVal.size() == 8)
					{
						prop.Value.ft = *reinterpret_cast<LPFILETIME>(const_cast<LPBYTE>(m_binVal.data()));
					}
					break;
				case PT_I8:
					if (m_binVal.size() == 8)
					{
						prop.Value.li.QuadPart = static_cast<LONGLONG>(*m_binVal.data());
					}
					break;
				case PT_LONG:
					if (m_binVal.size() == 4)
					{
						prop.Value.l = static_cast<DWORD>(*m_binVal.data());
					}
					break;
				case PT_BOOLEAN:
					if (m_binVal.size() == 2)
					{
						prop.Value.b = static_cast<WORD>(*m_binVal.data());
					}
					break;
				case PT_BINARY:
					prop.Value.bin.cb = m_binVal.size();
					prop.Value.bin.lpb = const_cast<LPBYTE>(m_binVal.data());
					break;
				case PT_UNICODE:
					unicodeVal = m_binVal;
					unicodeVal.push_back(0); // Add some null terminators just in case
					unicodeVal.push_back(0);
					prop.Value.lpszW = reinterpret_cast<LPWSTR>(const_cast<LPBYTE>(unicodeVal.data()));
					break;
				default:
					bParseMAPI = false;
					break;
				}
			}

			if (!bParseMAPI)
			{
				if (!ret->ulPropTag()) ret->ulPropTag(PROP_TAG(PT_BINARY, PROP_ID_NULL));
				ret->value(strings::BinToHexString(m_binVal, true));
				ret->altValue(strings::BinToTextString(m_binVal, true));
			}
		}
		else if (m_dwType == REG_DWORD)
		{
			if (ulPropTag)
			{
				bParseMAPI = true;
				prop.Value.l = m_dwVal;
			}
			else
			{
				ret->ulPropTag(PROP_TAG(PT_LONG, PROP_ID_NULL));
				ret->value(strings::format(L"%d", m_dwVal));
				ret->altValue(strings::format(L"0x%08X", m_dwVal));
			}
		}
		else if (m_dwType == REG_SZ)
		{
			if (!ret->ulPropTag()) ret->ulPropTag(PROP_TAG(PT_UNICODE, PROP_ID_NULL));
			ret->value(m_szVal);
		}

		if (bParseMAPI)
		{
			std::wstring PropString;
			std::wstring AltPropString;
			property::parseProperty(&prop, &PropString, &AltPropString);
			ret->value(PropString);
			ret->altValue(AltPropString);

			const auto szSmartView = smartview::parsePropertySmartView(&prop, nullptr, nullptr, nullptr, false, false);
			if (!szSmartView.empty()) ret->smartView(szSmartView);
		}

		// For debugging purposes right now
		ret->namedPropName(strings::BinToHexString(m_binVal, true));
		ret->namedPropGuid(strings::BinToTextString(m_binVal, true));

		return ret;
	}
} // namespace propertybag