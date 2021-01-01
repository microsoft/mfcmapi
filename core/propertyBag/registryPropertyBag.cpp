#include <core/stdafx.h>
#include <core/propertyBag/registryPropertyBag.h>
#include <core/mapi/mapiFunctions.h>
#include <core/utility/strings.h>
#include <core/utility/registry.h>
#include <core/utility/error.h>
#include <core/interpret/proptags.h>
#include <core/property/parseProperty.h>
#include <core/smartview/SmartView.h>

namespace propertybag
{
	registryPropertyBag::registryPropertyBag(HKEY hKey) { m_hKey = hKey; }

	registryPropertyBag ::~registryPropertyBag() {}

	propBagFlags registryPropertyBag::GetFlags() const
	{
		const auto ulFlags = propBagFlags::None;
		return ulFlags;
	}

	bool registryPropertyBag::IsEqual(const std::shared_ptr<IMAPIPropertyBag> lpPropBag) const
	{
		if (!lpPropBag) return false;
		if (GetType() != lpPropBag->GetType()) return false;

		const auto lpOther = std::dynamic_pointer_cast<registryPropertyBag>(lpPropBag);
		if (lpOther)
		{
			// Can two different hKeys be the same reg key?
			if (m_hKey != lpOther->m_hKey) return false;
			return true;
		}

		return false;
	}

	ULONG nameToPropTag(_In_ const std::wstring& name)
	{
		if (name.size() != 8) return 0;
		ULONG num{};
		if (strings::tryWstringToUlong(num, name, 16, false))
		{
			// abuse some macros to swap the order of the tag
			return PROP_TAG(PROP_ID(num), PROP_TYPE(num));
		}

		return 0;
	}

	_Check_return_ std::vector<std::shared_ptr<model::mapiRowModel>> registryPropertyBag::GetAllModels()
	{
		auto models = std::vector<std::shared_ptr<model::mapiRowModel>>{};
		auto cchMaxValueNameLen = DWORD{}; // Param in RegQueryInfoKeyW is misnamed
		auto cValues = DWORD{};
		auto hRes = WC_W32(RegQueryInfoKeyW(
			m_hKey,
			nullptr, // lpClass
			nullptr, // lpcchClass
			nullptr, // lpReserved
			nullptr, // lpcSubKeys
			nullptr, // lpcbMaxSubKeyLen
			nullptr, // lpcbMaxClassLen
			&cValues, // lpcValues
			&cchMaxValueNameLen, // lpcbMaxValueNameLen
			nullptr, // lpcbMaxValueLen
			nullptr, // lpcbSecurityDescriptor
			nullptr)); // lpftLastWriteTime

		if (cValues && cchMaxValueNameLen)
		{
			auto szBuf = std::wstring(cchMaxValueNameLen, L'\0');
			cchMaxValueNameLen++; // For null terminator
			for (DWORD dwIndex = 0; dwIndex < cValues; dwIndex++)
			{
				auto dwType = DWORD{};
				auto cchValLen = cchMaxValueNameLen;
				szBuf.clear();
				hRes = WC_W32(RegEnumValueW(
					m_hKey,
					dwIndex,
					const_cast<wchar_t*>(szBuf.c_str()),
					&cchValLen,
					nullptr, // lpReserved
					&dwType, // lpType
					nullptr, // lpData
					nullptr)); // lpcbData
				if (hRes == S_OK)
				{
					const auto valName =
						std::wstring(szBuf.c_str()); // szBuf.size() is 0, so make a copy with a proper size
					const auto ulPropTag = nameToPropTag(valName);
					auto dwVal = DWORD{};
					auto bVal = bool{};
					auto szVal = std::wstring{};
					auto binVal = std::vector<BYTE>{};
					switch (dwType)
					{
					case REG_BINARY:
						binVal = registry::ReadBinFromRegistry(m_hKey, valName);
						break;
					case REG_DWORD:
						dwVal = registry::ReadDWORDFromRegistry(m_hKey, valName);
						break;
					case REG_SZ:
						szVal = registry::ReadStringFromRegistry(m_hKey, valName);
						break;
					}
					models.push_back(regToModel(valName, ulPropTag, dwType, dwVal, szVal, binVal));
				}
			}
		}

		return models;
	}

	_Check_return_ std::shared_ptr<model::mapiRowModel> registryPropertyBag::GetOneModel(_In_ ULONG ulPropTag)
	{
		// TODO Implement
		return {};
	}

	_Check_return_ LPSPropValue registryPropertyBag::GetOneProp(ULONG ulPropTag)
	{
		// TODO: Implement
		return nullptr;
	}

	_Check_return_ HRESULT registryPropertyBag::SetProps(ULONG cValues, LPSPropValue lpPropArray)
	{
		if (!cValues || !lpPropArray) return MAPI_E_INVALID_PARAMETER;

		return E_NOTIMPL;
	}

	_Check_return_ HRESULT registryPropertyBag::SetProp(LPSPropValue lpProp) { return E_NOTIMPL; }

	_Check_return_ std::shared_ptr<model::mapiRowModel> registryPropertyBag::regToModel(
		_In_ const std::wstring& name,
		ULONG ulPropTag,
		DWORD dwType,
		DWORD dwVal,
		_In_ const std::wstring& szVal,
		_In_ const std::vector<BYTE>& binVal)
	{
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
			ret->name(name);
			ret->otherName(strings::format(L"%d", dwType)); // Just shoving in model to see it
		}

		auto bParseMAPI = false;
		auto prop = SPropValue{ulPropTag};
		if (dwType == REG_BINARY)
		{
			if (ulPropTag)
			{
				bParseMAPI = true;
				switch (PROP_TYPE(ulPropTag))
				{
				case PT_CLSID:
					if (binVal.size() == 16)
					{
						prop.Value.lpguid = reinterpret_cast<LPGUID>(const_cast<LPBYTE>(binVal.data()));
					}
					break;
				case PT_SYSTIME:
					if (binVal.size() == 8)
					{
						prop.Value.ft = *reinterpret_cast<LPFILETIME>(const_cast<LPBYTE>(binVal.data()));
					}
					break;
				case PT_I8:
					if (binVal.size() == 8)
					{
						prop.Value.li.QuadPart = static_cast<LONGLONG>(*binVal.data());
					}
					break;
				case PT_LONG:
					if (binVal.size() == 4)
					{
						prop.Value.l = static_cast<DWORD>(*binVal.data());
					}
					break;
				case PT_BOOLEAN:
					if (binVal.size() == 2)
					{
						prop.Value.b = static_cast<WORD>(*binVal.data());
					}
					break;
				case PT_BINARY:
					prop.Value.bin.cb = binVal.size();
					prop.Value.bin.lpb = const_cast<LPBYTE>(binVal.data());
					break;
				case PT_UNICODE:
					prop.Value.lpszW = reinterpret_cast<LPWSTR>(const_cast<LPBYTE>(binVal.data()));
					break;
				default:
					bParseMAPI = false;
					break;
				}
			}

			if (!bParseMAPI)
			{
				ret->ulPropTag(PROP_TAG(PT_BINARY, PROP_ID_NULL));
				ret->value(strings::BinToHexString(binVal, true));
				ret->altValue(strings::BinToTextString(binVal, true));
			}
		}
		else if (dwType == REG_DWORD)
		{
			if (ulPropTag)
			{
				bParseMAPI = true;
				prop.Value.l = dwVal;
			}
			else
			{
				ret->ulPropTag(PROP_TAG(PT_LONG, PROP_ID_NULL));
				ret->value(strings::format(L"%d", dwVal));
				ret->altValue(strings::format(L"0x%08X", dwVal));
			}
		}
		else if (dwType == REG_SZ)
		{
			ret->ulPropTag(PROP_TAG(PT_UNICODE, PROP_ID_NULL));
			ret->value(szVal);
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
		ret->namedPropName(strings::BinToHexString(binVal, true));
		ret->namedPropGuid(strings::BinToTextString(binVal, true));

		return ret;
	}
} // namespace propertybag