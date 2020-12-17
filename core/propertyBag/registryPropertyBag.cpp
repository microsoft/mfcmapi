#include <core/stdafx.h>
#include <core/propertyBag/registryPropertyBag.h>
#include <core/mapi/mapiFunctions.h>
#include <core/utility/strings.h>
#include <core/utility/registry.h>
#include <core/utility/error.h>
#include <core/interpret/proptags.h>

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
					auto valName = std::wstring(szBuf.c_str()); // szBuf.size() is 0, so make a copy with a proper size
					auto dwVal = DWORD{};
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
					models.push_back(regToModel(valName, dwType, dwVal, szVal, binVal));
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

	ULONG nameToULONG(_In_ const std::wstring& name)
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

	// TODO: Identify prop tags in the value name
	// Use type from tag to determine how to read data
	// Also use lpType
	// Convert data read to a MAPI prop Value
	// Figure out way to deal with S props
	// Figure out way to deal with named props
	_Check_return_ std::shared_ptr<model::mapiRowModel> registryPropertyBag::regToModel(
		_In_ const std::wstring& name,
		DWORD dwType,
		DWORD dwVal,
		_In_ const std::wstring& szVal,
		_In_ const std::vector<BYTE>& binVal)
	{
		auto ret = std::make_shared<model::mapiRowModel>();
		const auto ulPropTag = nameToULONG(name);
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

		switch (dwType)
		{
		case REG_BINARY:
			ret->value(strings::BinToHexString(binVal, true));
			ret->altValue(strings::BinToTextString(binVal, true));
			break;
		case REG_DWORD:
			ret->value(strings::format(L"0x%08X", dwVal));
			break;
		case REG_SZ:
			ret->value(szVal);
			break;
		}

		//const auto propTagNames = proptags::PropTagToPropName(ulPropTag, bIsAB);
		//const auto namePropNames = cache::NameIDToStrings(ulPropTag, lpProp, nullptr, sig, bIsAB);
		//if (!propTagNames.bestGuess.empty())
		//{
		//	ret->name(propTagNames.bestGuess);
		//}
		//else if (!namePropNames.bestPidLid.empty())
		//{
		//	ret->name(namePropNames.bestPidLid);
		//}
		//else if (!namePropNames.name.empty())
		//{
		//	ret->name(namePropNames.name);
		//}

		//std::wstring PropString;
		//std::wstring AltPropString;
		//property::parseProperty(lpPropVal, &PropString, &AltPropString);
		//ret->value(PropString);
		//ret->altValue(AltPropString);

		//if (!namePropNames.name.empty()) ret->namedPropName(namePropNames.name);
		//if (!namePropNames.guid.empty()) ret->namedPropGuid(namePropNames.guid);

		//const auto szSmartView = smartview::parsePropertySmartView(lpPropVal, lpProp, lpNameID, sig, bIsAB, false);
		//if (!szSmartView.empty()) ret->smartView(szSmartView);

		return ret;
	}
} // namespace propertybag