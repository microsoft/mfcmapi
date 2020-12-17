#include <core/stdafx.h>
#include <core/propertyBag/registryPropertyBag.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/mapiMemory.h>
#include <core/utility/strings.h>

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
			cchMaxValueNameLen++; // For null terminator
			for (DWORD dwIndex = 0; dwIndex < cValues; dwIndex++)
			{
				auto szBuf = std::wstring(cchMaxValueNameLen, L'\0');
				auto cchValLen = cchMaxValueNameLen;
				szBuf.clear();
				hRes = WC_W32(RegEnumValueW(
					m_hKey,
					dwIndex,
					const_cast<wchar_t*>(szBuf.c_str()),
					&cchValLen,
					nullptr, // lpReserved
					nullptr, // lpType
					nullptr, // lpData
					nullptr)); // lpcbData
				if (hRes == S_OK)
				{
					models.push_back(regToModel(szBuf));
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

	// TODO: Identify prop tags in the value name
	// Use type from tag to determine how to read data
	// Also use lpType
	// Convert data read to a MAPI prop Value
	// Figure out way to deal with S props
	// Figure out way to deal with named props
	_Check_return_ std::shared_ptr<model::mapiRowModel> registryPropertyBag::regToModel(_In_ const std::wstring val)
	{
		auto ret = std::make_shared<model::mapiRowModel>();
		ret->ulPropTag(1234);
		ret->name(val);
		ret->value(val);
		ret->altValue(L"foo");

		//ret->ulPropTag(ulPropTag);

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

		//ret->otherName(propTagNames.otherMatches);

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