#include <core/stdafx.h>
#include <core/model/mapiRowModel.h>
#include <core/smartview/SmartView.h>
#include <core/mapi/cache/namedProps.h>
#include <core/interpret/proptags.h>
#include <core/property/parseProperty.h>

namespace model
{
	_Check_return_ std::vector<std::shared_ptr<model::mapiRowModel>>
	propsToModels(ULONG cValues, const SPropValue* lpPropVals, const LPMAPIPROP lpProp, const bool bIsAB)
	{
		if (!cValues || !lpPropVals) return {};

		auto models = std::vector<std::shared_ptr<model::mapiRowModel>>{};
		for (ULONG i = 0; i < cValues; i++)
		{
			const auto prop = lpPropVals[i];
			models.push_back(model::propToModel(&prop, prop.ulPropTag, lpProp, bIsAB));
		}

		return models;
	}

	_Check_return_ std::shared_ptr<model::mapiRowModel>
	propToModel(const SPropValue* lpPropVal, const ULONG ulPropTag, const LPMAPIPROP lpProp, const bool bIsAB)
	{
		auto ret = std::make_shared<model::mapiRowModel>();
		ret->ulPropTag(ulPropTag);

		const auto PropTag = strings::format(L"0x%08X", ulPropTag);
		std::wstring PropString;
		std::wstring AltPropString;

		// TODO: nameid and mapping signature
		//const auto namePropNames = cache::NameIDToStrings(ulPropTag, m_lpProp, lpNameID, lpMappingSignature, m_bIsAB);
		const auto namePropNames = cache::NameIDToStrings(ulPropTag, lpProp, nullptr, nullptr, bIsAB);
		const auto propTagNames = proptags::PropTagToPropName(ulPropTag, bIsAB);

		if (!propTagNames.bestGuess.empty())
		{
			ret->name(propTagNames.bestGuess);
		}
		else if (!namePropNames.bestPidLid.empty())
		{
			ret->name(namePropNames.bestPidLid);
		}
		else if (!namePropNames.name.empty())
		{
			ret->name(namePropNames.name);
		}

		ret->otherName(propTagNames.otherMatches);

		property::parseProperty(lpPropVal, &PropString, &AltPropString);
		ret->value(PropString);
		ret->altValue(AltPropString);

		auto szSmartView = smartview::parsePropertySmartView(
			lpPropVal,
			lpProp,
			nullptr, // lpNameID,
			nullptr, // lpMappingSignature,
			bIsAB,
			false); // Built from lpProp & lpMAPIProp
		if (!szSmartView.empty()) ret->smartView(szSmartView);
		if (!namePropNames.name.empty()) ret->namedPropName(namePropNames.name);
		if (!namePropNames.guid.empty()) ret->namedPropGuid(namePropNames.guid);

		return ret;
	}
} // namespace model