#include <core/stdafx.h>
#include <core/model/mapiRowModel.h>
#include <core/smartview/SmartView.h>
#include <core/mapi/cache/namedProps.h>
#include <core/interpret/proptags.h>
#include <core/property/parseProperty.h>
#include <core/utility/error.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/mapiMemory.h>
#include <core/utility/registry.h>

namespace model
{
	_Check_return_ std::shared_ptr<model::mapiRowModel> propToModelInternal(
		_In_opt_ const SPropValue* lpPropVal,
		_In_ const ULONG ulPropTag,
		_In_opt_ const LPMAPIPROP lpProp,
		_In_ const bool bIsAB,
		_In_opt_ const SBinary* sig,
		_In_opt_ const MAPINAMEID* lpNameID)
	{
		if (!lpPropVal) return {};
		auto ret = std::make_shared<model::mapiRowModel>();
		ret->ulPropTag(ulPropTag);

		const auto propTagNames = proptags::PropTagToPropName(ulPropTag, bIsAB);
		const auto namePropNames = cache::NameIDToStrings(ulPropTag, lpProp, nullptr, sig, bIsAB);
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

		std::wstring PropString;
		std::wstring AltPropString;
		property::parseProperty(lpPropVal, &PropString, &AltPropString);
		ret->value(PropString);
		ret->altValue(AltPropString);

		if (!namePropNames.name.empty()) ret->namedPropName(namePropNames.name);
		if (!namePropNames.guid.empty()) ret->namedPropGuid(namePropNames.guid);

		const auto szSmartView = smartview::parsePropertySmartView(lpPropVal, lpProp, lpNameID, sig, bIsAB, false);
		if (!szSmartView.empty()) ret->smartView(szSmartView);

		return ret;
	}

	_Check_return_ std::vector<std::shared_ptr<model::mapiRowModel>> propsToModels(
		_In_ ULONG cValues,
		_In_opt_ const SPropValue* lpPropVals,
		_In_opt_ const LPMAPIPROP lpProp,
		_In_ const bool bIsAB)
	{
		if (!cValues || !lpPropVals) return {};
		const SBinary* lpSigBin = nullptr; // Will be borrowed. Do not free.
		LPSPropValue lpMappingSigFromObject = nullptr;
		std::vector<std::shared_ptr<cache::namedPropCacheEntry>> namedPropCacheEntries = {};

		if (!bIsAB && registry::parseNamedProps)
		{
			// Pre cache and prop tag to name mappings we need
			// First gather tags
			auto tags = std::vector<ULONG>{};
			for (ULONG i = 0; i < cValues; i++)
			{
				tags.push_back(lpPropVals[i].ulPropTag);
			}

			if (!tags.empty())
			{
				auto lpTags = mapi::allocate<LPSPropTagArray>(CbNewSPropTagArray(tags.size()));
				if (lpTags)
				{
					lpTags->cValues = tags.size();
					ULONG i = 0;
					for (const auto tag : tags)
					{
						mapi::setTag(lpTags, i++) = tag;
					}

					// Second get a mapping signature
					const SPropValue* lpMappingSig = mapi::FindProp(lpPropVals, cValues, PR_MAPPING_SIGNATURE);
					if (lpMappingSig && lpMappingSig->ulPropTag == PR_MAPPING_SIGNATURE)
						lpSigBin = &mapi::getBin(lpMappingSig);
					if (!lpSigBin && lpProp)
					{
						WC_H_S(HrGetOneProp(lpProp, PR_MAPPING_SIGNATURE, &lpMappingSigFromObject));
						if (lpMappingSigFromObject && lpMappingSigFromObject->ulPropTag == PR_MAPPING_SIGNATURE)
							lpSigBin = &mapi::getBin(lpMappingSigFromObject);
					}

					// Third, look up all the tags and cache named prop info
					namedPropCacheEntries = cache::GetNamesFromIDs(lpProp, lpSigBin, &lpTags, NULL);
					MAPIFreeBuffer(lpTags);
				}
			}
		}

		auto models = std::vector<std::shared_ptr<model::mapiRowModel>>{};
		for (ULONG i = 0; i < cValues; i++)
		{
			const auto prop = lpPropVals[i];
			const auto lpNameID =
				(namedPropCacheEntries.size() > i) ? namedPropCacheEntries[i]->getMapiNameId() : nullptr;
			models.push_back(model::propToModelInternal(&prop, prop.ulPropTag, lpProp, bIsAB, lpSigBin, lpNameID));
		}

		if (lpMappingSigFromObject) MAPIFreeBuffer(lpMappingSigFromObject);
		return models;
	}

	_Check_return_ std::shared_ptr<model::mapiRowModel> propToModel(
		_In_ const SPropValue* lpPropVal,
		_In_ const ULONG ulPropTag,
		_In_opt_ const LPMAPIPROP lpProp,
		_In_ const bool bIsAB)
	{
		return propToModelInternal(lpPropVal, ulPropTag, lpProp, bIsAB, nullptr, nullptr);
	}
} // namespace model