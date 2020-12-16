#include <core/stdafx.h>
#include <core/model/mapiRowModel.h>
#include <core/smartview/SmartView.h>
#include <core/mapi/cache/namedProps.h>
#include <core/interpret/proptags.h>
#include <core/property/parseProperty.h>
#include <core/utility/error.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/mapiMemory.h>

namespace model
{
	_Check_return_ std::shared_ptr<model::mapiRowModel>
	propToModelInternal(const SPropValue* lpPropVal, const ULONG ulPropTag, const LPMAPIPROP lpProp, const bool bIsAB)
	{
		auto ret = std::make_shared<model::mapiRowModel>();
		ret->ulPropTag(ulPropTag);

		const auto PropTag = strings::format(L"0x%08X", ulPropTag);
		std::wstring PropString;
		std::wstring AltPropString;

		const auto propTagNames = proptags::PropTagToPropName(ulPropTag, bIsAB);

		if (!propTagNames.bestGuess.empty())
		{
			ret->name(propTagNames.bestGuess);
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

		return ret;
	}

	void addNameToModel(
		_Check_return_ std::shared_ptr<model::mapiRowModel> model,
		const LPMAPIPROP lpProp,
		_In_opt_ const SBinary* sig,
		const bool bIsAB)
	{
		const auto namePropNames = cache::NameIDToStrings(model->ulPropTag(), lpProp, nullptr, sig, bIsAB);
		if (!namePropNames.bestPidLid.empty())
		{
			model->name(namePropNames.bestPidLid);
		}
		else if (!namePropNames.name.empty())
		{
			model->name(namePropNames.name);
		}

		if (!namePropNames.name.empty()) model->namedPropName(namePropNames.name);
		if (!namePropNames.guid.empty()) model->namedPropGuid(namePropNames.guid);
	}

	_Check_return_ std::vector<std::shared_ptr<model::mapiRowModel>>
	propsToModels(ULONG cValues, const SPropValue* lpPropVals, const LPMAPIPROP lpProp, const bool bIsAB)
	{
		if (!cValues || !lpPropVals) return {};
		const SBinary* lpSigBin = nullptr;
		LPSPropValue lpMappingSigFromObject = nullptr;

		auto models = std::vector<std::shared_ptr<model::mapiRowModel>>{};
		for (ULONG i = 0; i < cValues; i++)
		{
			const auto prop = lpPropVals[i];
			models.push_back(model::propToModelInternal(&prop, prop.ulPropTag, lpProp, bIsAB));
		}

		// Pre cache and prop tag to name mappings we need
		// First gather tags
		auto tags = std::vector<ULONG>{};
		for (const auto model : models)
		{
			tags.push_back(model->ulPropTag());
		}

		if (!tags.empty())
		{
			auto lpTags = mapi::allocate<LPSPropTagArray>(CbNewSPropTagArray(tags.size()));
			if (lpTags)
			{
				lpTags->cValues = tags.size();
				ULONG i = 0;
				for (const auto tag:tags)
				{
					lpTags->aulPropTag[i++] = tag;
				}

				// Second get a mapping signature
				const auto lpMappingSig = mapi::FindProp(lpPropVals, cValues, PR_MAPPING_SIGNATURE);
				if (lpMappingSig && lpMappingSig->ulPropTag == PR_MAPPING_SIGNATURE)
					lpSigBin = &mapi::getBin(lpMappingSig);
				if (!lpSigBin && lpProp)
				{
					WC_H(HrGetOneProp(lpProp, PR_MAPPING_SIGNATURE, &lpMappingSigFromObject));
					if (lpMappingSigFromObject && lpMappingSigFromObject->ulPropTag == PR_MAPPING_SIGNATURE)
						lpSigBin = &mapi::getBin(lpMappingSigFromObject);
				}

				// Third, look up all the tags and cache named prop info
				(void) cache::GetNamesFromIDs(lpProp, lpSigBin, &lpTags, NULL);
				MAPIFreeBuffer(lpTags);
			}
		}

		// Finally, update our models with named prop info from the cache
		for (const auto model : models)
		{
			addNameToModel(model, lpProp, lpSigBin, bIsAB);
		}

		if (lpMappingSigFromObject) MAPIFreeBuffer(lpMappingSigFromObject);
		return models;
	}

	_Check_return_ std::shared_ptr<model::mapiRowModel>
	propToModel(const SPropValue* lpPropVal, const ULONG ulPropTag, const LPMAPIPROP lpProp, const bool bIsAB)
	{
		auto ret = propToModelInternal(lpPropVal, ulPropTag, lpProp, bIsAB);
		addNameToModel(ret, lpProp, nullptr, bIsAB);
		return ret;
	}
} // namespace model