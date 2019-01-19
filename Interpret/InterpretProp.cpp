#include <StdAfx.h>
#include <core/interpret/guid.h>
#include <Interpret/InterpretProp.h>
#include <core/mapi/extraPropTags.h>
#include <core/mapi/cache/namedPropCache.h>
#include <Interpret/SmartView/SmartView.h>
#include <Property/ParseProperty.h>
#include <core/utility/strings.h>
#include <unordered_map>
#include <core/addin/addin.h>
#include <core/utility/registry.h>
#include <core/utility/output.h>

namespace interpretprop
{
#define ulNoMatch 0xffffffff
	static WCHAR szPropSeparator[] = L", "; // STRING_OK

	std::wstring TagToString(ULONG ulPropTag, _In_opt_ LPMAPIPROP lpObj, bool bIsAB, bool bSingleLine)
	{
		std::wstring szTemp;

		auto namePropNames = cache::NameIDToStrings(ulPropTag, lpObj, nullptr, nullptr, bIsAB);

		auto propTagNames = PropTagToPropName(ulPropTag, bIsAB);

		std::wstring szFormatString;
		if (bSingleLine)
		{
			szFormatString = L"0x%1!08X! (%2)"; // STRING_OK
			if (!propTagNames.bestGuess.empty()) szFormatString += L": %3!ws!"; // STRING_OK
			if (!propTagNames.otherMatches.empty()) szFormatString += L": (%4!ws!)"; // STRING_OK
			if (!namePropNames.name.empty())
			{
				szFormatString += strings::loadstring(IDS_NAMEDPROPSINGLELINE);
			}

			if (!namePropNames.guid.empty())
			{
				szFormatString += strings::loadstring(IDS_GUIDSINGLELINE);
			}
		}
		else
		{
			szFormatString = strings::loadstring(IDS_TAGMULTILINE);
			if (!propTagNames.bestGuess.empty())
			{
				szFormatString += strings::loadstring(IDS_PROPNAMEMULTILINE);
			}

			if (!propTagNames.otherMatches.empty())
			{
				szFormatString += strings::loadstring(IDS_OTHERNAMESMULTILINE);
			}

			if (PROP_ID(ulPropTag) < 0x8000)
			{
				szFormatString += strings::loadstring(IDS_DASLPROPTAG);
			}
			else if (!namePropNames.dasl.empty())
			{
				szFormatString += strings::loadstring(IDS_DASLNAMED);
			}

			if (!namePropNames.name.empty())
			{
				szFormatString += strings::loadstring(IDS_NAMEPROPNAMEMULTILINE);
			}

			if (!namePropNames.guid.empty())
			{
				szFormatString += strings::loadstring(IDS_NAMEPROPGUIDMULTILINE);
			}
		}

		auto szRet = strings::formatmessage(
			szFormatString.c_str(),
			ulPropTag,
			TypeToString(ulPropTag).c_str(),
			propTagNames.bestGuess.c_str(),
			propTagNames.otherMatches.c_str(),
			namePropNames.name.c_str(),
			namePropNames.guid.c_str(),
			namePropNames.dasl.c_str());

		if (fIsSet(DBGTest))
		{
			static size_t cchMaxBuff = 0;
			auto cchBuff = szRet.length();
			cchMaxBuff = max(cchBuff, cchMaxBuff);
			output::DebugPrint(
				DBGTest,
				L"TagToString parsing 0x%08X returned %u chars - max %u\n",
				ulPropTag,
				static_cast<UINT>(cchBuff),
				static_cast<UINT>(cchMaxBuff));
		}

		return szRet;
	}

	std::wstring ProblemArrayToString(_In_ const SPropProblemArray& problems)
	{
		std::wstring szOut;
		for (ULONG i = 0; i < problems.cProblem; i++)
		{
			szOut += strings::formatmessage(
				IDS_PROBLEMARRAY,
				problems.aProblem[i].ulIndex,
				TagToString(problems.aProblem[i].ulPropTag, nullptr, false, false).c_str(),
				problems.aProblem[i].scode,
				error::ErrorNameFromErrorCode(problems.aProblem[i].scode).c_str());
		}

		return szOut;
	}

	std::wstring MAPIErrToString(ULONG ulFlags, _In_ const MAPIERROR& err)
	{
		auto szOut = strings::formatmessage(
			ulFlags & MAPI_UNICODE ? IDS_MAPIERRUNICODE : IDS_MAPIERRANSI,
			err.ulVersion,
			err.lpszError,
			err.lpszComponent,
			err.ulLowLevelError,
			error::ErrorNameFromErrorCode(err.ulLowLevelError).c_str(),
			err.ulContext);

		return szOut;
	}

	std::wstring TnefProblemArrayToString(_In_ const STnefProblemArray& error)
	{
		std::wstring szOut;
		for (ULONG iError = 0; iError < error.cProblem; iError++)
		{
			szOut += strings::formatmessage(
				IDS_TNEFPROBARRAY,
				error.aProblem[iError].ulComponent,
				error.aProblem[iError].ulAttribute,
				TagToString(error.aProblem[iError].ulPropTag, nullptr, false, false).c_str(),
				error.aProblem[iError].scode,
				error::ErrorNameFromErrorCode(error.aProblem[iError].scode).c_str());
		}

		return szOut;
	}

	// There may be restrictions with over 100 nested levels, but we're not going to try to parse them
#define _MaxRestrictionNesting 100

	std::wstring RestrictionToString(_In_ const _SRestriction* lpRes, _In_opt_ LPMAPIPROP lpObj, ULONG ulTabLevel)
	{
		if (!lpRes)
		{
			return strings::loadstring(IDS_NULLRES);
		}
		if (ulTabLevel > _MaxRestrictionNesting)
		{
			return strings::loadstring(IDS_RESDEPTHEXCEEDED);
		}

		std::vector<std::wstring> resString;
		std::wstring szProp;
		std::wstring szAltProp;

		std::wstring szTabs;
		for (ULONG i = 0; i < ulTabLevel; i++)
		{
			szTabs += L"\t"; // STRING_OK
		}

		std::wstring szPropNum;
		auto szFlags = InterpretFlags(flagRestrictionType, lpRes->rt);
		resString.push_back(strings::formatmessage(IDS_RESTYPE, szTabs.c_str(), lpRes->rt, szFlags.c_str()));

		switch (lpRes->rt)
		{
		case RES_COMPAREPROPS:
			szFlags = InterpretFlags(flagRelop, lpRes->res.resCompareProps.relop);
			resString.push_back(strings::formatmessage(
				IDS_RESCOMPARE,
				szTabs.c_str(),
				szFlags.c_str(),
				lpRes->res.resCompareProps.relop,
				TagToString(lpRes->res.resCompareProps.ulPropTag1, lpObj, false, true).c_str(),
				TagToString(lpRes->res.resCompareProps.ulPropTag2, lpObj, false, true).c_str()));
			break;
		case RES_AND:
			resString.push_back(strings::formatmessage(IDS_RESANDCOUNT, szTabs.c_str(), lpRes->res.resAnd.cRes));
			if (lpRes->res.resAnd.lpRes)
			{
				for (ULONG i = 0; i < lpRes->res.resAnd.cRes; i++)
				{
					resString.push_back(strings::formatmessage(IDS_RESANDPOINTER, szTabs.c_str(), i));
					resString.push_back(RestrictionToString(&lpRes->res.resAnd.lpRes[i], lpObj, ulTabLevel + 1));
				}
			}
			break;
		case RES_OR:
			resString.push_back(strings::formatmessage(IDS_RESORCOUNT, szTabs.c_str(), lpRes->res.resOr.cRes));
			if (lpRes->res.resOr.lpRes)
			{
				for (ULONG i = 0; i < lpRes->res.resOr.cRes; i++)
				{
					resString.push_back(strings::formatmessage(IDS_RESORPOINTER, szTabs.c_str(), i));
					resString.push_back(RestrictionToString(&lpRes->res.resOr.lpRes[i], lpObj, ulTabLevel + 1));
				}
			}
			break;
		case RES_NOT:
			resString.push_back(strings::formatmessage(IDS_RESNOT, szTabs.c_str(), lpRes->res.resNot.ulReserved));
			resString.push_back(RestrictionToString(lpRes->res.resNot.lpRes, lpObj, ulTabLevel + 1));
			break;
		case RES_COUNT:
			// RES_COUNT and RES_NOT look the same, so we use the resNot member here
			resString.push_back(strings::formatmessage(IDS_RESCOUNT, szTabs.c_str(), lpRes->res.resNot.ulReserved));
			resString.push_back(RestrictionToString(lpRes->res.resNot.lpRes, lpObj, ulTabLevel + 1));
			break;
		case RES_CONTENT:
			szFlags = InterpretFlags(flagFuzzyLevel, lpRes->res.resContent.ulFuzzyLevel);
			resString.push_back(strings::formatmessage(
				IDS_RESCONTENT,
				szTabs.c_str(),
				szFlags.c_str(),
				lpRes->res.resContent.ulFuzzyLevel,
				TagToString(lpRes->res.resContent.ulPropTag, lpObj, false, true).c_str()));
			if (lpRes->res.resContent.lpProp)
			{
				InterpretProp(lpRes->res.resContent.lpProp, &szProp, &szAltProp);
				resString.push_back(strings::formatmessage(
					IDS_RESCONTENTPROP,
					szTabs.c_str(),
					TagToString(lpRes->res.resContent.lpProp->ulPropTag, lpObj, false, true).c_str(),
					szProp.c_str(),
					szAltProp.c_str()));
			}
			break;
		case RES_PROPERTY:
			szFlags = InterpretFlags(flagRelop, lpRes->res.resProperty.relop);
			resString.push_back(strings::formatmessage(
				IDS_RESPROP,
				szTabs.c_str(),
				szFlags.c_str(),
				lpRes->res.resProperty.relop,
				TagToString(lpRes->res.resProperty.ulPropTag, lpObj, false, true).c_str()));
			if (lpRes->res.resProperty.lpProp)
			{
				InterpretProp(lpRes->res.resProperty.lpProp, &szProp, &szAltProp);
				resString.push_back(strings::formatmessage(
					IDS_RESPROPPROP,
					szTabs.c_str(),
					TagToString(lpRes->res.resProperty.lpProp->ulPropTag, lpObj, false, true).c_str(),
					szProp.c_str(),
					szAltProp.c_str()));
				szPropNum = smartview::InterpretNumberAsString(
					lpRes->res.resProperty.lpProp->Value,
					lpRes->res.resProperty.lpProp->ulPropTag,
					NULL,
					nullptr,
					nullptr,
					false);
				if (!szPropNum.empty())
				{
					resString.push_back(
						strings::formatmessage(IDS_RESPROPPROPFLAGS, szTabs.c_str(), szPropNum.c_str()));
				}
			}
			break;
		case RES_BITMASK:
			szFlags = InterpretFlags(flagBitmask, lpRes->res.resBitMask.relBMR);
			resString.push_back(strings::formatmessage(
				IDS_RESBITMASK,
				szTabs.c_str(),
				szFlags.c_str(),
				lpRes->res.resBitMask.relBMR,
				lpRes->res.resBitMask.ulMask));
			szPropNum =
				smartview::InterpretNumberAsStringProp(lpRes->res.resBitMask.ulMask, lpRes->res.resBitMask.ulPropTag);
			if (!szPropNum.empty())
			{
				resString.push_back(strings::formatmessage(IDS_RESBITMASKFLAGS, szPropNum.c_str()));
			}

			resString.push_back(strings::formatmessage(
				IDS_RESBITMASKTAG,
				szTabs.c_str(),
				TagToString(lpRes->res.resBitMask.ulPropTag, lpObj, false, true).c_str()));
			break;
		case RES_SIZE:
			szFlags = InterpretFlags(flagRelop, lpRes->res.resSize.relop);
			resString.push_back(strings::formatmessage(
				IDS_RESSIZE,
				szTabs.c_str(),
				szFlags.c_str(),
				lpRes->res.resSize.relop,
				lpRes->res.resSize.cb,
				TagToString(lpRes->res.resSize.ulPropTag, lpObj, false, true).c_str()));
			break;
		case RES_EXIST:
			resString.push_back(strings::formatmessage(
				IDS_RESEXIST,
				szTabs.c_str(),
				TagToString(lpRes->res.resExist.ulPropTag, lpObj, false, true).c_str(),
				lpRes->res.resExist.ulReserved1,
				lpRes->res.resExist.ulReserved2));
			break;
		case RES_SUBRESTRICTION:
			resString.push_back(strings::formatmessage(
				IDS_RESSUBRES, szTabs.c_str(), TagToString(lpRes->res.resSub.ulSubObject, lpObj, false, true).c_str()));
			resString.push_back(RestrictionToString(lpRes->res.resSub.lpRes, lpObj, ulTabLevel + 1));
			break;
		case RES_COMMENT:
			resString.push_back(strings::formatmessage(IDS_RESCOMMENT, szTabs.c_str(), lpRes->res.resComment.cValues));
			if (lpRes->res.resComment.lpProp)
			{
				for (ULONG i = 0; i < lpRes->res.resComment.cValues; i++)
				{
					InterpretProp(&lpRes->res.resComment.lpProp[i], &szProp, &szAltProp);
					resString.push_back(strings::formatmessage(
						IDS_RESCOMMENTPROPS,
						szTabs.c_str(),
						i,
						TagToString(lpRes->res.resComment.lpProp[i].ulPropTag, lpObj, false, true).c_str(),
						szProp.c_str(),
						szAltProp.c_str()));
				}
			}

			resString.push_back(strings::formatmessage(IDS_RESCOMMENTRES, szTabs.c_str()));
			resString.push_back(RestrictionToString(lpRes->res.resComment.lpRes, lpObj, ulTabLevel + 1));
			break;
		case RES_ANNOTATION:
			resString.push_back(
				strings::formatmessage(IDS_RESANNOTATION, szTabs.c_str(), lpRes->res.resComment.cValues));
			if (lpRes->res.resComment.lpProp)
			{
				for (ULONG i = 0; i < lpRes->res.resComment.cValues; i++)
				{
					InterpretProp(&lpRes->res.resComment.lpProp[i], &szProp, &szAltProp);
					resString.push_back(strings::formatmessage(
						IDS_RESANNOTATIONPROPS,
						szTabs.c_str(),
						i,
						TagToString(lpRes->res.resComment.lpProp[i].ulPropTag, lpObj, false, true).c_str(),
						szProp.c_str(),
						szAltProp.c_str()));
				}
			}

			resString.push_back(strings::formatmessage(IDS_RESANNOTATIONRES, szTabs.c_str()));
			resString.push_back(RestrictionToString(lpRes->res.resComment.lpRes, lpObj, ulTabLevel + 1));
			break;
		}

		return strings::join(resString, L"\r\n");
	}

	std::wstring RestrictionToString(_In_ const _SRestriction* lpRes, _In_opt_ LPMAPIPROP lpObj)
	{
		return RestrictionToString(lpRes, lpObj, 0);
	}

	std::wstring AdrListToString(_In_ const ADRLIST& adrList)
	{
		std::wstring szProp;
		std::wstring szAltProp;
		auto adrstring = strings::formatmessage(IDS_ADRLISTCOUNT, adrList.cEntries);

		for (ULONG i = 0; i < adrList.cEntries; i++)
		{
			adrstring += strings::formatmessage(IDS_ADRLISTENTRIESCOUNT, i, adrList.aEntries[i].cValues);

			for (ULONG j = 0; j < adrList.aEntries[i].cValues; j++)
			{
				InterpretProp(&adrList.aEntries[i].rgPropVals[j], &szProp, &szAltProp);
				adrstring += strings::formatmessage(
					IDS_ADRLISTENTRY,
					i,
					j,
					TagToString(adrList.aEntries[i].rgPropVals[j].ulPropTag, nullptr, false, false).c_str(),
					szProp.c_str(),
					szAltProp.c_str());
			}
		}

		return adrstring;
	}

	_Check_return_ std::wstring ActionToString(_In_ const ACTION& action)
	{
		std::wstring szProp;
		std::wstring szAltProp;
		auto szFlags = InterpretFlags(flagAccountType, action.acttype);
		auto szFlags2 = InterpretFlags(flagRuleFlag, action.ulFlags);
		auto actstring = strings::formatmessage(
			IDS_ACTION,
			action.acttype,
			szFlags.c_str(),
			RestrictionToString(action.lpRes, nullptr).c_str(),
			action.ulFlags,
			szFlags2.c_str());

		switch (action.acttype)
		{
		case OP_MOVE:
		case OP_COPY:
		{
			auto sBinStore =
				SBinary{action.actMoveCopy.cbStoreEntryId, reinterpret_cast<LPBYTE>(action.actMoveCopy.lpStoreEntryId)};
			auto sBinFld =
				SBinary{action.actMoveCopy.cbFldEntryId, reinterpret_cast<LPBYTE>(action.actMoveCopy.lpFldEntryId)};

			actstring += strings::formatmessage(
				IDS_ACTIONOPMOVECOPY,
				strings::BinToHexString(&sBinStore, true).c_str(),
				strings::BinToTextString(&sBinStore, false).c_str(),
				strings::BinToHexString(&sBinFld, true).c_str(),
				strings::BinToTextString(&sBinFld, false).c_str());
			break;
		}
		case OP_REPLY:
		case OP_OOF_REPLY:
		{

			auto sBin = SBinary{action.actReply.cbEntryId, reinterpret_cast<LPBYTE>(action.actReply.lpEntryId)};
			auto szGUID = guid::GUIDToStringAndName(&action.actReply.guidReplyTemplate);

			actstring += strings::formatmessage(
				IDS_ACTIONOPREPLY,
				strings::BinToHexString(&sBin, true).c_str(),
				strings::BinToTextString(&sBin, false).c_str(),
				szGUID.c_str());
			break;
		}
		case OP_DEFER_ACTION:
		{
			auto sBin = SBinary{action.actDeferAction.cbData, static_cast<LPBYTE>(action.actDeferAction.pbData)};

			actstring += strings::formatmessage(
				IDS_ACTIONOPDEFER,
				strings::BinToHexString(&sBin, true).c_str(),
				strings::BinToTextString(&sBin, false).c_str());
			break;
		}
		case OP_BOUNCE:
		{
			szFlags = InterpretFlags(flagBounceCode, action.scBounceCode);
			actstring += strings::formatmessage(IDS_ACTIONOPBOUNCE, action.scBounceCode, szFlags.c_str());
			break;
		}
		case OP_FORWARD:
		case OP_DELEGATE:
		{
			actstring += strings::formatmessage(IDS_ACTIONOPFORWARDDEL);
			if (action.lpadrlist)
			{
				actstring += AdrListToString(*action.lpadrlist);
			}
			else
			{
				actstring += strings::loadstring(IDS_ADRLISTNULL);
			}

			break;
		}

		case OP_TAG:
		{
			InterpretProp(const_cast<LPSPropValue>(&action.propTag), &szProp, &szAltProp);
			actstring += strings::formatmessage(
				IDS_ACTIONOPTAG,
				TagToString(action.propTag.ulPropTag, nullptr, false, true).c_str(),
				szProp.c_str(),
				szAltProp.c_str());
			break;
		}
		default:
			break;
		}

		switch (action.acttype)
		{
		case OP_REPLY:
		{
			szFlags = InterpretFlags(flagOPReply, action.ulActionFlavor);
			break;
		}
		case OP_FORWARD:
		{
			szFlags = InterpretFlags(flagOpForward, action.ulActionFlavor);
			break;
		}
		default:
			break;
		}

		actstring += strings::formatmessage(IDS_ACTIONFLAVOR, action.ulActionFlavor, szFlags.c_str());

		if (!action.lpPropTagArray)
		{
			actstring += strings::loadstring(IDS_ACTIONTAGARRAYNULL);
		}
		else
		{
			actstring += strings::formatmessage(IDS_ACTIONTAGARRAYCOUNT, action.lpPropTagArray->cValues);
			for (ULONG i = 0; i < action.lpPropTagArray->cValues; i++)
			{
				actstring += strings::formatmessage(
					IDS_ACTIONTAGARRAYTAG,
					i,
					TagToString(action.lpPropTagArray->aulPropTag[i], nullptr, false, false).c_str());
			}
		}

		return actstring;
	}

	std::wstring ActionsToString(_In_ const ACTIONS& actions)
	{
		auto szFlags = InterpretFlags(flagRulesVersion, actions.ulVersion);
		auto actstring =
			strings::formatmessage(IDS_ACTIONSMEMBERS, actions.ulVersion, szFlags.c_str(), actions.cActions);

		for (ULONG i = 0; i < actions.cActions; i++)
		{
			actstring += strings::formatmessage(IDS_ACTIONSACTION, i);
			actstring += ActionToString(actions.lpAction[i]);
		}

		return actstring;
	}

	/***************************************************************************
	Name: InterpretProp
	Purpose: Evaluate a property value and return a string representing the property.
	Parameters:
	In:
	LPSPropValue lpProp: Property to be evaluated
	Out:
	wstring* tmpPropString: String representing property value
	wstring* tmpAltPropString: Alternative string representation
	Comment: Add new Property IDs as they become known
	***************************************************************************/
	void InterpretProp(
		_In_ const _SPropValue* lpProp,
		_In_opt_ std::wstring* PropString,
		_In_opt_ std::wstring* AltPropString)
	{
		if (!lpProp) return;

		auto parsedProperty = property::ParseProperty(lpProp);

		if (PropString) *PropString = parsedProperty.toString();
		if (AltPropString) *AltPropString = parsedProperty.toAltString();
	}

	std::wstring TypeToString(ULONG ulPropTag)
	{
		std::wstring tmpPropType;

		auto bNeedInstance = false;
		if (ulPropTag & MV_INSTANCE)
		{
			ulPropTag &= ~MV_INSTANCE;
			bNeedInstance = true;
		}

		auto bTypeFound = false;

		for (const auto& propType : PropTypeArray)
		{
			if (propType.ulValue == PROP_TYPE(ulPropTag))
			{
				tmpPropType = propType.lpszName;
				bTypeFound = true;
				break;
			}
		}

		if (!bTypeFound) tmpPropType = strings::format(L"0x%04x", PROP_TYPE(ulPropTag)); // STRING_OK

		if (bNeedInstance) tmpPropType += L" | MV_INSTANCE"; // STRING_OK
		return tmpPropType;
	}

	// Compare tag sort order.
	bool CompareTagsSortOrder(int a1, int a2)
	{
		const auto lpTag1 = &PropTagArray[a1];
		const auto lpTag2 = &PropTagArray[a2];

		if (lpTag1->ulSortOrder < lpTag2->ulSortOrder) return false;
		if (lpTag1->ulSortOrder == lpTag2->ulSortOrder)
		{
			return wcscmp(lpTag1->lpszName, lpTag2->lpszName) <= 0;
		}
		return true;
	}

	// Searches an array for a target number.
	// Search is done with a mask
	// Partial matches are those that match with the mask applied
	// Exact matches are those that match without the mask applied
	// lpUlNumPartials will exclude count of exact matches
	// if it wants just the true partial matches.
	// If no hits, then ulNoMatch should be returned for lpulFirstExact and/or lpulFirstPartial
	void FindTagArrayMatches(
		_In_ ULONG ulTarget,
		bool bIsAB,
		const std::vector<NAME_ARRAY_ENTRY_V2>& MyArray,
		std::vector<ULONG>& ulExacts,
		std::vector<ULONG>& ulPartials)
	{
		if (!(ulTarget & PROP_TAG_MASK)) // not dealing with a full prop tag
		{
			ulTarget = PROP_TAG(PT_UNSPECIFIED, ulTarget);
		}

		ULONG ulLowerBound = 0;
		auto ulUpperBound = static_cast<ULONG>(MyArray.size() - 1); // size-1 is the last entry
		auto ulMidPoint = (ulUpperBound + ulLowerBound) / 2;
		ULONG ulFirstMatch = ulNoMatch;
		const auto ulMaskedTarget = ulTarget & PROP_TAG_MASK;

		// Short circuit property IDs with the high bit set if bIsAB wasn't passed
		if (!bIsAB && ulTarget & 0x80000000) return;

		// Find A partial match
		while (ulUpperBound - ulLowerBound > 1)
		{
			if (ulMaskedTarget == (PROP_TAG_MASK & MyArray[ulMidPoint].ulValue))
			{
				ulFirstMatch = ulMidPoint;
				break;
			}

			if (ulMaskedTarget < (PROP_TAG_MASK & MyArray[ulMidPoint].ulValue))
			{
				ulUpperBound = ulMidPoint;
			}
			else if (ulMaskedTarget > (PROP_TAG_MASK & MyArray[ulMidPoint].ulValue))
			{
				ulLowerBound = ulMidPoint;
			}

			ulMidPoint = (ulUpperBound + ulLowerBound) / 2;
		}

		// When we get down to two points, we may have only checked one of them
		// Make sure we've checked the other
		if (ulMaskedTarget == (PROP_TAG_MASK & MyArray[ulUpperBound].ulValue))
		{
			ulFirstMatch = ulUpperBound;
		}
		else if (ulMaskedTarget == (PROP_TAG_MASK & MyArray[ulLowerBound].ulValue))
		{
			ulFirstMatch = ulLowerBound;
		}

		// Check that we got a match
		if (ulNoMatch != ulFirstMatch)
		{
			// Scan backwards to find the first partial match
			while (ulFirstMatch > 0 && ulMaskedTarget == (PROP_TAG_MASK & MyArray[ulFirstMatch - 1].ulValue))
			{
				ulFirstMatch = ulFirstMatch - 1;
			}

			for (auto ulCur = ulFirstMatch;
				 ulCur < MyArray.size() && ulMaskedTarget == (PROP_TAG_MASK & MyArray[ulCur].ulValue);
				 ulCur++)
			{
				if (ulTarget == MyArray[ulCur].ulValue)
				{
					ulExacts.push_back(ulCur);
				}
				else
				{
					ulPartials.push_back(ulCur);
				}
			}

			if (ulExacts.size()) sort(ulExacts.begin(), ulExacts.end(), CompareTagsSortOrder);
			if (ulPartials.size()) sort(ulPartials.begin(), ulPartials.end(), CompareTagsSortOrder);
		}
	}

	std::unordered_map<ULONG64, PropTagNames> g_PropNames;

	PropTagNames PropTagToPropName(ULONG ulPropTag, bool bIsAB)
	{
		auto ulKey = (bIsAB ? static_cast<ULONG64>(1) << 32 : 0) | ulPropTag;

		const auto match = g_PropNames.find(ulKey);
		if (match != g_PropNames.end())
		{
			return match->second;
		}

		std::vector<ULONG> ulExacts;
		std::vector<ULONG> ulPartials;
		FindTagArrayMatches(ulPropTag, bIsAB, PropTagArray, ulExacts, ulPartials);

		PropTagNames entry;

		if (ulExacts.size())
		{
			entry.bestGuess = PropTagArray[ulExacts.front()].lpszName;
			ulExacts.erase(ulExacts.begin());

			for (const auto& ulMatch : ulExacts)
			{
				if (!entry.otherMatches.empty())
				{
					entry.otherMatches += szPropSeparator;
				}

				entry.otherMatches += PropTagArray[ulMatch].lpszName;
			}
		}

		if (ulPartials.size())
		{
			if (entry.bestGuess.empty())
			{
				entry.bestGuess = PropTagArray[ulPartials.front()].lpszName;
				ulPartials.erase(ulPartials.begin());
			}

			for (const auto& ulMatch : ulPartials)
			{
				if (!entry.otherMatches.empty())
				{
					entry.otherMatches += szPropSeparator;
				}

				entry.otherMatches += PropTagArray[ulMatch].lpszName;
			}
		}

		g_PropNames.insert({ulKey, entry});

		return entry;
	}

	// Strictly does a lookup in the array. Does not convert otherwise
	_Check_return_ ULONG LookupPropName(_In_ const std::wstring& lpszPropName)
	{
		auto trimName = strings::trim(lpszPropName);
		if (trimName.empty()) return 0;

		for (auto& tag : PropTagArray)
		{
			if (0 == lstrcmpiW(trimName.c_str(), tag.lpszName))
			{
				return tag.ulValue;
			}
		}

		return 0;
	}

	_Check_return_ ULONG PropNameToPropTag(_In_ const std::wstring& lpszPropName)
	{
		if (lpszPropName.empty()) return 0;

		const auto ulTag = strings::wstringToUlong(lpszPropName, 16);
		if (ulTag != NULL)
		{
			return ulTag;
		}

		return LookupPropName(lpszPropName);
	}

	_Check_return_ ULONG PropTypeNameToPropType(_In_ const std::wstring& lpszPropType)
	{
		if (lpszPropType.empty() || PropTypeArray.empty()) return PT_UNSPECIFIED;

		// Check for numbers first before trying the string as an array lookup.
		// This will translate '0x102' to 0x102, 0x3 to 3, etc.
		const auto ulType = strings::wstringToUlong(lpszPropType, 16);
		if (ulType != NULL) return ulType;

		auto ulPropType = PT_UNSPECIFIED;

		for (const auto& propType : PropTypeArray)
		{
			if (0 == lstrcmpiW(lpszPropType.c_str(), propType.lpszName))
			{
				ulPropType = propType.ulValue;
				break;
			}
		}

		return ulPropType;
	}

	// Interprets a flag value according to a flag name and returns a string
	// Will not return a string if the flag name is not recognized
	std::wstring InterpretFlags(ULONG ulFlagName, LONG lFlagValue)
	{
		ULONG ulCurEntry = 0;

		if (FlagArray.empty()) return L"";

		while (ulCurEntry < FlagArray.size() && FlagArray[ulCurEntry].ulFlagName != ulFlagName)
		{
			ulCurEntry++;
		}

		// Don't run off the end of the array
		if (FlagArray.size() == ulCurEntry) return L"";
		if (FlagArray[ulCurEntry].ulFlagName != ulFlagName) return L"";

		// We've matched our flag name to the array - we SHOULD return a string at this point
		auto bNeedSeparator = false;

		auto lTempValue = lFlagValue;
		std::wstring szTempString;
		for (; FlagArray[ulCurEntry].ulFlagName == ulFlagName; ulCurEntry++)
		{
			if (flagFLAG == FlagArray[ulCurEntry].ulFlagType)
			{
				if (FlagArray[ulCurEntry].lFlagValue & lTempValue)
				{
					if (bNeedSeparator)
					{
						szTempString += L" | "; // STRING_OK
					}

					szTempString += FlagArray[ulCurEntry].lpszName;
					lTempValue &= ~FlagArray[ulCurEntry].lFlagValue;
					bNeedSeparator = true;
				}
			}
			else if (flagVALUE == FlagArray[ulCurEntry].ulFlagType)
			{
				if (FlagArray[ulCurEntry].lFlagValue == lTempValue)
				{
					if (bNeedSeparator)
					{
						szTempString += L" | "; // STRING_OK
					}

					szTempString += FlagArray[ulCurEntry].lpszName;
					lTempValue = 0;
					bNeedSeparator = true;
				}
			}
			else if (flagVALUEHIGHBYTES == FlagArray[ulCurEntry].ulFlagType)
			{
				if (FlagArray[ulCurEntry].lFlagValue == (lTempValue >> 16 & 0xFFFF))
				{
					if (bNeedSeparator)
					{
						szTempString += L" | "; // STRING_OK
					}

					szTempString += FlagArray[ulCurEntry].lpszName;
					lTempValue = lTempValue - (FlagArray[ulCurEntry].lFlagValue << 16);
					bNeedSeparator = true;
				}
			}
			else if (flagVALUE3RDBYTE == FlagArray[ulCurEntry].ulFlagType)
			{
				if (FlagArray[ulCurEntry].lFlagValue == (lTempValue >> 8 & 0xFF))
				{
					if (bNeedSeparator)
					{
						szTempString += L" | "; // STRING_OK
					}

					szTempString += FlagArray[ulCurEntry].lpszName;
					lTempValue = lTempValue - (FlagArray[ulCurEntry].lFlagValue << 8);
					bNeedSeparator = true;
				}
			}
			else if (flagVALUE4THBYTE == FlagArray[ulCurEntry].ulFlagType)
			{
				if (FlagArray[ulCurEntry].lFlagValue == (lTempValue & 0xFF))
				{
					if (bNeedSeparator)
					{
						szTempString += L" | "; // STRING_OK
					}

					szTempString += FlagArray[ulCurEntry].lpszName;
					lTempValue = lTempValue - FlagArray[ulCurEntry].lFlagValue;
					bNeedSeparator = true;
				}
			}
			else if (flagVALUELOWERNIBBLE == FlagArray[ulCurEntry].ulFlagType)
			{
				if (FlagArray[ulCurEntry].lFlagValue == (lTempValue & 0x0F))
				{
					if (bNeedSeparator)
					{
						szTempString += L" | "; // STRING_OK
					}

					szTempString += FlagArray[ulCurEntry].lpszName;
					lTempValue = lTempValue - FlagArray[ulCurEntry].lFlagValue;
					bNeedSeparator = true;
				}
			}
			else if (flagCLEARBITS == FlagArray[ulCurEntry].ulFlagType)
			{
				// find any bits we need to clear
				const auto lClearedBits = FlagArray[ulCurEntry].lFlagValue & lTempValue;
				// report what we found
				if (0 != lClearedBits)
				{
					if (bNeedSeparator)
					{
						szTempString += L" | "; // STRING_OK
					}

					szTempString += strings::format(L"0x%X", lClearedBits); // STRING_OK
						// clear the bits out
					lTempValue &= ~FlagArray[ulCurEntry].lFlagValue;
					bNeedSeparator = true;
				}
			}
		}

		// We know if we've found anything already because bNeedSeparator will be true
		// If bNeedSeparator isn't true, we found nothing and need to tack on
		// Otherwise, it's true, and we only tack if lTempValue still has something in it
		if (!bNeedSeparator || lTempValue)
		{
			if (bNeedSeparator)
			{
				szTempString += L" | "; // STRING_OK
			}

			szTempString += strings::format(L"0x%X", lTempValue); // STRING_OK
		}

		return szTempString;
	}

	// Returns a list of all known flags/values for a flag name.
	// For instance, for flagFuzzyLevel, would return:
	// \r\n0x00000000 FL_FULLSTRING\r\n\
	 // 0x00000001 FL_SUBSTRING\r\n\
	 // 0x00000002 FL_PREFIX\r\n\
	 // 0x00010000 FL_IGNORECASE\r\n\
	 // 0x00020000 FL_IGNORENONSPACE\r\n\
	 // 0x00040000 FL_LOOSE
	//
	// Since the string is always appended to a prompt we include \r\n at the start
	std::wstring AllFlagsToString(ULONG ulFlagName, bool bHex)
	{
		std::wstring szFlagString;
		if (!ulFlagName) return szFlagString;
		if (FlagArray.empty()) return szFlagString;

		ULONG ulCurEntry = 0;
		std::wstring szTempString;

		while (ulCurEntry < FlagArray.size() && FlagArray[ulCurEntry].ulFlagName != ulFlagName)
		{
			ulCurEntry++;
		}

		if (FlagArray[ulCurEntry].ulFlagName != ulFlagName) return szFlagString;

		// We've matched our flag name to the array - we SHOULD return a string at this point
		for (; FlagArray[ulCurEntry].ulFlagName == ulFlagName; ulCurEntry++)
		{
			if (flagCLEARBITS == FlagArray[ulCurEntry].ulFlagType)
			{
				// keep going
			}
			else
			{
				if (bHex)
				{
					szFlagString += strings::formatmessage(
						IDS_FLAGTOSTRINGHEX, FlagArray[ulCurEntry].lFlagValue, FlagArray[ulCurEntry].lpszName);
				}
				else
				{
					szFlagString += strings::formatmessage(
						IDS_FLAGTOSTRINGDEC, FlagArray[ulCurEntry].lFlagValue, FlagArray[ulCurEntry].lpszName);
				}
			}
		}

		return szFlagString;
	}
} // namespace interpretprop