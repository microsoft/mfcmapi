#include "StdAfx.h"
#include <Interpret/Guids.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/InterpretProp2.h>
#include <Interpret/ExtraPropTags.h>
#include <MAPI/NamedPropCache.h>
#include <Interpret/SmartView/SmartView.h>
#include "Property/ParseProperty.h"
#include <Interpret/String.h>

std::wstring TagToString(ULONG ulPropTag, _In_opt_ LPMAPIPROP lpObj, bool bIsAB, bool bSingleLine)
{
	std::wstring szTemp;

	auto namePropNames = NameIDToStrings(
		ulPropTag,
		lpObj,
		nullptr,
		nullptr,
		bIsAB);

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

	std::wstring szRet = strings::formatmessage(szFormatString.c_str(),
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
		DebugPrint(DBGTest, L"TagToString parsing 0x%08X returned %u chars - max %u\n", ulPropTag, static_cast<UINT>(cchBuff), static_cast<UINT>(cchMaxBuff));
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
			ErrorNameFromErrorCode(problems.aProblem[i].scode).c_str());
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
		ErrorNameFromErrorCode(err.ulLowLevelError).c_str(),
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
			ErrorNameFromErrorCode(error.aProblem[iError].scode).c_str());
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
		resString.push_back(strings::formatmessage(
			IDS_RESNOT,
			szTabs.c_str(),
			lpRes->res.resNot.ulReserved));
		resString.push_back(RestrictionToString(lpRes->res.resNot.lpRes, lpObj, ulTabLevel + 1));
		break;
	case RES_COUNT:
		// RES_COUNT and RES_NOT look the same, so we use the resNot member here
		resString.push_back(strings::formatmessage(
			IDS_RESCOUNT,
			szTabs.c_str(),
			lpRes->res.resNot.ulReserved));
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
			szPropNum = smartview::InterpretNumberAsString(lpRes->res.resProperty.lpProp->Value, lpRes->res.resProperty.lpProp->ulPropTag, NULL, nullptr, nullptr, false);
			if (!szPropNum.empty())
			{
				resString.push_back(strings::formatmessage(IDS_RESPROPPROPFLAGS, szTabs.c_str(), szPropNum.c_str()));
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
		szPropNum = smartview::InterpretNumberAsStringProp(lpRes->res.resBitMask.ulMask, lpRes->res.resBitMask.ulPropTag);
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
			IDS_RESSUBRES,
			szTabs.c_str(),
			TagToString(lpRes->res.resSub.ulSubObject, lpObj, false, true).c_str()));
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

		resString.push_back(strings::formatmessage(
			IDS_RESCOMMENTRES,
			szTabs.c_str()));
		resString.push_back(RestrictionToString(lpRes->res.resComment.lpRes, lpObj, ulTabLevel + 1));
		break;
	case RES_ANNOTATION:
		resString.push_back(strings::formatmessage(IDS_RESANNOTATION, szTabs.c_str(), lpRes->res.resComment.cValues));
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

		resString.push_back(strings::formatmessage(
			IDS_RESANNOTATIONRES,
			szTabs.c_str()));
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
		SBinary sBinStore = { 0 };
		SBinary sBinFld = { 0 };
		sBinStore.cb = action.actMoveCopy.cbStoreEntryId;
		sBinStore.lpb = reinterpret_cast<LPBYTE>(action.actMoveCopy.lpStoreEntryId);
		sBinFld.cb = action.actMoveCopy.cbFldEntryId;
		sBinFld.lpb = reinterpret_cast<LPBYTE>(action.actMoveCopy.lpFldEntryId);

		actstring += strings::formatmessage(IDS_ACTIONOPMOVECOPY,
			strings::BinToHexString(&sBinStore, true).c_str(),
			strings::BinToTextString(&sBinStore, false).c_str(),
			strings::BinToHexString(&sBinFld, true).c_str(),
			strings::BinToTextString(&sBinFld, false).c_str());
		break;
	}
	case OP_REPLY:
	case OP_OOF_REPLY:
	{

		SBinary sBin = { 0 };
		sBin.cb = action.actReply.cbEntryId;
		sBin.lpb = reinterpret_cast<LPBYTE>(action.actReply.lpEntryId);
		auto szGUID = guid::GUIDToStringAndName(&action.actReply.guidReplyTemplate);

		actstring += strings::formatmessage(IDS_ACTIONOPREPLY,
			strings::BinToHexString(&sBin, true).c_str(),
			strings::BinToTextString(&sBin, false).c_str(),
			szGUID.c_str());
		break;
	}
	case OP_DEFER_ACTION:
	{
		SBinary sBin = { 0 };
		sBin.cb = action.actDeferAction.cbData;
		sBin.lpb = static_cast<LPBYTE>(action.actDeferAction.pbData);

		actstring += strings::formatmessage(IDS_ACTIONOPDEFER,
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
		actstring += strings::formatmessage(IDS_ACTIONOPTAG,
			TagToString(action.propTag.ulPropTag, nullptr, false, true).c_str(),
			szProp.c_str(),
			szAltProp.c_str());
		break;
	}
	default: break;
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
	default: break;
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
			actstring += strings::formatmessage(IDS_ACTIONTAGARRAYTAG,
				i,
				TagToString(action.lpPropTagArray->aulPropTag[i], nullptr, false, false).c_str());
		}
	}

	return actstring;
}

std::wstring ActionsToString(_In_ const ACTIONS& actions)
{
	auto szFlags = InterpretFlags(flagRulesVersion, actions.ulVersion);
	auto actstring = strings::formatmessage(IDS_ACTIONSMEMBERS,
		actions.ulVersion,
		szFlags.c_str(),
		actions.cActions);

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
void InterpretProp(_In_ const _SPropValue* lpProp, _In_opt_ std::wstring* PropString, _In_opt_ std::wstring* AltPropString)
{
	if (!lpProp) return;

	auto parsedProperty = ParseProperty(lpProp);

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

	if (!bTypeFound)
		tmpPropType = strings::format(L"0x%04x", PROP_TYPE(ulPropTag)); // STRING_OK

	if (bNeedInstance) tmpPropType += L" | MV_INSTANCE"; // STRING_OK
	return tmpPropType;
}

// TagToString will prepend the http://schemas.microsoft.com/MAPI/ for us since it's a constant
// We don't compute a DASL string for non-named props as FormatMessage in TagToString can handle those
NamePropNames NameIDToStrings(_In_ LPMAPINAMEID lpNameID, ULONG ulPropTag)
{
	auto hRes = S_OK;
	NamePropNames namePropNames;

	// Can't generate strings without a MAPINAMEID structure
	if (!lpNameID) return namePropNames;

	LPNAMEDPROPCACHEENTRY lpNamedPropCacheEntry = nullptr;

	// If we're using the cache, look up the answer there and return
	if (fCacheNamedProps())
	{
		lpNamedPropCacheEntry = FindCacheEntry(PROP_ID(ulPropTag), lpNameID->lpguid, lpNameID->ulKind, lpNameID->Kind.lID, lpNameID->Kind.lpwstrName);
		if (lpNamedPropCacheEntry && lpNamedPropCacheEntry->bStringsCached)
		{
			namePropNames = lpNamedPropCacheEntry->namePropNames;
			return namePropNames;
		}

		// We shouldn't ever get here without a cached entry
		if (!lpNamedPropCacheEntry)
		{
			DebugPrint(DBGNamedProp, L"NameIDToStrings: Failed to find cache entry for ulPropTag = 0x%08X\n", ulPropTag);
			return namePropNames;
		}
	}

	DebugPrint(DBGNamedProp, L"Parsing named property\n");
	DebugPrint(DBGNamedProp, L"ulPropTag = 0x%08x\n", ulPropTag);
	namePropNames.guid = guid::GUIDToStringAndName(lpNameID->lpguid);
	DebugPrint(DBGNamedProp, L"lpNameID->lpguid = %ws\n", namePropNames.guid.c_str());

	auto szDASLGuid = guid::GUIDToString(lpNameID->lpguid);

	if (lpNameID->ulKind == MNID_ID)
	{
		DebugPrint(DBGNamedProp, L"lpNameID->Kind.lID = 0x%04X = %d\n", lpNameID->Kind.lID, lpNameID->Kind.lID);
		auto pidlids = NameIDToPropNames(lpNameID);

		if (!pidlids.empty())
		{
			namePropNames.bestPidLid = pidlids.front();
			pidlids.erase(pidlids.begin());
			namePropNames.otherPidLid = strings::join(pidlids, L", ");
			// Printing hex first gets a nice sort without spacing tricks
			namePropNames.name = strings::format(L"id: 0x%04X=%d = %ws", // STRING_OK
				lpNameID->Kind.lID,
				lpNameID->Kind.lID,
				namePropNames.bestPidLid.c_str());

			if (!namePropNames.otherPidLid.empty())
			{
				namePropNames.name += strings::format(L" (%ws)", namePropNames.otherPidLid.c_str());
			}
		}
		else
		{
			// Printing hex first gets a nice sort without spacing tricks
			namePropNames.name = strings::format(L"id: 0x%04X=%d", // STRING_OK
				lpNameID->Kind.lID,
				lpNameID->Kind.lID);
		}

		namePropNames.dasl = strings::format(L"id/%s/%04X%04X", // STRING_OK
			szDASLGuid.c_str(),
			lpNameID->Kind.lID,
			PROP_TYPE(ulPropTag));
	}
	else if (lpNameID->ulKind == MNID_STRING)
	{
		// lpwstrName is LPWSTR which means it's ALWAYS unicode
		// But some folks get it wrong and stuff ANSI data in there
		// So we check the string length both ways to make our best guess
		size_t cchShortLen = NULL;
		size_t cchWideLen = NULL;
		WC_H(StringCchLengthA(reinterpret_cast<LPSTR>(lpNameID->Kind.lpwstrName), STRSAFE_MAX_CCH, &cchShortLen));
		WC_H(StringCchLengthW(lpNameID->Kind.lpwstrName, STRSAFE_MAX_CCH, &cchWideLen));

		if (cchShortLen < cchWideLen)
		{
			// this is the *proper* case
			DebugPrint(DBGNamedProp, L"lpNameID->Kind.lpwstrName = \"%ws\"\n", lpNameID->Kind.lpwstrName);
			namePropNames.name = lpNameID->Kind.lpwstrName;

			namePropNames.dasl = strings::format(L"string/%ws/%ws", // STRING_OK
				szDASLGuid.c_str(),
				lpNameID->Kind.lpwstrName);
		}
		else
		{
			// this is the case where ANSI data was shoved into a unicode string.
			DebugPrint(DBGNamedProp, L"Warning: ANSI data was found in a unicode field. This is a bug on the part of the creator of this named property\n");
			DebugPrint(DBGNamedProp, L"lpNameID->Kind.lpwstrName = \"%hs\"\n", reinterpret_cast<LPCSTR>(lpNameID->Kind.lpwstrName));

			auto szComment = strings::loadstring(IDS_NAMEWASANSI);
			namePropNames.name = strings::format(L"%hs %ws", reinterpret_cast<LPSTR>(lpNameID->Kind.lpwstrName), szComment.c_str());

			namePropNames.dasl = strings::format(L"string/%ws/%hs", // STRING_OK
				szDASLGuid.c_str(),
				LPSTR(lpNameID->Kind.lpwstrName));
		}
	}

	// We've built our strings - if we're caching, put them in the cache
	if (lpNamedPropCacheEntry)
	{
		lpNamedPropCacheEntry->namePropNames = namePropNames;
		lpNamedPropCacheEntry->bStringsCached = true;
	}

	return namePropNames;
}

NamePropNames NameIDToStrings(
	ULONG ulPropTag, // optional 'original' prop tag
	_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
	_In_opt_ LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
	_In_opt_ const _SBinary* lpMappingSignature, // optional mapping signature for object to speed named prop lookups
	bool bIsAB) // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
{
	auto hRes = S_OK;
	NamePropNames namePropNames;

	// Named Props
	LPMAPINAMEID* lppPropNames = nullptr;

	// If we weren't passed named property information and we need it, look it up
	// We check bIsAB here - some address book providers return garbage which will crash us
	if (!lpNameID &&
		lpMAPIProp && // if we have an object
		!bIsAB &&
		RegKeys[regkeyPARSED_NAMED_PROPS].ulCurDWORD && // and we're parsing named props
		(RegKeys[regkeyGETPROPNAMES_ON_ALL_PROPS].ulCurDWORD || PROP_ID(ulPropTag) >= 0x8000)) // and it's either a named prop or we're doing all props
	{
		SPropTagArray tag = { 0 };
		auto lpTag = &tag;
		ULONG ulPropNames = 0;
		tag.cValues = 1;
		tag.aulPropTag[0] = ulPropTag;

		WC_H_GETPROPS(GetNamesFromIDs(lpMAPIProp,
			lpMappingSignature,
			&lpTag,
			nullptr,
			NULL,
			&ulPropNames,
			&lppPropNames));
		if (SUCCEEDED(hRes) && ulPropNames == 1 && lppPropNames && lppPropNames[0])
		{
			lpNameID = lppPropNames[0];
		}
	}

	if (lpNameID)
	{
		namePropNames = NameIDToStrings(lpNameID, ulPropTag);
	}

	// Avoid making the call if we don't have to so we don't accidently depend on MAPI
	if (lppPropNames) MAPIFreeBuffer(lppPropNames);

	return namePropNames;
}