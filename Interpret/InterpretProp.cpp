#include <StdAfx.h>
#include <Interpret/InterpretProp.h>
#include <core/interpret/guid.h>
#include <core/mapi/extraPropTags.h>
#include <Interpret/SmartView/SmartView.h>
#include <Property/ParseProperty.h>
#include <core/utility/strings.h>
#include <core/interpret/flags.h>
#include <core/interpret/proptags.h>

namespace interpretprop
{
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
		auto szFlags = flags::InterpretFlags(flagRestrictionType, lpRes->rt);
		resString.push_back(strings::formatmessage(IDS_RESTYPE, szTabs.c_str(), lpRes->rt, szFlags.c_str()));

		switch (lpRes->rt)
		{
		case RES_COMPAREPROPS:
			szFlags = flags::InterpretFlags(flagRelop, lpRes->res.resCompareProps.relop);
			resString.push_back(strings::formatmessage(
				IDS_RESCOMPARE,
				szTabs.c_str(),
				szFlags.c_str(),
				lpRes->res.resCompareProps.relop,
				proptags::TagToString(lpRes->res.resCompareProps.ulPropTag1, lpObj, false, true).c_str(),
				proptags::TagToString(lpRes->res.resCompareProps.ulPropTag2, lpObj, false, true).c_str()));
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
			szFlags = flags::InterpretFlags(flagFuzzyLevel, lpRes->res.resContent.ulFuzzyLevel);
			resString.push_back(strings::formatmessage(
				IDS_RESCONTENT,
				szTabs.c_str(),
				szFlags.c_str(),
				lpRes->res.resContent.ulFuzzyLevel,
				proptags::TagToString(lpRes->res.resContent.ulPropTag, lpObj, false, true).c_str()));
			if (lpRes->res.resContent.lpProp)
			{
				InterpretProp(lpRes->res.resContent.lpProp, &szProp, &szAltProp);
				resString.push_back(strings::formatmessage(
					IDS_RESCONTENTPROP,
					szTabs.c_str(),
					proptags::TagToString(lpRes->res.resContent.lpProp->ulPropTag, lpObj, false, true).c_str(),
					szProp.c_str(),
					szAltProp.c_str()));
			}
			break;
		case RES_PROPERTY:
			szFlags = flags::InterpretFlags(flagRelop, lpRes->res.resProperty.relop);
			resString.push_back(strings::formatmessage(
				IDS_RESPROP,
				szTabs.c_str(),
				szFlags.c_str(),
				lpRes->res.resProperty.relop,
				proptags::TagToString(lpRes->res.resProperty.ulPropTag, lpObj, false, true).c_str()));
			if (lpRes->res.resProperty.lpProp)
			{
				InterpretProp(lpRes->res.resProperty.lpProp, &szProp, &szAltProp);
				resString.push_back(strings::formatmessage(
					IDS_RESPROPPROP,
					szTabs.c_str(),
					proptags::TagToString(lpRes->res.resProperty.lpProp->ulPropTag, lpObj, false, true).c_str(),
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
			szFlags = flags::InterpretFlags(flagBitmask, lpRes->res.resBitMask.relBMR);
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
				proptags::TagToString(lpRes->res.resBitMask.ulPropTag, lpObj, false, true).c_str()));
			break;
		case RES_SIZE:
			szFlags = flags::InterpretFlags(flagRelop, lpRes->res.resSize.relop);
			resString.push_back(strings::formatmessage(
				IDS_RESSIZE,
				szTabs.c_str(),
				szFlags.c_str(),
				lpRes->res.resSize.relop,
				lpRes->res.resSize.cb,
				proptags::TagToString(lpRes->res.resSize.ulPropTag, lpObj, false, true).c_str()));
			break;
		case RES_EXIST:
			resString.push_back(strings::formatmessage(
				IDS_RESEXIST,
				szTabs.c_str(),
				proptags::TagToString(lpRes->res.resExist.ulPropTag, lpObj, false, true).c_str(),
				lpRes->res.resExist.ulReserved1,
				lpRes->res.resExist.ulReserved2));
			break;
		case RES_SUBRESTRICTION:
			resString.push_back(strings::formatmessage(
				IDS_RESSUBRES,
				szTabs.c_str(),
				proptags::TagToString(lpRes->res.resSub.ulSubObject, lpObj, false, true).c_str()));
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
						proptags::TagToString(lpRes->res.resComment.lpProp[i].ulPropTag, lpObj, false, true).c_str(),
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
						proptags::TagToString(lpRes->res.resComment.lpProp[i].ulPropTag, lpObj, false, true).c_str(),
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
					proptags::TagToString(adrList.aEntries[i].rgPropVals[j].ulPropTag, nullptr, false, false).c_str(),
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
		auto szFlags = flags::InterpretFlags(flagAccountType, action.acttype);
		auto szFlags2 = flags::InterpretFlags(flagRuleFlag, action.ulFlags);
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
			szFlags = flags::InterpretFlags(flagBounceCode, action.scBounceCode);
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
				proptags::TagToString(action.propTag.ulPropTag, nullptr, false, true).c_str(),
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
			szFlags = flags::InterpretFlags(flagOPReply, action.ulActionFlavor);
			break;
		}
		case OP_FORWARD:
		{
			szFlags = flags::InterpretFlags(flagOpForward, action.ulActionFlavor);
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
					proptags::TagToString(action.lpPropTagArray->aulPropTag[i], nullptr, false, false).c_str());
			}
		}

		return actstring;
	}

	std::wstring ActionsToString(_In_ const ACTIONS& actions)
	{
		auto szFlags = flags::InterpretFlags(flagRulesVersion, actions.ulVersion);
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
} // namespace interpretprop