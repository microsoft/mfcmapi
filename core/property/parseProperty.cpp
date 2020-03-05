#include <core/stdafx.h>
#include <core/property/parseProperty.h>
#include <core/property/property.h>
#include <core/mapi/extraPropTags.h>
#include <core/interpret/guid.h>
#include <core/utility/strings.h>
#include <core/utility/error.h>
#include <core/interpret/proptags.h>
#include <core/interpret/flags.h>
#include <core/smartview/SmartView.h>
#include <core/mapi/mapiFunctions.h>

namespace property
{
	std::wstring BuildErrorPropString(_In_ const _SPropValue* lpProp)
	{
		if (PROP_TYPE(lpProp->ulPropTag) != PT_ERROR) return L"";
		switch (PROP_ID(lpProp->ulPropTag))
		{
		case PROP_ID(PR_BODY):
		case PROP_ID(PR_BODY_HTML):
		case PROP_ID(PR_RTF_COMPRESSED):
			if (lpProp->Value.err == MAPI_E_NOT_ENOUGH_MEMORY || lpProp->Value.err == MAPI_E_NOT_FOUND)
			{
				return strings::loadstring(IDS_OPENBODY);
			}

			break;
		default:
			if (lpProp->Value.err == MAPI_E_NOT_ENOUGH_MEMORY)
			{
				return strings::loadstring(IDS_OPENSTREAM);
			}
		}

		return L"";
	}

	/***************************************************************************
	Name: parseProperty
	Purpose: Evaluate a property value and return a string representing the property.
	Parameters:
	In:
	LPSPropValue lpProp: Property to be evaluated
	Out:
	wstring* tmpPropString: String representing property value
	wstring* tmpAltPropString: Alternative string representation
	***************************************************************************/
	void parseProperty(
		_In_ const _SPropValue* lpProp,
		_In_opt_ std::wstring* PropString,
		_In_opt_ std::wstring* AltPropString)
	{
		if (!lpProp) return;

		auto parsedProperty = parseProperty(lpProp);

		if (PropString) *PropString = parsedProperty.toString();
		if (AltPropString) *AltPropString = parsedProperty.toAltString();
	}

	Property parseMVProperty(_In_ const _SPropValue* lpProp, ULONG ulMVRow)
	{
		if (!lpProp || ulMVRow > lpProp->Value.MVi.cValues) return Property();

		// We'll let parseProperty do all the work
		SPropValue sProp = {};
		sProp.ulPropTag = CHANGE_PROP_TYPE(lpProp->ulPropTag, PROP_TYPE(lpProp->ulPropTag) & ~MV_FLAG);

		// Only attempt to dereference our array if it's non-NULL
		if (PROP_TYPE(lpProp->ulPropTag) & MV_FLAG && lpProp->Value.MVi.lpi)
		{
			switch (PROP_TYPE(lpProp->ulPropTag))
			{
			case PT_MV_I2:
				sProp.Value.i = lpProp->Value.MVi.lpi[ulMVRow];
				break;
			case PT_MV_LONG:
				sProp.Value.l = lpProp->Value.MVl.lpl[ulMVRow];
				break;
			case PT_MV_DOUBLE:
				sProp.Value.dbl = lpProp->Value.MVdbl.lpdbl[ulMVRow];
				break;
			case PT_MV_CURRENCY:
				sProp.Value.cur = lpProp->Value.MVcur.lpcur[ulMVRow];
				break;
			case PT_MV_APPTIME:
				sProp.Value.at = lpProp->Value.MVat.lpat[ulMVRow];
				break;
			case PT_MV_SYSTIME:
				sProp.Value.ft = lpProp->Value.MVft.lpft[ulMVRow];
				break;
			case PT_MV_I8:
				sProp.Value.li = lpProp->Value.MVli.lpli[ulMVRow];
				break;
			case PT_MV_R4:
				sProp.Value.flt = lpProp->Value.MVflt.lpflt[ulMVRow];
				break;
			case PT_MV_STRING8:
				sProp.Value.lpszA = lpProp->Value.MVszA.lppszA[ulMVRow];
				break;
			case PT_MV_UNICODE:
				sProp.Value.lpszW = lpProp->Value.MVszW.lppszW[ulMVRow];
				break;
			case PT_MV_BINARY:
				mapi::setBin(sProp) = lpProp->Value.MVbin.lpbin[ulMVRow];
				break;
			case PT_MV_CLSID:
				sProp.Value.lpguid = &lpProp->Value.MVguid.lpguid[ulMVRow];
				break;
			default:
				break;
			}
		}

		return parseProperty(&sProp);
	}

	Property parseProperty(_In_ const _SPropValue* lpProp)
	{
		Property properties;
		if (!lpProp) return properties;

		if (MV_FLAG & PROP_TYPE(lpProp->ulPropTag))
		{
			// MV property
			properties.AddAttribute(L"mv", L"true"); // STRING_OK
			// All the MV structures are basically the same, so we can cheat when we pull the count
			properties.AddAttribute(L"count", std::to_wstring(lpProp->Value.MVi.cValues)); // STRING_OK

			// Don't bother with the loop if we don't have data
			if (lpProp->Value.MVi.lpi)
			{
				for (ULONG iMVCount = 0; iMVCount < lpProp->Value.MVi.cValues; iMVCount++)
				{
					properties.AddMVParsing(parseMVProperty(lpProp, iMVCount));
				}
			}
		}
		else
		{
			std::wstring szTmp;
			auto bPropXMLSafe = true;
			Attributes attributes;

			std::wstring szAltTmp;
			auto bAltPropXMLSafe = true;
			Attributes altAttributes;

			switch (PROP_TYPE(lpProp->ulPropTag))
			{
			case PT_I2:
				szTmp = std::to_wstring(lpProp->Value.i);
				szAltTmp = strings::format(L"0x%X", lpProp->Value.i); // STRING_OK
				break;
			case PT_LONG:
				szTmp = std::to_wstring(lpProp->Value.l);
				szAltTmp = strings::format(L"0x%X", lpProp->Value.l); // STRING_OK
				break;
			case PT_R4:
				szTmp = std::to_wstring(lpProp->Value.flt); // STRING_OK
				break;
			case PT_DOUBLE:
				szTmp = std::to_wstring(lpProp->Value.dbl); // STRING_OK
				break;
			case PT_CURRENCY:
				szTmp = strings::format(L"%05I64d", lpProp->Value.cur.int64); // STRING_OK
				if (szTmp.length() > 4)
				{
					szTmp.insert(szTmp.length() - 4, L".");
				}

				szAltTmp = strings::format(
					L"0x%08X:0x%08X",
					static_cast<int>(lpProp->Value.cur.Hi),
					static_cast<int>(lpProp->Value.cur.Lo)); // STRING_OK
				break;
			case PT_APPTIME:
				szTmp = std::to_wstring(lpProp->Value.at); // STRING_OK
				break;
			case PT_ERROR:
				szTmp = error::ErrorNameFromErrorCode(lpProp->Value.err); // STRING_OK
				szAltTmp = BuildErrorPropString(lpProp);

				attributes.AddAttribute(L"err", strings::format(L"0x%08X", lpProp->Value.err)); // STRING_OK
				break;
			case PT_BOOLEAN:
				szTmp = strings::loadstring(lpProp->Value.b ? IDS_TRUE : IDS_FALSE);
				break;
			case PT_OBJECT:
				szTmp = strings::loadstring(IDS_OBJECT);
				break;
			case PT_I8: // LARGE_INTEGER
				szTmp = strings::format(
					L"0x%08X:0x%08X",
					static_cast<int>(lpProp->Value.li.HighPart),
					static_cast<int>(lpProp->Value.li.LowPart)); // STRING_OK
				szAltTmp = strings::format(L"%I64d", lpProp->Value.li.QuadPart); // STRING_OK
				break;
			case PT_STRING8:
				if (strings::CheckStringProp(lpProp, PT_STRING8))
				{
					szTmp = strings::LPCSTRToWstring(lpProp->Value.lpszA);
					bPropXMLSafe = false;

					SBinary sBin = {};
					sBin.cb = static_cast<ULONG>(szTmp.length());
					sBin.lpb = reinterpret_cast<LPBYTE>(lpProp->Value.lpszA);
					szAltTmp = strings::BinToHexString(&sBin, false);

					altAttributes.AddAttribute(L"cb", std::to_wstring(sBin.cb)); // STRING_OK
				}
				break;
			case PT_UNICODE:
				if (strings::CheckStringProp(lpProp, PT_UNICODE))
				{
					szTmp = lpProp->Value.lpszW;
					bPropXMLSafe = false;

					SBinary sBin = {};
					sBin.cb = static_cast<ULONG>(szTmp.length()) * sizeof(WCHAR);
					sBin.lpb = reinterpret_cast<LPBYTE>(lpProp->Value.lpszW);
					szAltTmp = strings::BinToHexString(&sBin, false);

					altAttributes.AddAttribute(L"cb", std::to_wstring(sBin.cb)); // STRING_OK
				}
				break;
			case PT_SYSTIME:
				strings::FileTimeToString(lpProp->Value.ft, szTmp, szAltTmp);
				break;
			case PT_CLSID:
				// TODO: One string matches current behavior - look at splitting to two strings in future change
				szTmp = guid::GUIDToStringAndName(lpProp->Value.lpguid);
				break;
			case PT_BINARY:
			{
				const auto bin = mapi::getBin(lpProp);
				szTmp = strings::BinToHexString(&bin, false);
				szAltTmp = strings::BinToTextString(&bin, false);
				bAltPropXMLSafe = false;

				attributes.AddAttribute(L"cb", std::to_wstring(bin.cb)); // STRING_OK
				break;
			}
			case PT_SRESTRICTION:
				szTmp = RestrictionToString(reinterpret_cast<LPSRestriction>(lpProp->Value.lpszA), nullptr);
				bPropXMLSafe = false;
				break;
			case PT_ACTIONS:
				if (lpProp->Value.lpszA)
				{
					const auto actions = reinterpret_cast<const ACTIONS*>(lpProp->Value.lpszA);
					szTmp = ActionsToString(*actions);
				}
				else
				{
					szTmp = strings::loadstring(IDS_ACTIONSNULL);
				}

				bPropXMLSafe = false;

				break;
			default:
				break;
			}

			const Parsing mainParsing(szTmp, bPropXMLSafe, attributes);
			const Parsing altParsing(szAltTmp, bAltPropXMLSafe, altAttributes);
			properties.AddParsing(mainParsing, altParsing);
		}

		return properties;
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
			resString.push_back(strings::formatmessage(IDS_RESCOUNT, szTabs.c_str(), lpRes->res.resCount.ulCount));
			resString.push_back(RestrictionToString(lpRes->res.resCount.lpRes, lpObj, ulTabLevel + 1));
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
				parseProperty(lpRes->res.resContent.lpProp, &szProp, &szAltProp);
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
				parseProperty(lpRes->res.resProperty.lpProp, &szProp, &szAltProp);
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
		{
			szFlags = flags::InterpretFlags(flagBitmask, lpRes->res.resBitMask.relBMR);
			auto mask = strings::formatmessage(
				IDS_RESBITMASK,
				szTabs.c_str(),
				szFlags.c_str(),
				lpRes->res.resBitMask.relBMR,
				lpRes->res.resBitMask.ulMask);
			szPropNum =
				smartview::InterpretNumberAsStringProp(lpRes->res.resBitMask.ulMask, lpRes->res.resBitMask.ulPropTag);
			if (!szPropNum.empty())
			{
				mask += strings::formatmessage(IDS_RESBITMASKFLAGS, szPropNum.c_str());
			}

			resString.push_back(mask);

			resString.push_back(strings::formatmessage(
				IDS_RESBITMASKTAG,
				szTabs.c_str(),
				proptags::TagToString(lpRes->res.resBitMask.ulPropTag, lpObj, false, true).c_str()));
		}

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
					parseProperty(&lpRes->res.resComment.lpProp[i], &szProp, &szAltProp);
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
					parseProperty(&lpRes->res.resComment.lpProp[i], &szProp, &szAltProp);
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
				parseProperty(&adrList.aEntries[i].rgPropVals[j], &szProp, &szAltProp);
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
			auto sBin = SBinary{action.actDeferAction.cbData, action.actDeferAction.pbData};

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
			parseProperty(const_cast<LPSPropValue>(&action.propTag), &szProp, &szAltProp);
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
					proptags::TagToString(mapi::getTag(action.lpPropTagArray, i), nullptr, false, false).c_str());
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
} // namespace property