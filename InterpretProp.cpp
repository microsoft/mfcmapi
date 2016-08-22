#include "stdafx.h"
#include "InterpretProp.h"
#include "MAPIFunctions.h"
#include "InterpretProp2.h"
#include "ExtraPropTags.h"
#include "NamedPropCache.h"
#include "SmartView\SmartView.h"
#include "ParseProperty.h"
#include "String.h"
#include <vector>

static const char pBase64[] = {
 0x3e, 0x7f, 0x7f, 0x7f, 0x3f, 0x34, 0x35, 0x36,
 0x37, 0x38, 0x39, 0x3a, 0x3b, 0x3c, 0x3d, 0x7f,
 0x7f, 0x7f, 0x7f, 0x7f, 0x7f, 0x7f, 0x00, 0x01,
 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09,
 0x0a, 0x0b, 0x0c, 0x0d, 0x0e, 0x0f, 0x10, 0x11,
 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19,
 0x7f, 0x7f, 0x7f, 0x7f, 0x7f, 0x7f, 0x1a, 0x1b,
 0x1c, 0x1d, 0x1e, 0x1f, 0x20, 0x21, 0x22, 0x23,
 0x24, 0x25, 0x26, 0x27, 0x28, 0x29, 0x2a, 0x2b,
 0x2c, 0x2d, 0x2e, 0x2f, 0x30, 0x31, 0x32, 0x33
};

// allocates output buffer with new
// delete with delete[]
// suprisingly, this algorithm works in a unicode build as well
vector<BYTE> Base64Decode(wstring szEncodedStr)
{
	HRESULT hRes = S_OK;
	LPCWSTR szEncodedStrPtr = szEncodedStr.c_str();
	size_t cchLen = szEncodedStr.length();
	vector<BYTE> lpb;
	if (cchLen % 4) return lpb;

	// look for padding at the end
	static const WCHAR szPadding[] = L"=="; // STRING_OK
	const WCHAR* szPaddingLoc = wcschr(szEncodedStrPtr, szPadding[0]);
	if (NULL != szPaddingLoc)
	{
		size_t cchPaddingLen = 0;
		// check padding length
		EC_H(StringCchLength(szPaddingLoc, STRSAFE_MAX_CCH, &cchPaddingLen));
		if (cchPaddingLen >= 3) return lpb;

		// check for bad characters after the first '='
		if (wcsncmp(szPaddingLoc, (WCHAR*)szPadding, cchPaddingLen)) return lpb;

		// cchPaddingLen == 0,1,2 now
	}

	WCHAR c[4] = { 0 };
	BYTE bTmp[3] = { 0 }; // output

	while (*szEncodedStrPtr)
	{
		int i = 0;
		int iOutlen = 3;
		for (i = 0; i < 4; i++)
		{
			c[i] = *(szEncodedStrPtr + i);
			if (c[i] == L'=')
			{
				iOutlen = i - 1;
				break;
			}

			if ((c[i] < 0x2b) || (c[i] > 0x7a)) return vector<BYTE>();

			c[i] = pBase64[c[i] - 0x2b];
		}

		bTmp[0] = (BYTE)((c[0] << 2) | (c[1] >> 4));
		bTmp[1] = (BYTE)((c[1] & 0x0f) << 4 | (c[2] >> 2));
		bTmp[2] = (BYTE)((c[2] & 0x03) << 6 | c[3]);

		for (i = 0; i < iOutlen; i++)
		{
			lpb.push_back(bTmp[i]);
		}

		szEncodedStrPtr += 4;
	}

	return lpb;
}

static const // Base64 Index into encoding
char pIndex[] = { // and decoding table.
 0x41, 0x42, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48,
 0x49, 0x4a, 0x4b, 0x4c, 0x4d, 0x4e, 0x4f, 0x50,
 0x51, 0x52, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58,
 0x59, 0x5a, 0x61, 0x62, 0x63, 0x64, 0x65, 0x66,
 0x67, 0x68, 0x69, 0x6a, 0x6b, 0x6c, 0x6d, 0x6e,
 0x6f, 0x70, 0x71, 0x72, 0x73, 0x74, 0x75, 0x76,
 0x77, 0x78, 0x79, 0x7a, 0x30, 0x31, 0x32, 0x33,
 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x2b, 0x2f
};

wstring Base64Encode(size_t cbSourceBuf, _In_count_(cbSourceBuf) LPBYTE lpSourceBuffer)
{
	wstring szEncodedStr;
	size_t cbBuf = 0;

	// Using integer division to round down here
	while (cbBuf < (cbSourceBuf / 3) * 3) // encode each 3 byte octet.
	{
		szEncodedStr += pIndex[lpSourceBuffer[cbBuf] >> 2];
		szEncodedStr += pIndex[((lpSourceBuffer[cbBuf] & 0x03) << 4) + (lpSourceBuffer[cbBuf + 1] >> 4)];
		szEncodedStr += pIndex[((lpSourceBuffer[cbBuf + 1] & 0x0f) << 2) + (lpSourceBuffer[cbBuf + 2] >> 6)];
		szEncodedStr += pIndex[lpSourceBuffer[cbBuf + 2] & 0x3f];
		cbBuf += 3; // Next octet.
	}

	if (cbSourceBuf - cbBuf != 0) // Partial octet remaining?
	{
		szEncodedStr += pIndex[lpSourceBuffer[cbBuf] >> 2]; // Yes, encode it.

		if (cbSourceBuf - cbBuf == 1) // End of octet?
		{
			szEncodedStr += pIndex[(lpSourceBuffer[cbBuf] & 0x03) << 4];
			szEncodedStr += L'=';
			szEncodedStr += L'=';
		}
		else
		{ // No, one more part.
			szEncodedStr += pIndex[((lpSourceBuffer[cbBuf] & 0x03) << 4) + (lpSourceBuffer[cbBuf + 1] >> 4)];
			szEncodedStr += pIndex[(lpSourceBuffer[cbBuf + 1] & 0x0f) << 2];
			szEncodedStr += L'=';
		}
	}

	return szEncodedStr;
}

wstring CurrencyToString(CURRENCY curVal)
{
	wstring szCur = format(L"%05I64d", curVal.int64); // STRING_OK
	if (szCur.length() > 4)
	{
		szCur.insert(szCur.length() - 4, L"."); // STRING_OK
	}

	return szCur;
}

wstring TagToString(ULONG ulPropTag, _In_opt_ LPMAPIPROP lpObj, bool bIsAB, bool bSingleLine)
{
	wstring szRet;
	wstring szTemp;

	wstring szExactMatches;
	wstring szPartialMatches;
	wstring szNamedPropName;
	wstring szNamedPropGUID;
	wstring szNamedPropDASL;

	NameIDToStrings(
		ulPropTag,
		lpObj,
		NULL,
		NULL,
		bIsAB,
		szNamedPropName, // Built from lpProp & lpMAPIProp
		szNamedPropGUID, // Built from lpProp & lpMAPIProp
		szNamedPropDASL); // Built from ulPropTag & lpMAPIProp

	PropTagToPropName(ulPropTag, bIsAB, &szExactMatches, &szPartialMatches);

	wstring szFormatString;
	if (bSingleLine)
	{
		szFormatString = L"0x%1!08X! (%2)"; // STRING_OK
		if (!szExactMatches.empty()) szFormatString += L": %3!ws!"; // STRING_OK
		if (!szPartialMatches.empty()) szFormatString += L": (%4!ws!)"; // STRING_OK
		if (!szNamedPropName.empty())
		{
			szFormatString += loadstring(IDS_NAMEDPROPSINGLELINE);
		}

		if (!szNamedPropGUID.empty())
		{
			szFormatString += loadstring(IDS_GUIDSINGLELINE);
		}
	}
	else
	{
		szFormatString = loadstring(IDS_TAGMULTILINE);
		if (!szExactMatches.empty())
		{
			szFormatString += loadstring(IDS_PROPNAMEMULTILINE);
		}

		if (!szPartialMatches.empty())
		{
			szFormatString += loadstring(IDS_OTHERNAMESMULTILINE);
		}

		if (PROP_ID(ulPropTag) < 0x8000)
		{
			szFormatString += loadstring(IDS_DASLPROPTAG);
		}
		else if (!szNamedPropDASL.empty())
		{
			szFormatString += loadstring(IDS_DASLNAMED);
		}

		if (!szNamedPropName.empty())
		{
			szFormatString += loadstring(IDS_NAMEPROPNAMEMULTILINE);
		}

		if (!szNamedPropGUID.empty())
		{
			szFormatString += loadstring(IDS_NAMEPROPGUIDMULTILINE);
		}
	}

	szRet = formatmessage(szFormatString,
		ulPropTag,
		TypeToString(ulPropTag).c_str(),
		szExactMatches.c_str(),
		szPartialMatches.c_str(),
		szNamedPropName.c_str(),
		szNamedPropGUID.c_str(),
		szNamedPropDASL.c_str());

	if (fIsSet(DBGTest))
	{
		static size_t cchMaxBuff = 0;
		size_t cchBuff = szRet.length();
		cchMaxBuff = max(cchBuff, cchMaxBuff);
		DebugPrint(DBGTest, L"TagToString parsing 0x%08X returned %u chars - max %u\n", ulPropTag, (UINT)cchBuff, (UINT)cchMaxBuff);
	}

	return szRet;
}

wstring ProblemArrayToString(_In_ LPSPropProblemArray lpProblems)
{
	wstring szOut;
	if (lpProblems)
	{
		ULONG i = 0;
		for (i = 0; i < lpProblems->cProblem; i++)
		{
			wstring szTemp = formatmessage(
				IDS_PROBLEMARRAY,
				lpProblems->aProblem[i].ulIndex,
				TagToString(lpProblems->aProblem[i].ulPropTag, NULL, false, false).c_str(),
				lpProblems->aProblem[i].scode,
				ErrorNameFromErrorCode(lpProblems->aProblem[i].scode).c_str());
			szOut += szTemp;
		}
	}

	return szOut;
}

wstring MAPIErrToString(ULONG ulFlags, _In_ LPMAPIERROR lpErr)
{
	wstring szOut;
	if (lpErr)
	{
		szOut = formatmessage(
			ulFlags & MAPI_UNICODE ? IDS_MAPIERRUNICODE : IDS_MAPIERRANSI,
			lpErr->ulVersion,
			lpErr->lpszError,
			lpErr->lpszComponent,
			lpErr->ulLowLevelError,
			ErrorNameFromErrorCode(lpErr->ulLowLevelError).c_str(),
			lpErr->ulContext);
	}
	return szOut;
}

wstring TnefProblemArrayToString(_In_ LPSTnefProblemArray lpError)
{
	wstring szOut;
	if (lpError)
	{
		for (ULONG iError = 0; iError < lpError->cProblem; iError++)
		{
			wstring szTemp = formatmessage(
				IDS_TNEFPROBARRAY,
				lpError->aProblem[iError].ulComponent,
				lpError->aProblem[iError].ulAttribute,
				TagToString(lpError->aProblem[iError].ulPropTag, NULL, false, false).c_str(),
				lpError->aProblem[iError].scode,
				ErrorNameFromErrorCode(lpError->aProblem[iError].scode).c_str());
			szOut += szTemp;
		}
	}
	return szOut;
}

// There may be restrictions with over 100 nested levels, but we're not going to try to parse them
#define _MaxRestrictionNesting 100

wstring RestrictionToString(_In_ LPSRestriction lpRes, _In_opt_ LPMAPIPROP lpObj, ULONG ulTabLevel)
{
	ULONG i = 0;
	if (!lpRes)
	{
		return loadstring(IDS_NULLRES);
	}
	if (ulTabLevel > _MaxRestrictionNesting)
	{
		return loadstring(IDS_RESDEPTHEXCEEDED);
	}

	wstring resString;
	wstring szProp;
	wstring szAltProp;

	wstring szTabs;
	for (i = 0; i < ulTabLevel; i++)
	{
		szTabs += L"\t"; // STRING_OK
	}

	wstring szPropNum;
	wstring szFlags = InterpretFlags(flagRestrictionType, lpRes->rt);
	resString += formatmessage(IDS_RESTYPE, szTabs.c_str(), lpRes->rt, szFlags.c_str());

	switch (lpRes->rt)
	{
	case RES_COMPAREPROPS:
		szFlags = InterpretFlags(flagRelop, lpRes->res.resCompareProps.relop);
		resString += formatmessage(
			IDS_RESCOMPARE,
			szTabs.c_str(),
			szFlags.c_str(),
			lpRes->res.resCompareProps.relop,
			TagToString(lpRes->res.resCompareProps.ulPropTag1, lpObj, false, true).c_str(),
			TagToString(lpRes->res.resCompareProps.ulPropTag2, lpObj, false, true).c_str());
		break;
	case RES_AND:
		resString += formatmessage(IDS_RESANDCOUNT, szTabs.c_str(), lpRes->res.resAnd.cRes);
		if (lpRes->res.resAnd.lpRes)
		{
			for (i = 0; i < lpRes->res.resAnd.cRes; i++)
			{
				resString += formatmessage(IDS_RESANDPOINTER, szTabs.c_str(), i);
				resString += RestrictionToString(&lpRes->res.resAnd.lpRes[i], lpObj, ulTabLevel + 1);
			}
		}
		break;
	case RES_OR:
		resString += formatmessage(IDS_RESORCOUNT, szTabs.c_str(), lpRes->res.resOr.cRes);
		if (lpRes->res.resOr.lpRes)
		{
			for (i = 0; i < lpRes->res.resOr.cRes; i++)
			{
				resString += formatmessage(IDS_RESORPOINTER, szTabs.c_str(), i);
				resString += RestrictionToString(&lpRes->res.resOr.lpRes[i], lpObj, ulTabLevel + 1);
			}
		}
		break;
	case RES_NOT:
		resString += formatmessage(
			IDS_RESNOT,
			szTabs.c_str(),
			lpRes->res.resNot.ulReserved);
		resString += RestrictionToString(lpRes->res.resNot.lpRes, lpObj, ulTabLevel + 1);
		break;
	case RES_COUNT:
		// RES_COUNT and RES_NOT look the same, so we use the resNot member here
		resString += formatmessage(
			IDS_RESCOUNT,
			szTabs.c_str(),
			lpRes->res.resNot.ulReserved);
		resString += RestrictionToString(lpRes->res.resNot.lpRes, lpObj, ulTabLevel + 1);
		break;
	case RES_CONTENT:
		szFlags = InterpretFlags(flagFuzzyLevel, lpRes->res.resContent.ulFuzzyLevel);
		resString += formatmessage(
			IDS_RESCONTENT,
			szTabs.c_str(),
			szFlags.c_str(),
			lpRes->res.resContent.ulFuzzyLevel,
			TagToString(lpRes->res.resContent.ulPropTag, lpObj, false, true).c_str());
		if (lpRes->res.resContent.lpProp)
		{
			InterpretProp(lpRes->res.resContent.lpProp, &szProp, &szAltProp);
			resString += formatmessage(
				IDS_RESCONTENTPROP,
				szTabs.c_str(),
				TagToString(lpRes->res.resContent.lpProp->ulPropTag, lpObj, false, true).c_str(),
				szProp.c_str(),
				szAltProp.c_str());
		}
		break;
	case RES_PROPERTY:
		szFlags = InterpretFlags(flagRelop, lpRes->res.resProperty.relop);
		resString += formatmessage(
			IDS_RESPROP,
			szTabs.c_str(),
			szFlags.c_str(),
			lpRes->res.resProperty.relop,
			TagToString(lpRes->res.resProperty.ulPropTag, lpObj, false, true).c_str());
		if (lpRes->res.resProperty.lpProp)
		{
			InterpretProp(lpRes->res.resProperty.lpProp, &szProp, &szAltProp);
			resString += formatmessage(
				IDS_RESPROPPROP,
				szTabs.c_str(),
				TagToString(lpRes->res.resProperty.lpProp->ulPropTag, lpObj, false, true).c_str(),
				szProp.c_str(),
				szAltProp.c_str());
			szPropNum = InterpretNumberAsString(lpRes->res.resProperty.lpProp->Value, lpRes->res.resProperty.lpProp->ulPropTag, NULL, NULL, NULL, false);
			if (!szPropNum.empty())
			{
				resString += formatmessage(IDS_RESPROPPROPFLAGS, szTabs.c_str(), szPropNum.c_str());
			}
		}
		break;
	case RES_BITMASK:
		szFlags = InterpretFlags(flagBitmask, lpRes->res.resBitMask.relBMR);
		resString += formatmessage(
			IDS_RESBITMASK,
			szTabs.c_str(),
			szFlags.c_str(),
			lpRes->res.resBitMask.relBMR,
			lpRes->res.resBitMask.ulMask);
		szPropNum = InterpretNumberAsStringProp(lpRes->res.resBitMask.ulMask, lpRes->res.resBitMask.ulPropTag);
		if (!szPropNum.empty())
		{
			resString += formatmessage(IDS_RESBITMASKFLAGS, szPropNum.c_str());
		}
		resString += formatmessage(
			IDS_RESBITMASKTAG,
			szTabs.c_str(),
			TagToString(lpRes->res.resBitMask.ulPropTag, lpObj, false, true).c_str());
		break;
	case RES_SIZE:
		szFlags = InterpretFlags(flagRelop, lpRes->res.resSize.relop);
		resString += formatmessage(
			IDS_RESSIZE,
			szTabs.c_str(),
			szFlags.c_str(),
			lpRes->res.resSize.relop,
			lpRes->res.resSize.cb,
			TagToString(lpRes->res.resSize.ulPropTag, lpObj, false, true).c_str());
		break;
	case RES_EXIST:
		resString += formatmessage(
			IDS_RESEXIST,
			szTabs.c_str(),
			TagToString(lpRes->res.resExist.ulPropTag, lpObj, false, true).c_str(),
			lpRes->res.resExist.ulReserved1,
			lpRes->res.resExist.ulReserved2);
		break;
	case RES_SUBRESTRICTION:
		resString += formatmessage(
			IDS_RESSUBRES,
			szTabs.c_str(),
			TagToString(lpRes->res.resSub.ulSubObject, lpObj, false, true).c_str());
		resString += RestrictionToString(lpRes->res.resSub.lpRes, lpObj, ulTabLevel + 1);
		break;
	case RES_COMMENT:
		resString += formatmessage(IDS_RESCOMMENT, szTabs.c_str(), lpRes->res.resComment.cValues);
		if (lpRes->res.resComment.lpProp)
		{
			for (i = 0; i < lpRes->res.resComment.cValues; i++)
			{
				InterpretProp(&lpRes->res.resComment.lpProp[i], &szProp, &szAltProp);
				resString += formatmessage(
					IDS_RESCOMMENTPROPS,
					szTabs.c_str(),
					i,
					TagToString(lpRes->res.resComment.lpProp[i].ulPropTag, lpObj, false, true).c_str(),
					szProp.c_str(),
					szAltProp.c_str());
			}
		}
		resString += formatmessage(
			IDS_RESCOMMENTRES,
			szTabs.c_str());
		resString += RestrictionToString(lpRes->res.resComment.lpRes, lpObj, ulTabLevel + 1);
		break;
	case RES_ANNOTATION:
		resString += formatmessage(IDS_RESANNOTATION, szTabs.c_str(), lpRes->res.resComment.cValues);
		if (lpRes->res.resComment.lpProp)
		{
			for (i = 0; i < lpRes->res.resComment.cValues; i++)
			{
				InterpretProp(&lpRes->res.resComment.lpProp[i], &szProp, &szAltProp);
				resString += formatmessage(
					IDS_RESANNOTATIONPROPS,
					szTabs.c_str(),
					i,
					TagToString(lpRes->res.resComment.lpProp[i].ulPropTag, lpObj, false, true).c_str(),
					szProp.c_str(),
					szAltProp.c_str());
			}
		}

		resString += formatmessage(
			IDS_RESANNOTATIONRES,
			szTabs.c_str());
		resString += RestrictionToString(lpRes->res.resComment.lpRes, lpObj, ulTabLevel + 1);
		break;
	}

	return resString;
}

wstring RestrictionToString(_In_ LPSRestriction lpRes, _In_opt_ LPMAPIPROP lpObj)
{
	return RestrictionToString(lpRes, lpObj, 0);
}

wstring AdrListToString(_In_ LPADRLIST lpAdrList)
{
	if (!lpAdrList)
	{
		return loadstring(IDS_ADRLISTNULL);
	}

	wstring adrstring;
	wstring szProp;
	wstring szAltProp;
	adrstring = formatmessage(IDS_ADRLISTCOUNT, lpAdrList->cEntries);

	ULONG i = 0;
	for (i = 0; i < lpAdrList->cEntries; i++)
	{
		adrstring += formatmessage(IDS_ADRLISTENTRIESCOUNT, i, lpAdrList->aEntries[i].cValues);

		ULONG j = 0;
		for (j = 0; j < lpAdrList->aEntries[i].cValues; j++)
		{
			InterpretProp(&lpAdrList->aEntries[i].rgPropVals[j], &szProp, &szAltProp);
			adrstring += formatmessage(
				IDS_ADRLISTENTRY,
				i,
				j,
				TagToString(lpAdrList->aEntries[i].rgPropVals[j].ulPropTag, NULL, false, false).c_str(),
				szProp.c_str(),
				szAltProp.c_str());
		}
	}

	return adrstring;
}

_Check_return_ wstring ActionToString(_In_ ACTION* lpAction)
{
	if (!lpAction)
	{
		return loadstring(IDS_ACTIONNULL);
	}

	wstring actstring;
	wstring szProp;
	wstring szAltProp;
	wstring szFlags = InterpretFlags(flagAccountType, lpAction->acttype);
	wstring szFlags2 = InterpretFlags(flagRuleFlag, lpAction->ulFlags);
	actstring = formatmessage(
		IDS_ACTION,
		lpAction->acttype,
		szFlags.c_str(),
		RestrictionToString(lpAction->lpRes, NULL).c_str(),
		lpAction->ulFlags,
		szFlags2.c_str());

	switch (lpAction->acttype)
	{
	case OP_MOVE:
	case OP_COPY:
	{
		SBinary sBinStore = { 0 };
		SBinary sBinFld = { 0 };
		sBinStore.cb = lpAction->actMoveCopy.cbStoreEntryId;
		sBinStore.lpb = (LPBYTE)lpAction->actMoveCopy.lpStoreEntryId;
		sBinFld.cb = lpAction->actMoveCopy.cbFldEntryId;
		sBinFld.lpb = (LPBYTE)lpAction->actMoveCopy.lpFldEntryId;

		actstring += formatmessage(IDS_ACTIONOPMOVECOPY,
			BinToHexString(&sBinStore, true).c_str(),
			BinToTextString(&sBinStore, false).c_str(),
			BinToHexString(&sBinFld, true).c_str(),
			BinToTextString(&sBinFld, false).c_str());
		break;
	}
	case OP_REPLY:
	case OP_OOF_REPLY:
	{

		SBinary sBin = { 0 };
		sBin.cb = lpAction->actReply.cbEntryId;
		sBin.lpb = (LPBYTE)lpAction->actReply.lpEntryId;
		wstring szGUID = GUIDToStringAndName(&lpAction->actReply.guidReplyTemplate);

		actstring += formatmessage(IDS_ACTIONOPREPLY,
			BinToHexString(&sBin, true).c_str(),
			BinToTextString(&sBin, false).c_str(),
			szGUID.c_str());
		break;
	}
	case OP_DEFER_ACTION:
	{
		SBinary sBin = { 0 };
		sBin.cb = lpAction->actDeferAction.cbData;
		sBin.lpb = (LPBYTE)lpAction->actDeferAction.pbData;

		actstring += formatmessage(IDS_ACTIONOPDEFER,
			BinToHexString(&sBin, true).c_str(),
			BinToTextString(&sBin, false).c_str());
		break;
	}
	case OP_BOUNCE:
	{
		szFlags = InterpretFlags(flagBounceCode, lpAction->scBounceCode);
		actstring += formatmessage(IDS_ACTIONOPBOUNCE, lpAction->scBounceCode, szFlags.c_str());
		break;
	}
	case OP_FORWARD:
	case OP_DELEGATE:
	{
		actstring += formatmessage(IDS_ACTIONOPFORWARDDEL);
		actstring += AdrListToString(lpAction->lpadrlist);
		break;
	}

	case OP_TAG:
	{
		InterpretProp(&lpAction->propTag, &szProp, &szAltProp);
		actstring += formatmessage(IDS_ACTIONOPTAG,
			TagToString(lpAction->propTag.ulPropTag, NULL, false, true).c_str(),
			szProp.c_str(),
			szAltProp.c_str());
		break;
	}
	}

	switch (lpAction->acttype)
	{
	case OP_REPLY:
	{
		szFlags = InterpretFlags(flagOPReply, lpAction->ulActionFlavor);
		break;
	}
	case OP_FORWARD:
	{
		szFlags = InterpretFlags(flagOpForward, lpAction->ulActionFlavor);
		break;
	}
	}

	actstring += formatmessage(IDS_ACTIONFLAVOR, lpAction->ulActionFlavor, szFlags.c_str());

	if (!lpAction->lpPropTagArray)
	{
		actstring += loadstring(IDS_ACTIONTAGARRAYNULL);
	}
	else
	{
		actstring += formatmessage(IDS_ACTIONTAGARRAYCOUNT, lpAction->lpPropTagArray->cValues);
		ULONG i = 0;
		for (i = 0; i < lpAction->lpPropTagArray->cValues; i++)
		{
			actstring += formatmessage(IDS_ACTIONTAGARRAYTAG,
				i,
				TagToString(lpAction->lpPropTagArray->aulPropTag[i], NULL, false, false).c_str());
		}
	}

	return actstring;
}

wstring ActionsToString(_In_ ACTIONS* lpActions)
{
	if (!lpActions)
	{
		return loadstring(IDS_ACTIONSNULL);
	}

	wstring szFlags = InterpretFlags(flagRulesVersion, lpActions->ulVersion);
	wstring actstring = formatmessage(IDS_ACTIONSMEMBERS,
		lpActions->ulVersion,
		szFlags.c_str(),
		lpActions->cActions);

	UINT i = 0;
	for (i = 0; i < lpActions->cActions; i++)
	{
		actstring += formatmessage(IDS_ACTIONSACTION, i);
		actstring += ActionToString(&lpActions->lpAction[i]);
	}

	return actstring;
}

void FileTimeToString(_In_ FILETIME* lpFileTime, _In_ wstring& PropString, _In_opt_ wstring& AltPropString)
{
	HRESULT hRes = S_OK;
	SYSTEMTIME SysTime = { 0 };

	if (!lpFileTime) return;

	WC_B(FileTimeToSystemTime((FILETIME*)lpFileTime, &SysTime));

	if (S_OK == hRes)
	{
		int iRet = 0;
		wchar_t szTimeStr[MAX_PATH] = { 0 };
		wchar_t szDateStr[MAX_PATH] = { 0 };

		// shove millisecond info into our format string since GetTimeFormat doesn't use it
		wstring szFormatStr = formatmessage(IDS_FILETIMEFORMAT, SysTime.wMilliseconds);

		WC_D(iRet, GetTimeFormatW(
			LOCALE_USER_DEFAULT,
			NULL,
			&SysTime,
			szFormatStr.c_str(),
			szTimeStr,
			MAX_PATH));

		WC_D(iRet, GetDateFormatW(
			LOCALE_USER_DEFAULT,
			NULL,
			&SysTime,
			NULL,
			szDateStr,
			MAX_PATH));

		PropString = format(L"%ws %ws", szTimeStr, szDateStr); // STRING_OK
	}
	else
	{
		PropString = loadstring(IDS_INVALIDSYSTIME);
	}

	AltPropString = formatmessage(IDS_FILETIMEALTFORMAT, lpFileTime->dwLowDateTime, lpFileTime->dwHighDateTime);
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
void InterpretProp(_In_ LPSPropValue lpProp, _In_opt_ wstring* PropString, _In_opt_ wstring* AltPropString)
{
	if (!lpProp) return;

	Property parsedProperty = ParseProperty(lpProp);

	if (PropString) *PropString = parsedProperty.toString();
	if (AltPropString) *AltPropString = parsedProperty.toAltString();
}

wstring TypeToString(ULONG ulPropTag)
{
	wstring tmpPropType;

	bool bNeedInstance = false;
	if (ulPropTag & MV_INSTANCE)
	{
		ulPropTag &= ~MV_INSTANCE;
		bNeedInstance = true;
	}

	ULONG ulCur = 0;
	bool bTypeFound = false;

	for (ulCur = 0; ulCur < ulPropTypeArray; ulCur++)
	{
		if (PropTypeArray[ulCur].ulValue == PROP_TYPE(ulPropTag))
		{
			tmpPropType = PropTypeArray[ulCur].lpszName;
			bTypeFound = true;
			break;
		}
	}

	if (!bTypeFound)
		tmpPropType = format(L"0x%04x", PROP_TYPE(ulPropTag)); // STRING_OK

	if (bNeedInstance) tmpPropType += L" | MV_INSTANCE"; // STRING_OK
	return tmpPropType;
}

// TagToString will prepend the http://schemas.microsoft.com/MAPI/ for us since it's a constant
// We don't compute a DASL string for non-named props as FormatMessage in TagToString can handle those
void NameIDToStrings(_In_ LPMAPINAMEID lpNameID,
	ULONG ulPropTag,
	_In_ wstring& szPropName,
	_In_ wstring& szPropGUID,
	_In_ wstring& szDASL)
{
	HRESULT hRes = S_OK;

	// Can't generate strings without a MAPINAMEID structure
	if (!lpNameID) return;

	LPNAMEDPROPCACHEENTRY lpNamedPropCacheEntry = NULL;

	// If we're using the cache, look up the answer there and return
	if (fCacheNamedProps())
	{
		lpNamedPropCacheEntry = FindCacheEntry(PROP_ID(ulPropTag), lpNameID->lpguid, lpNameID->ulKind, lpNameID->Kind.lID, lpNameID->Kind.lpwstrName);
		if (lpNamedPropCacheEntry && lpNamedPropCacheEntry->bStringsCached)
		{
			szPropName = lpNamedPropCacheEntry->lpszPropName;
			szPropGUID = lpNamedPropCacheEntry->lpszPropGUID;
			szDASL = lpNamedPropCacheEntry->lpszDASL;
			return;
		}

		// We shouldn't ever get here without a cached entry
		if (!lpNamedPropCacheEntry)
		{
			DebugPrint(DBGNamedProp, L"NameIDToStrings: Failed to find cache entry for ulPropTag = 0x%08X\n", ulPropTag);
			return;
		}
	}

	DebugPrint(DBGNamedProp, L"Parsing named property\n");
	DebugPrint(DBGNamedProp, L"ulPropTag = 0x%08x\n", ulPropTag);
	szPropGUID = GUIDToStringAndName(lpNameID->lpguid);
	DebugPrint(DBGNamedProp, L"lpNameID->lpguid = %ws\n", szPropGUID.c_str());

	wstring szDASLGuid = GUIDToString(lpNameID->lpguid);

	if (lpNameID->ulKind == MNID_ID)
	{
		DebugPrint(DBGNamedProp, L"lpNameID->Kind.lID = 0x%04X = %d\n", lpNameID->Kind.lID, lpNameID->Kind.lID);
		wstring szName = NameIDToPropName(lpNameID);

		if (!szName.empty())
		{
			// Printing hex first gets a nice sort without spacing tricks
			szPropName = format(L"id: 0x%04X=%d = %ws", // STRING_OK
				lpNameID->Kind.lID,
				lpNameID->Kind.lID,
				szName.c_str());

		}
		else
		{
			// Printing hex first gets a nice sort without spacing tricks
			szPropName = format(L"id: 0x%04X=%d", // STRING_OK
				lpNameID->Kind.lID,
				lpNameID->Kind.lID);
		}

		szDASL = format(L"id/%s/%04X%04X", // STRING_OK
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
		WC_H(StringCchLengthA((LPSTR)lpNameID->Kind.lpwstrName, STRSAFE_MAX_CCH, &cchShortLen));
		WC_H(StringCchLengthW(lpNameID->Kind.lpwstrName, STRSAFE_MAX_CCH, &cchWideLen));

		if (cchShortLen < cchWideLen)
		{
			// this is the *proper* case
			DebugPrint(DBGNamedProp, L"lpNameID->Kind.lpwstrName = \"%ws\"\n", lpNameID->Kind.lpwstrName);
			szPropName = format(L"sz: \"%ws\"", lpNameID->Kind.lpwstrName);

			szDASL = format(L"string/%ws/%ws", // STRING_OK
				szDASLGuid.c_str(),
				lpNameID->Kind.lpwstrName);
		}
		else
		{
			// this is the case where ANSI data was shoved into a unicode string.
			DebugPrint(DBGNamedProp, L"Warning: ANSI data was found in a unicode field. This is a bug on the part of the creator of this named property\n");
			DebugPrint(DBGNamedProp, L"lpNameID->Kind.lpwstrName = \"%hs\"\n", (LPCSTR)lpNameID->Kind.lpwstrName);

			wstring szComment = loadstring(IDS_NAMEWASANSI);
			szPropName = format(L"sz: \"%hs\" %ws", (LPSTR)lpNameID->Kind.lpwstrName, szComment);

			szDASL = format(L"string/%ws/%hs", // STRING_OK
				szDASLGuid.c_str(),
				lpNameID->Kind.lpwstrName);
		}
	}

	// We've built our strings - if we're caching, put them in the cache
	if (lpNamedPropCacheEntry)
	{
		lpNamedPropCacheEntry->lpszPropName = szPropName;
		lpNamedPropCacheEntry->lpszPropGUID = szPropGUID;
		lpNamedPropCacheEntry->lpszDASL = szDASL;
		lpNamedPropCacheEntry->bStringsCached = true;
	}
}

// lpszNamedPropName, lpszNamedPropGUID, lpszNamedPropDASL freed with FreeNameIDStrings
void NameIDToStrings(
	ULONG ulPropTag, // optional 'original' prop tag
	_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
	_In_opt_ LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
	_In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
	bool bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
	_In_ wstring& lpszNamedPropName, // Built from ulPropTag & lpMAPIProp
	_In_ wstring& lpszNamedPropGUID, // Built from ulPropTag & lpMAPIProp
	_In_ wstring& lpszNamedPropDASL) // Built from ulPropTag & lpMAPIProp
{
	HRESULT hRes = S_OK;

	// Named Props
	LPMAPINAMEID* lppPropNames = 0;

	// If we weren't passed named property information and we need it, look it up
	// We check bIsAB here - some address book providers return garbage which will crash us
	if (!lpNameID &&
		lpMAPIProp && // if we have an object
		!bIsAB &&
		RegKeys[regkeyPARSED_NAMED_PROPS].ulCurDWORD && // and we're parsing named props
		(RegKeys[regkeyGETPROPNAMES_ON_ALL_PROPS].ulCurDWORD || PROP_ID(ulPropTag) >= 0x8000)) // and it's either a named prop or we're doing all props
	{
		SPropTagArray tag = { 0 };
		LPSPropTagArray lpTag = &tag;
		ULONG ulPropNames = 0;
		tag.cValues = 1;
		tag.aulPropTag[0] = ulPropTag;

		WC_H_GETPROPS(GetNamesFromIDs(lpMAPIProp,
			lpMappingSignature,
			&lpTag,
			NULL,
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
		NameIDToStrings(lpNameID,
			ulPropTag,
			lpszNamedPropName,
			lpszNamedPropGUID,
			lpszNamedPropDASL);
	}

	// Avoid making the call if we don't have to so we don't accidently depend on MAPI
	if (lppPropNames) MAPIFreeBuffer(lppPropNames);
}