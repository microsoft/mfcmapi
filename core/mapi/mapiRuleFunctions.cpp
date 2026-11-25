#include <core/stdafx.h>
#include <core/mapi/mapiRuleFunctions.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/mapiMemory.h>
#include <core/mapi/extraPropTags.h>
#include <core/interpret/flags.h>
#include <core/utility/error.h>
#include <core/utility/strings.h>

namespace mapi
{
	_Check_return_ HRESULT CopyFolderRules(_In_ LPMAPIFOLDER lpSrcFolder, _In_ LPMAPIFOLDER lpDestFolder, bool bReplace)
	{
		if (!lpSrcFolder || !lpDestFolder) return MAPI_E_INVALID_PARAMETER;
		LPEXCHANGEMODIFYTABLE lpSrcTbl = nullptr;
		LPEXCHANGEMODIFYTABLE lpDstTbl = nullptr;

		auto hRes = EC_MAPI(lpSrcFolder->OpenProperty(
			PR_RULES_TABLE,
			const_cast<LPGUID>(&IID_IExchangeModifyTable),
			0,
			MAPI_DEFERRED_ERRORS,
			reinterpret_cast<LPUNKNOWN*>(&lpSrcTbl)));

		if (SUCCEEDED(hRes))
		{
			hRes = EC_MAPI(lpDestFolder->OpenProperty(
				PR_RULES_TABLE,
				const_cast<LPGUID>(&IID_IExchangeModifyTable),
				0,
				MAPI_DEFERRED_ERRORS,
				reinterpret_cast<LPUNKNOWN*>(&lpDstTbl)));
		}

		if (lpSrcTbl && lpDstTbl)
		{
			LPMAPITABLE lpTable = nullptr;
			lpSrcTbl->GetTable(0, &lpTable);

			if (lpTable)
			{
				static const SizedSPropTagArray(9, ruleTags) = {
					9,
					{PR_RULE_ACTIONS,
					 PR_RULE_CONDITION,
					 PR_RULE_LEVEL,
					 PR_RULE_NAME,
					 PR_RULE_PROVIDER,
					 PR_RULE_PROVIDER_DATA,
					 PR_RULE_SEQUENCE,
					 PR_RULE_STATE,
					 PR_RULE_USER_FLAGS},
				};

				hRes = EC_MAPI(lpTable->SetColumns(LPSPropTagArray(&ruleTags), 0));

				LPSRowSet lpRows = nullptr;

				if (SUCCEEDED(hRes))
				{
					hRes = EC_MAPI(HrQueryAllRows(lpTable, nullptr, nullptr, nullptr, 0, &lpRows));
				}

				if (lpRows && lpRows->cRows < MAXNewROWLIST)
				{
					auto lpTempList = mapi::allocate<LPROWLIST>(CbNewROWLIST(lpRows->cRows));
					if (lpTempList)
					{
						lpTempList->cEntries = lpRows->cRows;

						for (ULONG iArrayPos = 0; iArrayPos < lpRows->cRows; iArrayPos++)
						{
							lpTempList->aEntries[iArrayPos].ulRowFlags = ROW_ADD;
							lpTempList->aEntries[iArrayPos].rgPropVals = mapi::allocate<LPSPropValue>(
								lpRows->aRow[iArrayPos].cValues * sizeof(SPropValue), lpTempList);
							if (lpTempList->aEntries[iArrayPos].rgPropVals)
							{
								ULONG ulDst = 0;
								for (ULONG ulSrc = 0; ulSrc < lpRows->aRow[iArrayPos].cValues; ulSrc++)
								{
									if (lpRows->aRow[iArrayPos].lpProps[ulSrc].ulPropTag == PR_RULE_PROVIDER_DATA)
									{
										const auto& bin = mapi::getBin(lpRows->aRow[iArrayPos].lpProps[ulSrc]);
										if (!bin.cb || !bin.lpb)
										{
											// PR_RULE_PROVIDER_DATA was NULL - we don't want this
											continue;
										}
									}

									hRes = EC_H(MyPropCopyMore(
										&lpTempList->aEntries[iArrayPos].rgPropVals[ulDst],
										&lpRows->aRow[iArrayPos].lpProps[ulSrc],
										MAPIAllocateMore,
										lpTempList));
									ulDst++;
								}

								lpTempList->aEntries[iArrayPos].cValues = ulDst;
							}
						}

						ULONG ulFlags = 0;
						if (bReplace) ulFlags = ROWLIST_REPLACE;

						if (SUCCEEDED(hRes))
						{
							hRes = EC_MAPI(lpDstTbl->ModifyTable(ulFlags, lpTempList));
						}

						MAPIFreeBuffer(lpTempList);
					}

					FreeProws(lpRows);
				}

				lpTable->Release();
			}
		}

		if (lpDstTbl) lpDstTbl->Release();
		if (lpSrcTbl) lpSrcTbl->Release();
		return hRes;
	}

	static RuleReportValue DescribeRuleReportValue(const SRow& row, ULONG tag)
	{
		RuleReportValue value;
		if (!row.lpProps)
		{
			if (row.cValues) value.status = MAPI_E_INVALID_PARAMETER;
			return value;
		}

		auto prop = FindProp(row.lpProps, row.cValues, tag);
		if (!prop) prop = FindProp(row.lpProps, row.cValues, CHANGE_PROP_TYPE(tag, PT_ERROR));
		if (!prop)
		{
			for (ULONG i = 0; i < row.cValues; ++i)
			{
				if (PROP_ID(row.lpProps[i].ulPropTag) == PROP_ID(tag))
				{
					value.status = MAPI_E_INVALID_TYPE;
					break;
				}
			}

			return value;
		}
		if (PROP_TYPE(prop->ulPropTag) == PT_ERROR)
		{
			value.status = FAILED(prop->Value.err) ? prop->Value.err : MAPI_E_CALL_FAILED;
			return value;
		}

		if (prop->ulPropTag != tag)
		{
			value.status = MAPI_E_INVALID_TYPE;
			return value;
		}

		switch (PROP_TYPE(tag))
		{
		case PT_LONG:
			value.raw = strings::format(L"0x%08X", prop->Value.ul);
			if (tag == PR_RULE_STATE) value.decoded = flags::InterpretFlags(PROP_ID(tag), prop->Value.l);
			break;
		case PT_I8:
			value.raw =
				strings::format(L"0x%08X%08X", static_cast<ULONG>(prop->Value.li.HighPart), prop->Value.li.LowPart);
			break;
		case PT_UNICODE:
			if (!prop->Value.lpszW)
			{
				value.status = MAPI_E_INVALID_PARAMETER;
				return value;
			}

			value.raw = prop->Value.lpszW;
			break;
		default:
			value.status = MAPI_E_INVALID_TYPE;
			return value;
		}

		value.status = S_OK;
		return value;
	}

	_Check_return_ RuleReportRow DescribeRuleReportRow(const SRow& row)
	{
		return {
			DescribeRuleReportValue(row, PR_RULE_ID),
			DescribeRuleReportValue(row, CHANGE_PROP_TYPE(PR_RULE_NAME, PT_UNICODE)),
			DescribeRuleReportValue(row, CHANGE_PROP_TYPE(PR_RULE_PROVIDER, PT_UNICODE)),
			DescribeRuleReportValue(row, PR_RULE_SEQUENCE),
			DescribeRuleReportValue(row, PR_RULE_STATE)};
	}

	_Check_return_ HRESULT ReadRuleReportRows(_In_ LPMAPITABLE table, _Out_ std::vector<RuleReportRow>& rows)
	{
		rows.clear();
		if (!table) return MAPI_E_INVALID_PARAMETER;

		SizedSPropTagArray(5, columns) = {
			5,
			{PR_RULE_ID,
			 CHANGE_PROP_TYPE(PR_RULE_NAME, PT_UNICODE),
			 CHANGE_PROP_TYPE(PR_RULE_PROVIDER, PT_UNICODE),
			 PR_RULE_SEQUENCE,
			 PR_RULE_STATE}};
		auto hRes = WC_MAPI(table->SetColumns(reinterpret_cast<LPSPropTagArray>(&columns), 0));
		if (hRes != S_OK) return hRes;

		for (;;)
		{
			LPSRowSet batch = nullptr;
			hRes = WC_MAPI(table->QueryRows(100, 0, &batch));
			if (FAILED(hRes))
			{
				if (batch) FreeProws(batch);
				return hRes;
			}

			if (!batch) return MAPI_E_CALL_FAILED;
			const auto count = batch->cRows;
			for (ULONG i = 0; i < count; ++i)
				rows.push_back(DescribeRuleReportRow(batch->aRow[i]));
			FreeProws(batch);
			if (hRes != S_OK || !count) return hRes;
		}
	}
} // namespace mapi
