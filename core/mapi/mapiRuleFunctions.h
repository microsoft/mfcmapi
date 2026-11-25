// MAPI rule functions
#pragma once

namespace mapi
{
	_Check_return_ HRESULT
	CopyFolderRules(_In_ LPMAPIFOLDER lpSrcFolder, _In_ LPMAPIFOLDER lpDestFolder, bool bReplace);
	struct RuleReportValue
	{
		HRESULT status{MAPI_E_NOT_FOUND};
		std::wstring raw;
		std::wstring decoded;
	};

	struct RuleReportRow
	{
		RuleReportValue id;
		RuleReportValue name;
		RuleReportValue provider;
		RuleReportValue sequence;
		RuleReportValue state;
	};

	_Check_return_ RuleReportRow DescribeRuleReportRow(const SRow& row);
	_Check_return_ HRESULT ReadRuleReportRows(_In_ LPMAPITABLE table, _Out_ std::vector<RuleReportRow>& rows);
} // namespace mapi
