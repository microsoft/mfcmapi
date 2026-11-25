#pragma once
// Rule table dumping for MrMAPI

#include <core/mapi/mapiRuleFunctions.h>

void DoRules(_In_opt_ LPMAPIFOLDER lpFolder);
_Check_return_ HRESULT DoRulesReport(_In_opt_ LPMAPIFOLDER folder, const std::wstring& path);
_Check_return_ HRESULT
FormatRulesReport(const std::vector<mapi::RuleReportRow>& rows, HRESULT tableStatus, _Out_ std::string& json);
