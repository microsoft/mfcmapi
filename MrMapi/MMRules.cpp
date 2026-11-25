#include <StdAfx.h>
#include <MrMapi/MMRules.h>
#include <MrMapi/mmcli.h>
#include <MrMapi/MMAcls.h>
#include <core/utility/error.h>
#include <core/utility/csv.h>
#include <winrt/Windows.Data.Json.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <roapi.h>

using namespace winrt::Windows::Data::Json;

// Fails on text that cannot convert to UTF-8, such as a lone surrogate, so the report is not written with replacement characters.
static _Check_return_ HRESULT ValidateReportText(const std::wstring& text)
{
	if (text.empty()) return S_OK;
	if (text.size() > static_cast<size_t>(INT_MAX)) return WC_H(HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE));
	return WC_B(WideCharToMultiByte(
		CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr));
}

// Windows.Data.Json can only emit compact text, which has no whitespace outside strings.
// Structural characters are ASCII and never appear inside a multibyte UTF-8 sequence.
static std::string IndentJson(const std::string& compact)
{
	std::string result;
	result.reserve(compact.size() * 2);
	size_t depth = 0;
	auto inString = false;
	const auto newline = [&] {
		result += '\n';
		result.append(depth * 2, ' ');
	};

	for (size_t i = 0; i < compact.size(); ++i)
	{
		const auto c = compact[i];
		if (inString)
		{
			result += c;
			if (c == '\\' && i + 1 < compact.size())
				result += compact[++i];
			else if (c == '"')
				inString = false;
			continue;
		}

		switch (c)
		{
		case '"':
			inString = true;
			result += c;
			break;
		case '{':
		case '[':
			result += c;
			if (i + 1 < compact.size() && (compact[i + 1] == '}' || compact[i + 1] == ']'))
			{
				result += compact[++i];
			}
			else
			{
				++depth;
				newline();
			}

			break;
		case '}':
		case ']':
			--depth;
			newline();
			result += c;
			break;
		case ',':
			result += c;
			newline();
			break;
		case ':':
			result += ": ";
			break;
		default:
			result += c;
			break;
		}
	}

	result += '\n';
	return result;
}

static JsonObject ReportValue(const wchar_t* property, ULONG tag, const mapi::RuleReportValue& value)
{
	JsonObject object;
	object.Insert(L"property", JsonValue::CreateStringValue(property));
	object.Insert(L"tag", JsonValue::CreateStringValue(strings::format(L"0x%08X", tag)));
	if (value.status != S_OK)
		object.Insert(L"status", JsonValue::CreateStringValue(strings::format(L"0x%08X", value.status)));
	if (SUCCEEDED(value.status))
	{
		object.Insert(L"value", JsonValue::CreateStringValue(value.raw));
		if (!value.decoded.empty()) object.Insert(L"decoded", JsonValue::CreateStringValue(value.decoded));
	}
	else
	{
		object.Insert(L"value", JsonValue::CreateNullValue());
	}

	return object;
}

_Check_return_ HRESULT
FormatRulesReport(const std::vector<mapi::RuleReportRow>& rows, HRESULT tableStatus, _Out_ std::string& json)
{
	json.clear();
	for (const auto& row : rows)
	{
		for (const auto value : {&row.id, &row.name, &row.provider, &row.sequence, &row.state})
		{
			if (FAILED(value->status)) continue;
			const auto rawResult = ValidateReportText(value->raw);
			if (FAILED(rawResult)) return rawResult;
			const auto decodedResult = ValidateReportText(value->decoded);
			if (FAILED(decodedResult)) return decodedResult;
		}
	}

	const auto initialized = RoInitialize(RO_INIT_SINGLETHREADED);
	if (FAILED(initialized) && initialized != RPC_E_CHANGED_MODE) return WC_H(initialized);

	struct RuntimeCleanup
	{
		bool uninitialize;
		~RuntimeCleanup()
		{
			if (uninitialize)
			{
				winrt::clear_factory_cache();
				RoUninitialize();
			}
		}
	} cleanup{SUCCEEDED(initialized)};

	try
	{
		JsonObject report;
		report.Insert(L"schemaVersion", JsonValue::CreateNumberValue(1));
		report.Insert(L"scope", JsonValue::CreateStringValue(L"standard-rule metadata only"));
		report.Insert(L"complete", JsonValue::CreateBooleanValue(tableStatus == S_OK));
		if (tableStatus != S_OK)
			report.Insert(L"tableStatus", JsonValue::CreateStringValue(strings::format(L"0x%08X", tableStatus)));
		JsonArray values;
		for (const auto& row : rows)
		{
			JsonObject item;
			item.Insert(L"id", ReportValue(L"PR_RULE_ID", PR_RULE_ID, row.id));
			item.Insert(L"name", ReportValue(L"PR_RULE_NAME", CHANGE_PROP_TYPE(PR_RULE_NAME, PT_UNICODE), row.name));
			item.Insert(
				L"provider",
				ReportValue(L"PR_RULE_PROVIDER", CHANGE_PROP_TYPE(PR_RULE_PROVIDER, PT_UNICODE), row.provider));
			item.Insert(L"sequence", ReportValue(L"PR_RULE_SEQUENCE", PR_RULE_SEQUENCE, row.sequence));
			item.Insert(L"state", ReportValue(L"PR_RULE_STATE", PR_RULE_STATE, row.state));
			values.Append(item);
		}

		report.Insert(L"rules", values);
		json = IndentJson(winrt::to_string(report.Stringify()));
		return S_OK;
	} catch (const winrt::hresult_error& error)
	{
		return WC_H(error.code());
	}
}

void DoRules(_In_opt_ LPMAPIFOLDER lpFolder) { DumpExchangeTable(PR_RULES_TABLE, lpFolder); }

_Check_return_ HRESULT DoRulesReport(_In_opt_ LPMAPIFOLDER folder, const std::wstring& path)
{
	std::vector<mapi::RuleReportRow> rows;
	LPEXCHANGEMODIFYTABLE exchange = nullptr;
	LPMAPITABLE table = nullptr;
	auto tableStatus = folder ? WC_MAPI(folder->OpenProperty(
									PR_RULES_TABLE,
									const_cast<LPGUID>(&IID_IExchangeModifyTable),
									0,
									MAPI_DEFERRED_ERRORS,
									reinterpret_cast<LPUNKNOWN*>(&exchange)))
							  : MAPI_E_NOT_FOUND;
	if (tableStatus == S_OK && !exchange) tableStatus = MAPI_E_CALL_FAILED;
	if (tableStatus == S_OK) tableStatus = WC_MAPI(exchange->GetTable(0, &table));
	if (tableStatus == S_OK && !table) tableStatus = MAPI_E_CALL_FAILED;
	if (tableStatus == S_OK) tableStatus = mapi::ReadRuleReportRows(table, rows);
	if (table) table->Release();
	if (exchange) exchange->Release();

	std::string json;
	auto result = FormatRulesReport(rows, tableStatus, json);
	if (SUCCEEDED(result)) result = csv::Write(path, json);
	if (SUCCEEDED(result) && tableStatus != S_OK) result = FAILED(tableStatus) ? tableStatus : MAPI_E_CALL_FAILED;
	if (FAILED(result)) wprintf_s(L"Rules report incomplete or not written: 0x%08lX\n", result);
	return result;
}
