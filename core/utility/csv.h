#pragma once
#include <string>
#include <vector>

namespace csv
{
	struct Table
	{
		std::vector<std::wstring> headers;
		std::vector<std::vector<std::wstring>> rows;
	};

	// Call on the list's UI thread, outside loading, without pumping messages.
	_Check_return_ HRESULT Snapshot(HWND list, Table& table);

	// UTF-8 BOM, quoted fields, CRLF records. Cell contents are never sanitized or normalized.
	_Check_return_ HRESULT Format(const Table& table, std::string& bytes);
	_Check_return_ HRESULT Write(const std::wstring& path, const std::string& bytes) noexcept;
} // namespace csv
