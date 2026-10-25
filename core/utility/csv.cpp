#include <core/stdafx.h>
#include <core/utility/csv.h>
#include <CommCtrl.h>
#include <limits>
#include <core/utility/error.h>

namespace csv
{
	_Check_return_ HRESULT Snapshot(HWND list, Table& table)
	{
		table = {};

		if (GetWindowThreadProcessId(list, nullptr) != GetCurrentThreadId())
		{
			return WC_H(HRESULT_FROM_WIN32(ERROR_INVALID_THREAD_ID));
		}

		const auto header = reinterpret_cast<HWND>(::SendMessageW(list, LVM_GETHEADER, 0, 0));
		if (!header) return S_OK;

		const auto count = static_cast<int>(::SendMessageW(header, HDM_GETITEMCOUNT, 0, 0));
		if (count < 0) return WC_H(HRESULT_FROM_WIN32(ERROR_INVALID_DATA));

		std::vector<int> columns(count);
		if (!columns.empty() &&
			!::SendMessageW(list, LVM_GETCOLUMNORDERARRAY, columns.size(), reinterpret_cast<LPARAM>(columns.data())))
		{
			return WC_H(HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
		}

		columns.erase(
			std::remove_if(
				columns.begin(),
				columns.end(),
				[list](int column) { return ::SendMessageW(list, LVM_GETCOLUMNWIDTH, column, 0) == 0; }),
			columns.end());

		const auto readText = [list](int row, int column, std::wstring& value) -> HRESULT {
			std::vector<wchar_t> text(256);
			for (;;)
			{
				if (row < 0)
				{
					LVCOLUMNW info{};
					info.mask = LVCF_TEXT;
					info.pszText = text.data();
					info.cchTextMax = static_cast<int>(text.size());
					if (!::SendMessageW(list, LVM_GETCOLUMNW, column, reinterpret_cast<LPARAM>(&info)))
					{
						return WC_H(HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
					}
				}
				else
				{
					LVITEMW info{};
					info.iSubItem = column;
					info.pszText = text.data();
					info.cchTextMax = static_cast<int>(text.size());
					::SendMessageW(list, LVM_GETITEMTEXTW, row, reinterpret_cast<LPARAM>(&info));
				}

				if (wcslen(text.data()) < text.size() - 1)
				{
					value = text.data();
					return S_OK;
				}

				if (text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()) / 2)
				{
					return WC_H(E_OUTOFMEMORY);
				}

				text.resize(text.size() * 2);
			}
		};

		Table snapshot;
		for (const auto column : columns)
		{
			const auto hRes = readText(-1, column, snapshot.headers.emplace_back());
			if (FAILED(hRes)) return hRes;
		}

		const auto rows = static_cast<int>(::SendMessageW(list, LVM_GETITEMCOUNT, 0, 0));
		if (rows < 0) return WC_H(HRESULT_FROM_WIN32(ERROR_INVALID_DATA));

		snapshot.rows.reserve(rows);
		for (auto row = 0; row < rows; ++row)
		{
			auto& values = snapshot.rows.emplace_back();
			for (const auto column : columns)
			{
				const auto hRes = readText(row, column, values.emplace_back());
				if (FAILED(hRes)) return hRes;
			}
		}

		table = std::move(snapshot);

		return S_OK;
	}

	_Check_return_ HRESULT Format(const Table& table, std::string& bytes)
	{
		bytes.clear();
		std::wstring text;
		const auto appendRow = [&text](const std::vector<std::wstring>& row) {
			auto first = true;
			for (const auto& value : row)
			{
				if (!first) text += L',';
				first = false;
				text += L'"';
				for (const auto ch : value)
				{
					text += ch;
					if (ch == L'"') text += L'"';
				}

				text += L'"';
			}

			text += L"\r\n";
		};

		if (!table.headers.empty())
		{
			appendRow(table.headers);
			for (const auto& row : table.rows)
			{
				appendRow(row);
			}
		}

		std::string output = "\xEF\xBB\xBF";
		if (text.empty())
		{
			bytes = std::move(output);
			return S_OK;
		}

		if (text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
		{
			return WC_H(HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE));
		}

		const auto length = static_cast<int>(text.size());
		const auto size =
			WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), length, nullptr, 0, nullptr, nullptr);
		if (!size) return WC_H(HRESULT_FROM_WIN32(GetLastError()));

		output.resize(3 + static_cast<size_t>(size));
		const auto hRes = WC_B(WideCharToMultiByte(
			CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), length, output.data() + 3, size, nullptr, nullptr));
		if (SUCCEEDED(hRes)) bytes = std::move(output);

		return hRes;
	}

	_Check_return_ HRESULT Write(const std::wstring& path, const std::string& bytes) noexcept
	{
		const auto handle =
			CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
		if (handle == INVALID_HANDLE_VALUE) return WC_H(HRESULT_FROM_WIN32(GetLastError()));

		auto hRes = S_OK;
		size_t offset = 0;
		while (offset < bytes.size())
		{
			const auto remaining = bytes.size() - offset;
			const auto count = static_cast<DWORD>((std::min)(remaining, size_t{1024 * 1024}));
			DWORD written = 0;
			hRes = WC_B(WriteFile(handle, bytes.data() + offset, count, &written, nullptr));
			if (FAILED(hRes)) break;

			if (!written)
			{
				hRes = WC_H(STG_E_WRITEFAULT);
				break;
			}

			offset += written;
		}

		if (SUCCEEDED(hRes)) hRes = WC_B(FlushFileBuffers(handle));
		const auto closeResult = WC_B(CloseHandle(handle));

		return FAILED(hRes) ? hRes : closeResult;
	}
} // namespace csv
