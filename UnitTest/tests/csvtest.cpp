#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/utility/csv.h>
#include <CommCtrl.h>

#pragma comment(lib, "comctl32.lib")

namespace csvtest
{
	csv::Table Snapshot(HWND list)
	{
		csv::Table table;
		Assert::AreEqual(S_OK, csv::Snapshot(list, table));
		return table;
	}

	std::string Format(const csv::Table& table)
	{
		std::string bytes;
		Assert::AreEqual(S_OK, csv::Format(table, bytes));
		return bytes;
	}
	int CALLBACK CompareRows(LPARAM left, LPARAM right, LPARAM) { return left < right ? -1 : left > right ? 1 : 0; }

	class NativeList
	{
	public:
		NativeList()
		{
			INITCOMMONCONTROLSEX controls{sizeof(controls), ICC_LISTVIEW_CLASSES};
			Assert::IsTrue(InitCommonControlsEx(&controls) != FALSE);
			window = CreateWindowExW(
				0,
				WC_LISTVIEWW,
				L"",
				WS_POPUP | LVS_REPORT,
				0,
				0,
				100,
				100,
				nullptr,
				nullptr,
				GetModuleHandleW(nullptr),
				nullptr);
			Assert::IsNotNull(window);
		}
		~NativeList() { DestroyWindow(window); }
		NativeList(const NativeList&) = delete;
		NativeList& operator=(const NativeList&) = delete;

		void Column(int index, const std::wstring& name, int width = 100)
		{
			LVCOLUMNW column{};
			column.mask = LVCF_TEXT | LVCF_WIDTH;
			column.pszText = const_cast<wchar_t*>(name.c_str());
			column.cx = width;
			Assert::AreEqual(
				index,
				static_cast<int>(SendMessageW(window, LVM_INSERTCOLUMNW, index, reinterpret_cast<LPARAM>(&column))));
		}
		void Row(int index, const std::wstring& first, const std::wstring& second, LPARAM sortKey = 0)
		{
			LVITEMW item{};
			item.mask = LVIF_TEXT | LVIF_PARAM;
			item.iItem = index;
			item.lParam = sortKey;
			item.pszText = const_cast<wchar_t*>(first.c_str());
			Assert::AreEqual(
				index, static_cast<int>(SendMessageW(window, LVM_INSERTITEMW, 0, reinterpret_cast<LPARAM>(&item))));
			item.iSubItem = 1;
			item.pszText = const_cast<wchar_t*>(second.c_str());
			Assert::IsTrue(SendMessageW(window, LVM_SETITEMTEXTW, index, reinterpret_cast<LPARAM>(&item)) != FALSE);
		}
		HWND window{};
	};

	class TestFile
	{
	public:
		TestFile()
		{
			path = L"csv-test-" + std::to_wstring(GetCurrentProcessId()) + L"-" + std::to_wstring(GetTickCount64()) +
				   L".csv";
		}
		~TestFile() { DeleteFileW(path.c_str()); }
		TestFile(const TestFile&) = delete;
		TestFile& operator=(const TestFile&) = delete;
		std::wstring path;
	};

	TEST_CLASS(csvtest)
	{
	public:
		static const bool dummy_var = true;

		TEST_METHOD(ExactQuotedUtf8)
		{
			const csv::Table table{
				{L"Sub,ject", L"Quote\"", L"Empty"},
				{{L"a,b", L"say \"hello\"", L""},
				 {L"line1\r\nline2\nline3\rend", L" \tkeep spaces ", L"\u00E9\u4E2D\U0001F600"}}};
			const std::string expected =
				"\xEF\xBB\xBF"
				"\"Sub,ject\",\"Quote\"\"\",\"Empty\"\r\n"
				"\"a,b\",\"say \"\"hello\"\"\",\"\"\r\n"
				"\"line1\r\nline2\nline3\rend\",\" \tkeep spaces \",\"\xC3\xA9\xE4\xB8\xAD\xF0\x9F\x98\x80\"\r\n";
			Assert::AreEqual(expected, Format(table));
		}

		TEST_METHOD(FormulaLikeTextIsNotSilentlyChanged)
		{
			Assert::AreEqual(
				std::string(
					"\xEF\xBB\xBF\"Header\"\r\n\"=1+1\"\r\n\"+1\"\r\n\"-1\"\r\n\"@name\"\r\n\"\t=1\"\r\n\"'=1\"\r\n"),
				Format({{L"Header"}, {{L"=1+1"}, {L"+1"}, {L"-1"}, {L"@name"}, {L"\t=1"}, {L"'=1"}}}));
		}

		TEST_METHOD(EmptyTableAndHeaderOnly)
		{
			Assert::AreEqual(std::string("\xEF\xBB\xBF"), Format({}));
			Assert::AreEqual(std::string("\xEF\xBB\xBF\"A\",\"\"\r\n"), Format({{L"A", L""}, {}}));
		}

		TEST_METHOD(InvalidUnicodeFailsInsteadOfReplacingText)
		{
			std::string bytes{"stale"};
			Assert::AreEqual(
				HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION),
				csv::Format({{L"Header"}, {{std::wstring(1, wchar_t{0xD800})}}}, bytes));
			Assert::IsTrue(bytes.empty());
		}

		TEST_METHOD(NativeSnapshotVisualOrderAllRowsAndImmutableText)
		{
			NativeList list;
			list.Column(0, L"Subject");
			list.Column(1, L"Sender");
			list.Column(2, L"Hidden", 0);
			list.Row(0, L"second", L"B");
			list.Row(0, L"first", L"A");
			int order[] = {2, 1, 0};
			Assert::IsTrue(
				SendMessageW(list.window, LVM_SETCOLUMNORDERARRAY, 3, reinterpret_cast<LPARAM>(order)) != FALSE);
			LVITEMW state{};
			state.stateMask = LVIS_SELECTED;
			state.state = LVIS_SELECTED;
			SendMessageW(list.window, LVM_SETITEMSTATE, 1, reinterpret_cast<LPARAM>(&state));
			const auto table = Snapshot(list.window);
			SendMessageW(list.window, LVM_DELETEALLITEMS, 0, 0);
			Assert::AreEqual(
				std::string("\xEF\xBB\xBF\"Sender\",\"Subject\"\r\n\"A\",\"first\"\r\n\"B\",\"second\"\r\n"),
				Format(table));
		}

		TEST_METHOD(NativeSnapshotDoesNotTruncateHeadersOrCells)
		{
			NativeList list;
			const auto header = std::wstring(4096, L'H');
			const auto value = std::wstring(16384, L'V') + L"\u00E9";
			list.Column(0, header);
			list.Column(1, L"");
			list.Row(0, value, L"");
			const auto snapshot = Snapshot(list.window);
			unittest::AreEqualEx(header, snapshot.headers.at(0));
			unittest::AreEqualEx(value, snapshot.rows.at(0).at(0));
			Assert::AreEqual(Format({{header, L""}, {{value, L""}}}), Format(snapshot));
		}

		TEST_METHOD(NativeSnapshotFollowsClientSideSort)
		{
			NativeList list;
			list.Column(0, L"Subject");
			list.Column(1, L"Sender");
			list.Row(0, L"second", L"B", 2);
			list.Row(1, L"first", L"A", 1);
			Assert::IsTrue(SendMessageW(list.window, LVM_SORTITEMS, 0, reinterpret_cast<LPARAM>(CompareRows)) != FALSE);
			Assert::AreEqual(
				std::string("\xEF\xBB\xBF\"Subject\",\"Sender\"\r\n\"first\",\"A\"\r\n\"second\",\"B\"\r\n"),
				Format(Snapshot(list.window)));
		}

		TEST_METHOD(NativeSnapshotWithoutDisplayedColumns)
		{
			NativeList list;
			Assert::AreEqual(std::string("\xEF\xBB\xBF"), Format(Snapshot(list.window)));
			list.Column(0, L"Hidden", 0);
			Assert::AreEqual(std::string("\xEF\xBB\xBF"), Format(Snapshot(list.window)));
		}

		TEST_METHOD(NativeEmptyListStillExportsHeaders)
		{
			NativeList list;
			list.Column(0, L"Header");
			Assert::AreEqual(std::string("\xEF\xBB\xBF\"Header\"\r\n"), Format(Snapshot(list.window)));
		}

		TEST_METHOD(InvalidWindowFails)
		{
			csv::Table table{{L"stale"}, {}};
			Assert::AreEqual(HRESULT_FROM_WIN32(ERROR_INVALID_THREAD_ID), csv::Snapshot(nullptr, table));
			Assert::IsTrue(table.headers.empty());
			Assert::IsTrue(table.rows.empty());
		}

		TEST_METHOD(WriteExactBytesAcrossChunksAndReportSharingErrors)
		{
			TestFile file;
			const auto bytes = Format({{L"Header"}, {{std::wstring(1024 * 1024 + 17, L'X')}}});
			Assert::AreEqual(S_OK, csv::Write(file.path, bytes));
			const auto handle = CreateFileW(
				file.path.c_str(),
				GENERIC_READ,
				FILE_SHARE_READ,
				nullptr,
				OPEN_EXISTING,
				FILE_ATTRIBUTE_NORMAL,
				nullptr);
			Assert::IsTrue(handle != INVALID_HANDLE_VALUE);
			std::string actual(bytes.size(), '\0');
			DWORD read = 0;
			const auto readOK = ReadFile(handle, actual.data(), static_cast<DWORD>(actual.size()), &read, nullptr);
			const auto writeError = csv::Write(file.path, "must not replace");
			CloseHandle(handle);
			Assert::IsTrue(readOK != FALSE);
			Assert::AreEqual(static_cast<DWORD>(bytes.size()), read);
			Assert::AreEqual(bytes, actual);
			Assert::AreEqual(HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION), writeError);
		}

		TEST_METHOD(WriteReportsMissingDirectory)
		{
			TestFile file;
			Assert::AreEqual(
				HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND), csv::Write(file.path + L"\\missing.csv", "data"));
		}
	};
} // namespace csvtest
