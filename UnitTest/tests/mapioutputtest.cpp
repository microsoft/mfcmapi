#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <core/mapi/mapiMemory.h>
#include <core/mapi/mapiOutput.h>
#include <core/mapi/mapiFunctions.h>
#include <core/utility/error.h>
#include <core/utility/output.h>
#include <core/utility/registry.h>
#include <fcntl.h>
#include <io.h>
#include <cstddef>
#include <limits>
#include <memory>

// Resolve test memory APIs through this DLL, not an installed MAPI provider.
#ifdef _M_IX86
#pragma comment(linker, "/export:MAPIAllocateBuffer@8=_TestMAPIAllocateBuffer@8")
#pragma comment(linker, "/export:MAPIFreeBuffer@4=_TestMAPIFreeBuffer@4")
#pragma comment(linker, "/export:FreeProws@4=_TestFreeProws@4")
#pragma comment(linker, "/export:FreePadrlist@4=_TestFreePadrlist@4")
#pragma comment(linker, "/export:PpropFindProp@12=_TestPpropFindProp@12")
#else
#pragma comment(linker, "/export:MAPIAllocateBuffer=TestMAPIAllocateBuffer")
#pragma comment(linker, "/export:MAPIFreeBuffer=TestMAPIFreeBuffer")
#pragma comment(linker, "/export:FreeProws=TestFreeProws")
#pragma comment(linker, "/export:FreePadrlist=TestFreePadrlist")
#pragma comment(linker, "/export:PpropFindProp=TestPpropFindProp")
#endif

namespace
{
	union Buffer {
		Buffer* next;
		std::max_align_t alignment;
	};

	Buffer* buffers{};
	ULONG invalidFrees{};
	ULONG freedRowSets{};
	std::vector<ULONG> freedPropertyTags;
} // namespace

extern "C" SCODE STDAPICALLTYPE TestMAPIAllocateBuffer(ULONG size, LPVOID* result)
{
	if (!result) return MAPI_E_INVALID_PARAMETER;
	*result = nullptr;
	if (static_cast<size_t>(size) > (std::numeric_limits<size_t>::max)() - sizeof(Buffer))
		return MAPI_E_NOT_ENOUGH_MEMORY;
	const auto buffer = static_cast<Buffer*>(HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(Buffer) + size));
	if (!buffer) return MAPI_E_NOT_ENOUGH_MEMORY;
	buffer->next = buffers;
	buffers = buffer;
	*result = buffer + 1;
	return S_OK;
}

extern "C" ULONG STDAPICALLTYPE TestMAPIFreeBuffer(LPVOID buffer)
{
	if (!buffer) return 0;
	auto link = &buffers;
	while (*link && *link + 1 != buffer)
		link = &(*link)->next;
	if (!*link)
	{
		++invalidFrees;
		return static_cast<ULONG>(MAPI_E_INVALID_PARAMETER);
	}

	const auto allocation = *link;
	*link = allocation->next;
	HeapFree(GetProcessHeap(), 0, allocation);
	return 0;
}

extern "C" void STDAPICALLTYPE TestFreeProws(LPSRowSet rows)
{
	if (!rows) return;
	++freedRowSets;
	for (ULONG row = 0; row < rows->cRows; ++row)
	{
		for (ULONG prop = 0; prop < rows->aRow[row].cValues; ++prop)
			freedPropertyTags.push_back(rows->aRow[row].lpProps[prop].ulPropTag);
		TestMAPIFreeBuffer(rows->aRow[row].lpProps);
	}

	TestMAPIFreeBuffer(rows);
}

extern "C" void STDAPICALLTYPE TestFreePadrlist(LPADRLIST recipients)
{
	if (!recipients) return;
	for (ULONG row = 0; row < recipients->cEntries; ++row)
		TestMAPIFreeBuffer(recipients->aEntries[row].rgPropVals);
	TestMAPIFreeBuffer(recipients);
}

extern "C" LPSPropValue STDAPICALLTYPE TestPpropFindProp(LPSPropValue props, ULONG count, ULONG tag)
{
	if (!props) return nullptr;
	for (ULONG index = 0; index < count; ++index)
	{
		if (props[index].ulPropTag == tag ||
			(PROP_TYPE(tag) == PT_UNSPECIFIED && PROP_ID(props[index].ulPropTag) == PROP_ID(tag)))
			return &props[index];
	}

	return nullptr;
}

namespace mapioutputtest
{
	constexpr ULONG numberTag = 0x60010003;
	constexpr ULONG stringTag = 0x6002001F;

	class TestTable final : public IMAPITable
	{
	public:
		TestTable(ULONG rowCount, ULONG batchSize, bool withProperties = false)
			: m_rowCount(rowCount), m_batchSize(batchSize), m_withProperties(withProperties)
		{
		}

		HRESULT STDMETHODCALLTYPE QueryInterface(REFIID, void** object) override
		{
			*object = nullptr;
			return E_NOINTERFACE;
		}
		ULONG STDMETHODCALLTYPE AddRef() override { return ++m_refs; }
		ULONG STDMETHODCALLTYPE Release() override { return --m_refs; }
		HRESULT STDMETHODCALLTYPE GetLastError(HRESULT, ULONG, LPMAPIERROR*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE Advise(ULONG, LPMAPIADVISESINK, ULONG_PTR*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE Unadvise(ULONG_PTR) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE GetStatus(ULONG*, ULONG*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE SetColumns(LPSPropTagArray, ULONG) override
		{
			++setColumnsCount;
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE QueryColumns(ULONG flags, LPSPropTagArray* tags) override
		{
			++queryColumnsCount;
			Assert::AreEqual(0UL, flags);
			*tags = nullptr;
			if (nullColumns) return columnsResult;
			const auto allocated = MAPIAllocateBuffer(
				static_cast<ULONG>(CbNewSPropTagArray(columns.size())), reinterpret_cast<void**>(tags));
			if (FAILED(allocated)) return allocated;
			(*tags)->cValues = static_cast<ULONG>(columns.size());
			std::copy(columns.begin(), columns.end(), (*tags)->aulPropTag);
			return columnsResult;
		}
		HRESULT STDMETHODCALLTYPE GetRowCount(ULONG, ULONG*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE SeekRow(BOOKMARK origin, LONG count, LONG* sought) override
		{
			Assert::AreEqual(static_cast<BOOKMARK>(BOOKMARK_BEGINNING), origin);
			Assert::AreEqual(0L, count);
			++seekCount;
			if (FAILED(seekResult)) return seekResult;
			position = 0;
			if (sought) *sought = 0;
			return S_OK;
		}
		HRESULT STDMETHODCALLTYPE SeekRowApprox(ULONG, ULONG) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE QueryPosition(ULONG*, ULONG*, ULONG*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE FindRow(LPSRestriction, BOOKMARK, ULONG) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE Restrict(LPSRestriction, ULONG) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE CreateBookmark(BOOKMARK*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE FreeBookmark(BOOKMARK) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE SortTable(LPSSortOrderSet, ULONG) override
		{
			++sortCount;
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE QuerySortOrder(LPSSortOrderSet*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE QueryRows(LONG requested, ULONG flags, LPSRowSet* rows) override
		{
			Assert::AreEqual(expectedRequested, requested);
			Assert::AreEqual(expectedFlags, flags);
			++queryCount;
			*rows = nullptr;
			if (queryCount == failOnQuery && !returnRowsOnFailure) return queryResult;
			if (position == m_rowCount && nullAtEnd) return S_OK;
			const auto count = (std::min)({static_cast<ULONG>(requested), m_batchSize, m_rowCount - position});
			*rows = mapi::allocate<LPSRowSet>(CbNewSRowSet(count));
			if (!*rows) return MAPI_E_NOT_ENOUGH_MEMORY;
			(*rows)->cRows = count;
			if (m_withProperties)
			{
				for (ULONG row = 0; row < count; ++row)
				{
					auto& outputRow = (*rows)->aRow[row];
					outputRow.lpProps = mapi::allocate<LPSPropValue>(2 * sizeof(SPropValue));
					if (!outputRow.lpProps) return MAPI_E_NOT_ENOUGH_MEMORY;
					outputRow.cValues = 2;
					outputRow.lpProps[0].ulPropTag = stringTag;
					outputRow.lpProps[0].Value.lpszW = m_text.data();
					outputRow.lpProps[1].ulPropTag = numberTag;
					outputRow.lpProps[1].Value.l = static_cast<LONG>(position + row);
				}
			}

			position += count;
			return queryCount == failOnQuery ? queryResult : S_OK;
		}
		HRESULT STDMETHODCALLTYPE Abort() override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE ExpandRow(ULONG, LPBYTE, ULONG, ULONG, LPSRowSet*, ULONG*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE CollapseRow(ULONG, LPBYTE, ULONG, ULONG*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE WaitForCompletion(ULONG, ULONG, ULONG*) override { return MAPI_E_NO_SUPPORT; }
		HRESULT STDMETHODCALLTYPE GetCollapseState(ULONG, ULONG, LPBYTE, ULONG*, LPBYTE*) override
		{
			return MAPI_E_NO_SUPPORT;
		}
		HRESULT STDMETHODCALLTYPE SetCollapseState(ULONG, ULONG, LPBYTE, BOOKMARK*) override
		{
			return MAPI_E_NO_SUPPORT;
		}

		ULONG queryCount{};
		ULONG seekCount{};
		ULONG position{};
		ULONG failOnQuery{};
		ULONG queryColumnsCount{};
		ULONG setColumnsCount{};
		ULONG sortCount{};
		LONG expectedRequested{20};
		ULONG expectedFlags{};
		HRESULT queryResult{MAPI_E_CALL_FAILED};
		HRESULT columnsResult{S_OK};
		std::vector<ULONG> columns;
		HRESULT seekResult{S_OK};
		bool returnRowsOnFailure{};
		bool nullAtEnd{};
		bool nullColumns{};

	private:
		ULONG m_refs{1};
		ULONG m_rowCount{};
		ULONG m_batchSize{};
		bool m_withProperties{};
		std::wstring m_text{L"\u03A9<&"};
	};

	class OutputFile
	{
	public:
		OutputFile()
		{
			FILE* file{};
			Assert::AreEqual(0, tmpfile_s(&file));
			m_file.reset(file);
			Assert::IsTrue(_setmode(_fileno(file), _O_U16TEXT) != -1);
		}

		FILE* get() const noexcept { return m_file.get(); }

		std::wstring read() const
		{
			Assert::AreEqual(0, fflush(get()));
			Assert::AreEqual(0, fseek(get(), 0, SEEK_SET));
			auto actual = std::wstring{};
			for (auto ch = fgetwc(get()); ch != WEOF; ch = fgetwc(get()))
				actual.push_back(static_cast<wchar_t>(ch));
			Assert::AreEqual(0, ferror(get()));
			return actual;
		}

	private:
		std::unique_ptr<FILE, decltype(&fclose)> m_file{nullptr, fclose};
	};

	std::wstring ExpectedNumber(ULONG tag, ULONG value)
	{
		const auto names =
			tag == numberTag ? L"\t\t<Name>PR_DOTSTUFF_STATE</Name>\n"
							   L"\t\t<OtherNames>PR_NICK_NAME, PR_NICK_NAME_A, PR_NICK_NAME_W</OtherNames>\n"
			: tag == 0x60030003
				? L"\t\t<Name>PR_DROPDOWN_DISPLAY_NAME</Name>\n"
				  L"\t\t<OtherNames>PR_DROPDOWN_DISPLAY_NAME_A, PR_DROPDOWN_DISPLAY_NAME_W</OtherNames>\n"
				: L"";
		return strings::format(
			L"\t<property tag = \"0x%08X\" type = \"PT_LONG\" >\n"
			L"%ws"
			L"\t\t<Value>%u</Value>\n"
			L"\t\t<AltValue>0x%X</AltValue>\n"
			L"\t</property>\n",
			tag,
			names,
			value,
			value);
	}

	std::wstring ExpectedProperties(ULONG value)
	{
		return ExpectedNumber(numberTag, value) + L"\t<property tag = \"0x6002001F\" type = \"PT_UNICODE\" >\n"
												  L"\t\t<Value><![CDATA[\u03A9<&]]></Value>\n"
												  L"\t\t<AltValue cb=\"6\" >A9033C002600</AltValue>\n"
												  L"\t</property>\n";
	}

	std::wstring ExpectedTable(ULONG rowCount, bool withProperties = false, ULONG firstValue = 0)
	{
		auto expected = std::wstring(L"<?xml version=\"1.0\" encoding=\"UTF-16\"?>\n<table>\n");
		for (ULONG row = 0; row < rowCount; ++row)
		{
			expected += strings::format(L"<row index = \"0x%08X\">\n", row);
			if (withProperties) expected += ExpectedProperties(firstValue + row);

			expected += L"</row>\n";
		}

		return expected + L"</table>\n";
	}

	TEST_CLASS(mapioutputtest)
	{
	public:
		TEST_CLASS_INITIALIZE(initialize) { unittest::init(); }

		TEST_METHOD_INITIALIZE(initializeMemory)
		{
			Assert::IsTrue(buffers == nullptr);
			invalidFrees = 0;
			freedRowSets = 0;
			freedPropertyTags.clear();
			m_previousFreedRowSets = 0;
			m_errors.clear();
			m_memory = std::make_unique<unittest::testMapiMemory>();
			m_originalDisplayError = std::move(error::displayError);
			error::displayError = [this](const std::wstring& text) { m_errors.push_back(text); };
		}

		TEST_METHOD_CLEANUP(uninitializeMemory)
		{
			ULONG leakedBuffers{};
			while (buffers)
			{
				++leakedBuffers;
				TestMAPIFreeBuffer(buffers + 1);
			}

			if (m_memory)
			{
				error::displayError = std::move(m_originalDisplayError);
				m_memory.reset();
			}

			Assert::AreEqual(0UL, leakedBuffers, L"The exporter leaked MAPI allocations.");
			Assert::AreEqual(0UL, invalidFrees, L"The exporter freed an unowned MAPI allocation.");
		}

		TEST_METHOD(Test_TableRowIndicesAcrossBatches) { CheckTable(41, 20); }

		TEST_METHOD(QueryRowsExplainsUnsupportedMVInstanceColumns)
		{
			for (const ULONG tag : {0x81AF3003UL, 0x800A301FUL, 0x800B3102UL})
			{
				auto table = TestTable(0, 20);
				table.columns = {PR_SUBJECT_W, tag};
				table.failOnQuery = 1;
				table.queryResult = MAPI_E_NO_SUPPORT;
				CheckQueryDiagnostic(table, true);
				Assert::AreEqual(1UL, table.queryColumnsCount);
			}
		}

		TEST_METHOD(QueryRowsDoesNotSuggestMVSortForOrdinaryColumns)
		{
			for (const auto& columns :
				 {std::vector<ULONG>{}, std::vector<ULONG>{PR_SUBJECT_W}, std::vector<ULONG>{0x81AF1003}})
			{
				auto table = TestTable(0, 20);
				table.columns = columns;
				table.failOnQuery = 1;
				table.queryResult = MAPI_E_NO_SUPPORT;
				CheckQueryDiagnostic(table, false);
				Assert::AreEqual(1UL, table.queryColumnsCount);
			}
		}

		TEST_METHOD(QueryRowsPreservesOtherFailuresWithoutQueryingColumns)
		{
			for (const auto result : {MAPI_E_CALL_FAILED, MAPI_E_NOT_FOUND, MAPI_E_INVALID_PARAMETER})
			{
				auto table = TestTable(0, 20);
				table.columns = {0x81AF3003};
				table.failOnQuery = 1;
				table.queryResult = result;
				CheckQueryDiagnostic(table, false);
				Assert::AreEqual(0UL, table.queryColumnsCount);
			}
		}

		TEST_METHOD(QueryRowsPreservesPrimaryFailureWhenColumnQueryFails)
		{
			auto table = TestTable(0, 20);
			table.columns = {0x81AF3003};
			table.columnsResult = MAPI_E_CALL_FAILED;
			table.failOnQuery = 1;
			table.queryResult = MAPI_E_NO_SUPPORT;
			CheckQueryDiagnostic(table, false);
			Assert::AreEqual(1UL, table.queryColumnsCount);
		}

		TEST_METHOD(QueryRowsHandlesUnavailableColumns)
		{
			for (const auto result : {S_OK, MAPI_E_CALL_FAILED})
			{
				auto table = TestTable(0, 20);
				table.nullColumns = true;
				table.columnsResult = result;
				table.failOnQuery = 1;
				table.queryResult = MAPI_E_NO_SUPPORT;
				CheckQueryDiagnostic(table, false);
			}
		}

		TEST_METHOD(QueryRowsForwardsCountFlagsAndRowsWithoutColumnInspection)
		{
			for (const LONG requested : {1L, 20L})
			{
				auto table = TestTable(1, 1, true);
				table.expectedRequested = requested;
				table.expectedFlags = TBL_NOADVANCE;
				table.columns = {0x81AF3003};
				LPSRowSet rows = nullptr;
				Assert::AreEqual(S_OK, mapi::QueryRowsWithDiagnostics(&table, requested, TBL_NOADVANCE, &rows));
				Assert::IsNotNull(rows);
				Assert::AreEqual(1UL, rows->cRows);
				Assert::AreEqual(numberTag, rows->aRow[0].lpProps[1].ulPropTag);
				FreeProws(rows);
				Assert::AreEqual(1UL, table.queryCount);
				Assert::AreEqual(0UL, table.queryColumnsCount);
				Assert::IsTrue(m_errors.empty());
			}
		}

		TEST_METHOD(QueryRowsPreservesFailureRowOwnership)
		{
			auto table = TestTable(1, 1, true);
			table.columns = {0x81AF3003};
			table.failOnQuery = 1;
			table.queryResult = MAPI_E_NO_SUPPORT;
			table.returnRowsOnFailure = true;
			CheckQueryDiagnostic(table, true, 1);
			Assert::AreEqual(1UL, freedRowSets);
		}

		TEST_METHOD(QueryRowsPreservesWarningsWithoutMVGuidance)
		{
			auto table = TestTable(0, 20);
			table.columns = {0x81AF3003};
			table.failOnQuery = 1;
			table.queryResult = MAPI_W_PARTIAL_COMPLETION;
			CheckQueryDiagnostic(table, false);
			Assert::AreEqual(0UL, table.queryColumnsCount);
		}

		TEST_METHOD(QueryRowsReportsInvalidArguments)
		{
			auto table = TestTable(0, 20);
			LPSRowSet rows = nullptr;
			Assert::AreEqual(MAPI_E_INVALID_PARAMETER, mapi::QueryRowsWithDiagnostics(nullptr, 20, 0, &rows));
			Assert::AreEqual(MAPI_E_INVALID_PARAMETER, mapi::QueryRowsWithDiagnostics(&table, 20, 0, nullptr));
			Assert::AreEqual(size_t{2}, m_errors.size());
			Assert::AreEqual(0UL, table.queryCount);
			Assert::AreEqual(0UL, table.queryColumnsCount);
		}
		TEST_METHOD(Test_TableRowIndicesShortBatches) { CheckTable(41, 7); }
		TEST_METHOD(Test_TableRowIndicesSingleRowBatches) { CheckTable(21, 1); }

		TEST_METHOD(Test_TableRowIndicesBoundaries)
		{
			for (const auto count : {0UL, 1UL, 19UL, 20UL, 21UL, 40UL})
				CheckTable(count, 20);
		}

		TEST_METHOD(Test_TablePropertiesStayWithRowsAcrossBatches)
		{
			CheckTable(41, 7, true);
			Assert::AreEqual(size_t{82}, freedPropertyTags.size());
			for (size_t prop = 0; prop < freedPropertyTags.size(); prop += 2)
			{
				Assert::AreEqual(stringTag, freedPropertyTags[prop]);
				Assert::AreEqual(numberTag, freedPropertyTags[prop + 1]);
			}
		}

		TEST_METHOD(Test_TableSeeksToBeginning)
		{
			auto table = TestTable(3, 2, true);
			table.position = 2;
			auto file = OutputFile{};
			output::outputTable(output::dbgLevel::NoDebug, file.get(), &table);
			unittest::AreEqualEx(ExpectedTable(3, true), file.read());
			Assert::AreEqual(1UL, table.seekCount);
			Assert::IsTrue(m_errors.empty());
		}

		TEST_METHOD(Test_TableAcceptsNullTerminalRowSet)
		{
			auto table = TestTable(21, 20, true);
			table.nullAtEnd = true;
			auto file = OutputFile{};
			output::outputTable(output::dbgLevel::NoDebug, file.get(), &table);
			unittest::AreEqualEx(ExpectedTable(21, true), file.read());
			Assert::AreEqual(3UL, table.queryCount);
			Assert::AreEqual(2UL, freedRowSets);
			Assert::IsTrue(m_errors.empty());
		}

		TEST_METHOD(Test_TableQueryFailureBeforeFirstRow) { CheckQueryFailure(1, false); }
		TEST_METHOD(Test_TableQueryFailureAfterBatch) { CheckQueryFailure(2, false); }
		TEST_METHOD(Test_TableQueryFailureFreesReturnedRows) { CheckQueryFailure(2, true); }

		TEST_METHOD(Test_TableSeekFailureIsReportedAndExportContinues)
		{
			auto table = TestTable(3, 20, true);
			table.position = 1;
			table.seekResult = MAPI_E_CALL_FAILED;
			auto file = OutputFile{};
			output::outputTable(output::dbgLevel::NoDebug, file.get(), &table);
			unittest::AreEqualEx(ExpectedTable(2, true, 1), file.read());
			CheckError(L"SeekRow");
		}

		TEST_METHOD(Test_NullTableProducesNoOutput)
		{
			auto file = OutputFile{};
			output::outputTable(output::dbgLevel::NoDebug, file.get(), nullptr);
			Assert::IsTrue(file.read().empty());
			Assert::IsTrue(m_errors.empty());
		}

		TEST_METHOD(Test_DisabledDebugDoesNotReadTable)
		{
			auto table = TestTable(3, 20);
			output::outputTable(output::dbgLevel::Generic, nullptr, &table);
			Assert::AreEqual(0UL, table.queryCount);
			Assert::AreEqual(0UL, table.seekCount);
			Assert::IsTrue(m_errors.empty());
		}

		TEST_METHOD(Test_PropertiesAreSortedWithoutChangingInput)
		{
			SPropValue props[3]{};
			props[0].ulPropTag = 0x60030003;
			props[0].Value.l = 3;
			props[1].ulPropTag = numberTag;
			props[1].Value.l = 1;
			props[2].ulPropTag = 0x60020003;
			props[2].Value.l = 2;
			const auto original = std::vector<SPropValue>(std::begin(props), std::end(props));
			auto file = OutputFile{};
			output::outputProperties(output::dbgLevel::NoDebug, file.get(), 3, props, nullptr, false);
			unittest::AreEqualEx(
				ExpectedNumber(numberTag, 1) + ExpectedNumber(0x60020003, 2) + ExpectedNumber(0x60030003, 3),
				file.read());
			Assert::IsTrue(memcmp(original.data(), props, sizeof(props)) == 0);
			Assert::IsTrue(m_errors.empty());
		}

		TEST_METHOD(Test_RowSetOutputsEveryRowsProperties)
		{
			auto table = TestTable(2, 20, true);
			LPSRowSet rows{};
			Assert::AreEqual(S_OK, table.QueryRows(20, 0, &rows));
			auto file = OutputFile{};
			output::outputSRowSet(output::dbgLevel::NoDebug, file.get(), rows, nullptr);
			FreeProws(rows);
			unittest::AreEqualEx(ExpectedProperties(0) + ExpectedProperties(1), file.read());
			Assert::IsTrue(m_errors.empty());
		}

		TEST_METHOD(Test_EmptyRowSetProducesNoOutput)
		{
			SRowSet rows{};
			auto file = OutputFile{};
			output::outputSRowSet(output::dbgLevel::NoDebug, file.get(), &rows, nullptr);
			Assert::IsTrue(file.read().empty());
			Assert::IsTrue(m_errors.empty());
		}

		TEST_METHOD(Test_InvalidRowInputsProduceDiagnostics)
		{
			auto file = OutputFile{};
			output::outputSRowSet(output::dbgLevel::NoDebug, file.get(), nullptr, nullptr);
			output::outputSRow(output::dbgLevel::NoDebug, file.get(), nullptr, nullptr);
			SRow row{};
			row.cValues = 1;
			output::outputSRow(output::dbgLevel::NoDebug, file.get(), &row, nullptr);
			output::outputProperties(output::dbgLevel::NoDebug, file.get(), 1, nullptr, nullptr, false);
			const auto actual = file.read();
			for (const auto message :
				 {L"OutputSRowSet called with NULL lpRowSet!",
				  L"OutputSRow called with NULL lpSRow!",
				  L"OutputSRow called with NULL lpSRow->lpProps!",
				  L"OutputProperties called with NULL lpProps!"})
				Assert::IsTrue(actual.find(message) != std::wstring::npos);
			Assert::IsTrue(m_errors.empty());
		}

		TEST_METHOD(Test_DisabledDebugSkipsRowAndPropertyOutput)
		{
			output::outputSRowSet(output::dbgLevel::Generic, nullptr, nullptr, nullptr);
			output::outputSRow(output::dbgLevel::Generic, nullptr, nullptr, nullptr);
			output::outputProperties(output::dbgLevel::Generic, nullptr, 1, nullptr, nullptr, false);
			output::outputProperty(output::dbgLevel::Generic, nullptr, nullptr, nullptr, false);
			Assert::IsTrue(m_errors.empty());
		}

		TEST_METHOD(Test_NullPropertyProducesNoOutput)
		{
			auto file = OutputFile{};
			output::outputProperty(output::dbgLevel::NoDebug, file.get(), nullptr, nullptr, false);
			Assert::IsTrue(file.read().empty());
			Assert::IsTrue(m_errors.empty());
		}

	private:
		void CheckQueryDiagnostic(TestTable & table, bool expectGuidance, ULONG expectedRows = 0)
		{
			m_errors.clear();
			LPSRowSet rows = nullptr;
			const auto result = mapi::QueryRowsWithDiagnostics(&table, table.expectedRequested, 0, &rows);
			Assert::AreEqual(table.queryResult, result);
			if (expectedRows)
			{
				Assert::IsNotNull(rows);
				Assert::AreEqual(expectedRows, rows->cRows);
			}
			else
			{
				Assert::IsNull(rows);
			}

			FreeProws(rows);
			Assert::AreEqual(size_t{1}, m_errors.size());
			Assert::IsTrue(m_errors[0].find(L"QueryRows") != std::wstring::npos);
			Assert::IsTrue(m_errors[0].find(error::ErrorNameFromErrorCode(result)) != std::wstring::npos);
			Assert::AreEqual(expectGuidance, m_errors[0].find(L"MV_INSTANCE") != std::wstring::npos);
			if (expectGuidance)
			{
				const auto guidance = strings::loadstring(IDS_QUERYROWSMVINSTANCE);
				Assert::IsFalse(guidance.empty());
				Assert::IsTrue(m_errors[0].find(guidance) != std::wstring::npos);
				Assert::IsTrue(guidance.find(L"sort order") != std::wstring::npos);
				Assert::IsTrue(guidance.find(L"Table > Sort Table") != std::wstring::npos);
			}

			Assert::AreEqual(1UL, table.queryCount);
			Assert::AreEqual(0UL, table.setColumnsCount);
			Assert::AreEqual(0UL, table.sortCount);
			Assert::AreEqual(0UL, table.seekCount);
			Assert::IsTrue(buffers == nullptr, L"QueryRows diagnostics leaked or took ownership of MAPI buffers.");
		}

		void CheckTable(ULONG rowCount, ULONG batchSize, bool withProperties = false)
		{
			auto table = TestTable(rowCount, batchSize, withProperties);
			auto file = OutputFile{};
			output::outputTable(output::dbgLevel::NoDebug, file.get(), &table);
			unittest::AreEqualEx(
				ExpectedTable(rowCount, withProperties),
				file.read(),
				strings::format(L"Rows: %u, batch size: %u", rowCount, batchSize).c_str());
			Assert::AreEqual((rowCount + batchSize - 1) / batchSize + 1, table.queryCount);
			Assert::AreEqual(table.queryCount, freedRowSets - m_previousFreedRowSets);
			m_previousFreedRowSets = freedRowSets;
			Assert::AreEqual(1UL, table.seekCount);
			Assert::IsTrue(m_errors.empty());
		}

		void CheckQueryFailure(ULONG failOnQuery, bool returnRows)
		{
			auto table = TestTable(25, 20, true);
			table.failOnQuery = failOnQuery;
			table.returnRowsOnFailure = returnRows;
			auto file = OutputFile{};
			output::outputTable(output::dbgLevel::NoDebug, file.get(), &table);
			unittest::AreEqualEx(ExpectedTable((failOnQuery - 1) * 20, true), file.read());
			Assert::AreEqual(failOnQuery, table.queryCount);
			Assert::AreEqual(failOnQuery - (returnRows ? 0UL : 1UL), freedRowSets);
			CheckError(L"QueryRows");
		}

		void CheckError(const std::wstring& function)
		{
			Assert::AreEqual(size_t{1}, m_errors.size());
			Assert::IsTrue(m_errors[0].find(function) != std::wstring::npos);
			Assert::IsTrue(m_errors[0].find(L"0x80004005") != std::wstring::npos);
		}

		std::unique_ptr<unittest::testMapiMemory> m_memory;
		ULONG m_previousFreedRowSets{};
		std::function<void(const std::wstring&)> m_originalDisplayError;
		std::vector<std::wstring> m_errors;
	};
} // namespace mapioutputtest
