#pragma once
#include <core/sortlistdata/data.h>

namespace mapi
{
	class adviseSink;
} // namespace mapi

namespace sortlistdata
{
	class sortListData;

	class nodeData : public IData
	{
	public:
		static void init(
			sortListData* lpData,
			ULONG cProps,
			_In_opt_ LPSPropValue lpProps,
			_In_opt_ const SBinary* lpEntryID,
			_In_opt_ const SBinary* lpInstanceKey,
			ULONG bSubfolders,
			ULONG ulContainerFlags);
		static void init(sortListData* lpData, _In_ LPSRow lpsRow);

		nodeData(
			_In_opt_ const SBinary* lpEntryID,
			_In_opt_ const SBinary* lpInstanceKey,
			ULONG bSubfolders,
			ULONG ulContainerFlags);
		~nodeData();

		bool hasSink() noexcept { return !!m_lpAdviseSink; }
		mapi::adviseSink* getSink() noexcept { return m_lpAdviseSink; }
		bool advise(HWND m_hWnd, HTREEITEM hItem, LPMDB lpMDB);
		bool hasChildren();
		LONG getSubfolders() noexcept { return m_cSubfolders; }
		void setSubfolders(bool hasFolders) noexcept { m_cSubfolders = hasFolders ? 1 : 0; }
		LPMAPITABLE getTable() noexcept { return m_lpHierarchyTable; }
		void setTable(LPMAPITABLE table) noexcept { m_lpHierarchyTable = table; } // we assume the caller did AddRef
		bool hasTable() noexcept { return !!m_lpHierarchyTable; }
		bool matchInstanceKey(_In_opt_ const SBinary* lpInstanceKey);

		LPSBinary m_lpEntryID{}; // Allocated with MAPIAllocateBuffer

	private:
		void unadvise();

		mapi::adviseSink* m_lpAdviseSink{}; // Object - free with Release
		ULONG_PTR m_ulAdviseConnection{};
		LPMAPITABLE m_lpHierarchyTable{}; // Object - free with Release
		LONG m_cSubfolders{-1}; // -1 for unknown, 0 for no subfolders, >0 for at least one subfolder
		std::vector<BYTE> m_lpInstanceKey{};
	};
} // namespace sortlistdata