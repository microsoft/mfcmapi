#pragma once
#include <core/sortlistdata/data.h>

namespace controls
{
	namespace sortlistdata
	{
		class contentsData;
		class NodeData;
		class propListData;
		class mvPropData;
		class resData;
		class commentData;
		class binaryData;

		class SortListData
		{
		public:
			~SortListData();
			void Clean();
			void InitializeContents(_In_ LPSRow lpsRowData);
			void InitializeNode(
				ULONG cProps,
				_In_opt_ LPSPropValue lpProps,
				_In_opt_ LPSBinary lpEntryID,
				_In_opt_ LPSBinary lpInstanceKey,
				ULONG bSubfolders,
				ULONG ulContainerFlags);
			void InitializeNode(_In_ LPSRow lpsRow);
			void InitializePropList(_In_ ULONG ulPropTag);
			void InitializeMV(_In_ const _SPropValue* lpProp, ULONG iProp);
			void InitializeMV(_In_opt_ const _SPropValue* lpProp);
			void InitializeRes(_In_opt_ const _SRestriction* lpOldRes);
			void InitializeComment(_In_opt_ const _SPropValue* lpOldProp);
			void InitializeBinary(_In_opt_ LPSBinary lpOldBin);

			template <typename T> T* cast() { return reinterpret_cast<T*>(m_lpData); }
			NodeData* Node() const { return reinterpret_cast<NodeData*>(m_lpData); }
			commentData* Comment() const { return reinterpret_cast<commentData*>(m_lpData); }
			binaryData* Binary() const { return reinterpret_cast<binaryData*>(m_lpData); }

			const std::wstring& getSortText() const noexcept { return sortText; }
			void setSortText(const std::wstring& _sortText);
			const ULARGE_INTEGER& getSortValue() const noexcept { return sortValue; }
			void setSortValue(const ULARGE_INTEGER _sortValue) noexcept { sortValue = _sortValue; }
			void clearSortValues() noexcept
			{
				sortText.clear();
				sortValue = {};
			}

			ULONG cSourceProps{};
			LPSPropValue
				lpSourceProps{}; // Stolen from lpsRowData in SortListData::InitializeContents - free with MAPIFreeBuffer
			bool bItemFullyLoaded{};

		private:
			IData* m_lpData{};
			std::wstring sortText{};
			ULARGE_INTEGER sortValue{};
		};
	} // namespace sortlistdata
} // namespace controls