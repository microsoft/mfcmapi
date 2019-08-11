#pragma once
#include <core/sortlistdata/data.h>

namespace controls
{
	namespace sortlistdata
	{
		class SortListData
		{
		public:
			~SortListData();
			void Clean();
			void Init(IData* lpData, ULONG cValues, LPSPropValue lpProps)
			{
				Clean();
				lpSourceProps = lpProps;
				cSourceProps = cValues;
				m_lpData = lpData;
			}

			void Init(IData* lpData, bool _bItemFullyLoaded = false)
			{
				Clean();
				bItemFullyLoaded = _bItemFullyLoaded;  
				m_lpData = lpData;
			}

			template <typename T> T* cast() { return reinterpret_cast<T*>(m_lpData); }

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

		void InitContents(SortListData* lpData, _In_ LPSRow lpsRowData);
		void InitNode(
			SortListData* lpData,
			ULONG cProps,
			_In_opt_ LPSPropValue lpProps,
			_In_opt_ LPSBinary lpEntryID,
			_In_opt_ LPSBinary lpInstanceKey,
			ULONG bSubfolders,
			ULONG ulContainerFlags);
		void InitNode(SortListData* lpData, _In_ LPSRow lpsRow);
		void InitPropList(SortListData* data, _In_ ULONG ulPropTag);
		void InitMV(SortListData* data, _In_ const _SPropValue* lpProp, ULONG iProp);
		void InitMV(SortListData* data, _In_opt_ const _SPropValue* lpProp);
		void InitRes(SortListData* data, _In_opt_ const _SRestriction* lpOldRes);
		void InitComment(SortListData* data, _In_opt_ const _SPropValue* lpOldProp);
		void InitBinary(SortListData* data, _In_opt_ LPSBinary lpOldBin);
	} // namespace sortlistdata
} // namespace controls