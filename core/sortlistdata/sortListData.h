#pragma once
#include <core/sortlistdata/data.h>

namespace controls
{
	namespace sortlistdata
	{
		class sortListData
		{
		public:
			~sortListData();
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
				lpSourceProps{}; // Stolen from lpsRowData in sortListData::InitializeContents - free with MAPIFreeBuffer
			bool bItemFullyLoaded{};

		private:
			IData* m_lpData{};
			std::wstring sortText{};
			ULARGE_INTEGER sortValue{};
		};
	} // namespace sortlistdata
} // namespace controls