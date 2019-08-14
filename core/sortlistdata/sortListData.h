#pragma once
#include <core/sortlistdata/data.h>

namespace sortlistdata
{
	class sortListData
	{
	public:
		void init(IData* lpData, ULONG cValues, LPSPropValue lpProps)
		{
			clean();
			lpSourceProps = lpProps;
			cSourceProps = cValues;
			m_lpData = lpData;
		}

		void init(IData* lpData, bool _bItemFullyLoaded = false)
		{
			clean();
			bItemFullyLoaded = _bItemFullyLoaded;
			m_lpData = lpData;
		}
		~sortListData();
		void clean();

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