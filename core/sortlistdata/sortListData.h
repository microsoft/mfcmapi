#pragma once
#include <core/sortlistdata/data.h>

namespace sortlistdata
{
	class sortListData
	{
	public:
		void init(std::shared_ptr<IData> _lpData, ULONG cValues, LPSPropValue lpProps)
		{
			clean();
			lpSourceProps = lpProps;
			cSourceProps = cValues;
			lpData = _lpData;
		}

		void init(std::shared_ptr<IData> _lpData, bool _bItemFullyLoaded = false)
		{
			clean();
			bItemFullyLoaded = _bItemFullyLoaded;
			lpData = _lpData;
		}
		~sortListData();
		void clean() noexcept;

		template <typename T> std::shared_ptr<T> cast() { return std::dynamic_pointer_cast<T>(lpData); }

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
		std::shared_ptr<IData> lpData{};
		std::wstring sortText{};
		ULARGE_INTEGER sortValue{};
	};
} // namespace sortlistdata