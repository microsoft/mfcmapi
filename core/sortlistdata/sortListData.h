#pragma once
#include <core/sortlistdata/data.h>
#include <core/mapi/mapiFunctions.h>

namespace sortlistdata
{
	class sortListData
	{
	public:
		void init(std::shared_ptr<IData> _lpData, ULONG cValues, LPSPropValue lpProps) noexcept
		{
			clean();
			lpSourceProps = lpProps;
			cSourceProps = cValues;
			lpData = _lpData;
		}

		void init(std::shared_ptr<IData> _lpData, bool _bItemFullyLoaded = false) noexcept
		{
			clean();
			bItemFullyLoaded = _bItemFullyLoaded;
			lpData = _lpData;
		}
		~sortListData() { clean(); }
		void clean() noexcept;

		template <typename T> std::shared_ptr<T> cast() noexcept { return std::dynamic_pointer_cast<T>(lpData); }

		_Check_return_ SRow getRow() { return {0, cSourceProps, lpSourceProps}; }
		void setRow(_In_ ULONG cProps, _In_ SPropValue* lpProps)
		{
			MAPIFreeBuffer(lpSourceProps);
			cSourceProps = cProps;
			lpSourceProps = lpProps;
		}

		_Check_return_ bool getFullyLoaded() noexcept { return bItemFullyLoaded; }
		void setFullyLoaded(_In_ const bool _fullyLoaded) noexcept { bItemFullyLoaded = _fullyLoaded; }

		const std::wstring& getSortText() const noexcept { return sortText; }
		void setSortText(const std::wstring& _sortText);
		const ULARGE_INTEGER& getSortValue() const noexcept { return sortValue; }
		void setSortValue(const ULARGE_INTEGER _sortValue) noexcept { sortValue = _sortValue; }
		void clearSortValues() noexcept
		{
			sortText.clear();
			sortValue = {};
		}

		_Check_return_ SPropValue* GetOneProp(const ULONG ulPropTag)
		{
			return mapi::FindProp(lpSourceProps, cSourceProps, ulPropTag);
		}

	private:
		std::shared_ptr<IData> lpData{};
		std::wstring sortText{};
		ULARGE_INTEGER sortValue{};

		ULONG cSourceProps{};
		LPSPropValue
			lpSourceProps{}; // Stolen from lpsRowData in sortListData::InitializeContents - free with MAPIFreeBuffer
		bool bItemFullyLoaded{};
	};
} // namespace sortlistdata