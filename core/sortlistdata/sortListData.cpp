#include <core/stdafx.h>
#include <core/sortlistdata/sortListData.h>
#include <core/utility/strings.h>

namespace sortlistdata
{
	void sortListData::clean() noexcept
	{
		lpData = nullptr;

		MAPIFreeBuffer(lpSourceProps);
		lpSourceProps = nullptr;
		cSourceProps = 0;

		bItemFullyLoaded = false;
		clearSortValues();
	}

	void sortListData::setSortText(const std::wstring& _sortText) { sortText = strings::wstringToLower(_sortText); }
} // namespace sortlistdata