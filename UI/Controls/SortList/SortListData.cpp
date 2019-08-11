#include <StdAfx.h>
#include <UI/Controls/SortList/SortListData.h>
#include <core/utility/strings.h>

namespace controls
{
	namespace sortlistdata
	{
		SortListData::~SortListData() { Clean(); }

		void SortListData::Clean()
		{
			delete m_lpData;
			m_lpData = nullptr;

			MAPIFreeBuffer(lpSourceProps);
			lpSourceProps = nullptr;
			cSourceProps = 0;

			bItemFullyLoaded = false;
			clearSortValues();
		}

		void SortListData::setSortText(const std::wstring& _sortText) { sortText = strings::wstringToLower(_sortText); }
	} // namespace sortlistdata
} // namespace controls