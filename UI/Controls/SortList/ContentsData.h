#pragma once
#include <UI/Controls/SortList/Data.h>

namespace controls
{
	namespace sortlistdata
	{
		class ContentsData : public IData
		{
		public:
			ContentsData(_In_ LPSRow lpsRowData);
			~ContentsData();
			LPSBinary m_lpEntryID; // Allocated with MAPIAllocateBuffer
			LPSBinary m_lpLongtermID; // Allocated with MAPIAllocateBuffer
			LPSBinary m_lpInstanceKey; // Allocated with MAPIAllocateBuffer
			LPSBinary m_lpServiceUID; // Allocated with MAPIAllocateBuffer
			LPSBinary m_lpProviderUID; // Allocated with MAPIAllocateBuffer
			std::wstring m_szDN;
			std::string m_szProfileDisplayName;
			ULONG m_ulAttachNum;
			ULONG m_ulAttachMethod;
			ULONG m_ulRowID; // for recipients
			ULONG m_ulRowType; // PR_ROW_TYPE
		};
	}
}