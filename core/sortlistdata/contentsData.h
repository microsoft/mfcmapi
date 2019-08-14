#pragma once
#include <core/sortlistdata/data.h>

namespace sortlistdata
{
	class sortListData;

	void InitContents(sortListData* lpData, _In_ LPSRow lpsRowData);

	class contentsData : public IData
	{
	public:
		contentsData(_In_ LPSRow lpsRowData);
		~contentsData();
		LPSBinary m_lpEntryID{}; // Allocated with MAPIAllocateBuffer
		LPSBinary m_lpLongtermID{}; // Allocated with MAPIAllocateBuffer
		LPSBinary m_lpInstanceKey{}; // Allocated with MAPIAllocateBuffer
		LPSBinary m_lpServiceUID{}; // Allocated with MAPIAllocateBuffer
		LPSBinary m_lpProviderUID{}; // Allocated with MAPIAllocateBuffer
		std::wstring m_szDN{};
		std::string m_szProfileDisplayName{};
		ULONG m_ulAttachNum{};
		ULONG m_ulAttachMethod{};
		ULONG m_ulRowID{}; // for recipients
		ULONG m_ulRowType{}; // PR_ROW_TYPE
	};
} // namespace sortlistdata