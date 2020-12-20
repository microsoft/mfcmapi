#pragma once
#include <core/sortlistdata/data.h>

namespace sortlistdata
{
	class sortListData;

	class contentsData : public IData
	{
	public:
		static void init(sortListData* lpData, _In_ LPSRow lpsRowData);

		contentsData(_In_ LPSRow lpsRowData);
		~contentsData();

		_Check_return_ LPSBinary getEntryID() { return m_lpEntryID; }
		_Check_return_ LPSBinary getLongTermID() { return m_lpLongtermID; }
		_Check_return_ LPSBinary getInstanceKey() { return m_lpInstanceKey; }
		_Check_return_ LPSBinary getServiceUID() { return m_lpServiceUID; }
		_Check_return_ LPSBinary getProviderUID() { return m_lpProviderUID; }
		_Check_return_ std::wstring getDN() { return m_szDN; }
		_Check_return_ std::wstring getProfileDisplayName() { return m_szProfileDisplayName; }
		_Check_return_ ULONG getAttachNum() { return m_ulAttachNum; }
		_Check_return_ ULONG getAttachMethod() { return m_ulAttachMethod; }
		_Check_return_ ULONG getRowID() { return m_ulRowID; }
		_Check_return_ ULONG getRowType() { return m_ulRowType; }
		void setRowType(ULONG rowType) { m_ulRowType = rowType; }

	private:
		LPSBinary m_lpEntryID{}; // Allocated with MAPIAllocateBuffer
		LPSBinary m_lpLongtermID{}; // Allocated with MAPIAllocateBuffer
		LPSBinary m_lpInstanceKey{}; // Allocated with MAPIAllocateBuffer
		LPSBinary m_lpServiceUID{}; // Allocated with MAPIAllocateBuffer
		LPSBinary m_lpProviderUID{}; // Allocated with MAPIAllocateBuffer
		std::wstring m_szDN{};
		std::wstring m_szProfileDisplayName{};
		ULONG m_ulAttachNum{};
		ULONG m_ulAttachMethod{};
		ULONG m_ulRowID{}; // for recipients
		ULONG m_ulRowType{}; // PR_ROW_TYPE
	}; // namespace sortlistdata
} // namespace sortlistdata