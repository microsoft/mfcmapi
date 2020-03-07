#include <core/stdafx.h>
#include <core/sortlistdata/nodeData.h>
#include <core/mapi/mapiFunctions.h>
#include <core/sortlistdata/sortListData.h>
#include <core/utility/output.h>
#include <core/mapi/adviseSink.h>
#include <core/utility/error.h>
#include <core/utility/registry.h>

namespace sortlistdata
{
	void nodeData::init(sortListData* data, ULONG cProps, _In_opt_ LPSPropValue lpProps)
	{
		if (!data) return;

		if (!cProps || !lpProps)
		{
			data->clean();
			return;
		}

		data->init(std::make_shared<nodeData>(cProps, lpProps), cProps, lpProps);
	}

	nodeData::nodeData(ULONG cProps, _In_opt_ LPSPropValue lpProps) { init(cProps, lpProps); }

	nodeData::~nodeData()
	{
		unadvise();

		if (m_lpHierarchyTable) m_lpHierarchyTable->Release();
	}

	// Builds/rebuilds a node given a row without touching the hierarchy table or advise sink
	// Can be used for construction or modification of row
	void nodeData::init(ULONG cProps, _In_opt_ LPSPropValue lpProps)
	{
		m_lpEntryID.clear();
		m_lpInstanceKey.clear();

		if (!cProps || !lpProps) return;
		const auto lpEID = PpropFindProp(lpProps, cProps, PR_ENTRYID);
		if (lpEID)
		{
			const auto bin = mapi::getBin(lpEID);
			m_lpEntryID = {bin.lpb, bin.lpb + bin.cb};
		}

		const auto lpInstance = PpropFindProp(lpProps, cProps, PR_INSTANCE_KEY);
		if (lpInstance)
		{
			const auto bin = mapi::getBin(lpInstance);
			m_lpInstanceKey = {bin.lpb, bin.lpb + bin.cb};
		}

		const auto lpSubfolders = PpropFindProp(lpProps, cProps, PR_SUBFOLDERS);
		const auto lpContainerFlags = PpropFindProp(lpProps, cProps, PR_CONTAINER_FLAGS);

		if (lpSubfolders)
		{
			m_cSubfolders = lpSubfolders->Value.b != 0 ? 1 : 0;
		}
		else if (lpContainerFlags)
		{
			m_cSubfolders = lpContainerFlags->Value.ul & AB_SUBCONTAINERS ? 1 : 0;
		}
	}

	bool nodeData::advise(HWND m_hWnd, HTREEITEM hItem, LPMDB lpMDB)
	{
		if (!m_lpHierarchyTable) return false;
		auto lpAdviseSink = new (std::nothrow) mapi::adviseSink(m_hWnd, hItem);

		if (lpAdviseSink)
		{
			const auto hRes = WC_MAPI(m_lpHierarchyTable->Advise(
				fnevTableModified, static_cast<IMAPIAdviseSink*>(lpAdviseSink), &m_ulAdviseConnection));
			if (hRes == MAPI_E_NO_SUPPORT) // Some tables don't support this!
			{
				lpAdviseSink->Release();
				lpAdviseSink = nullptr;
				output::DebugPrint(output::dbgLevel::Notify, L"This table doesn't support notifications\n");
				return false;
			}
			else if (hRes == S_OK && lpMDB)
			{
				lpAdviseSink->SetAdviseTarget(lpMDB);
				mapi::ForceRop(lpMDB);
			}

			output::DebugPrint(
				output::dbgLevel::Notify,
				L"nodeData::advise",
				L"Advise sink %p, ulAdviseConnection = 0x%08X\n",
				lpAdviseSink,
				static_cast<int>(m_ulAdviseConnection));

			m_lpAdviseSink = lpAdviseSink;
			return true;
		}

		return false;
	}

	void nodeData::unadvise()
	{
		if (m_lpAdviseSink)
		{
			output::DebugPrint(
				output::dbgLevel::Hierarchy,
				L"nodeData::unadvise",
				L"Unadvising %p, ulAdviseConnection = 0x%08X\n",
				m_lpAdviseSink,
				static_cast<int>(m_ulAdviseConnection));

			// unadvise before releasing our sink
			if (m_lpAdviseSink && m_lpHierarchyTable)
			{
				m_lpHierarchyTable->Unadvise(m_ulAdviseConnection);
			}

			if (m_lpAdviseSink)
			{
				m_lpAdviseSink->Release();
				m_lpAdviseSink = nullptr;
			}
		}
	}

	bool nodeData::hasChildren()
	{
		// If we don't have a table, then we don't have children
		if (!m_lpHierarchyTable) return false;
		ULONG ulRowCount = NULL;
		const auto hRes = WC_MAPI(m_lpHierarchyTable->GetRowCount(NULL, &ulRowCount));
		if (S_OK != hRes || ulRowCount)
		{
			// If we have a row count, we have children.
			// If we can't get one, assume we might have children.
			return true;
		}

		return false;
	}

	bool nodeData::matchInstanceKey(_In_opt_ const SBinary* lpInstanceKey)
	{
		if (!lpInstanceKey) return m_lpInstanceKey.empty();
		return m_lpInstanceKey == std::vector<BYTE>{lpInstanceKey->lpb, lpInstanceKey->lpb + lpInstanceKey->cb};
	}

} // namespace sortlistdata