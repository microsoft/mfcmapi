#include <StdAfx.h>
#include <UI/Controls/SortList/SortListData.h>
#include <core/sortlistdata/contentsData.h>
#include <UI/Controls/SortList/NodeData.h>
#include <core/sortlistdata/propListData.h>
#include <core/sortlistdata/mvPropData.h>
#include <core/sortlistdata/resData.h>
#include <core/sortlistdata/commentData.h>
#include <core/sortlistdata/binaryData.h>
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

		contentsData* SortListData::Contents() const { return reinterpret_cast<contentsData*>(m_lpData); }

		NodeData* SortListData::Node() const { return reinterpret_cast<NodeData*>(m_lpData); }

		propListData* SortListData::Prop() const { return reinterpret_cast<propListData*>(m_lpData); }

		mvPropData* SortListData::MV() const { return reinterpret_cast<mvPropData*>(m_lpData); }

		resData* SortListData::Res() const { return reinterpret_cast<resData*>(m_lpData); }

		commentData* SortListData::Comment() const { return reinterpret_cast<commentData*>(m_lpData); }

		binaryData* SortListData::Binary() const { return reinterpret_cast<binaryData*>(m_lpData); }

		void SortListData::setSortText(const std::wstring& _sortText) { sortText = strings::wstringToLower(_sortText); }

		// Sets data from the LPSRow into the SortListData structure
		// Assumes the structure is either an existing structure or a new one which has been memset to 0
		// If it's an existing structure - we need to free up some memory
		// SORTLIST_CONTENTS
		void SortListData::InitializeContents(_In_ LPSRow lpsRowData)
		{
			Clean();

			if (!lpsRowData) return;
			lpSourceProps = lpsRowData->lpProps;
			cSourceProps = lpsRowData->cValues;
			m_lpData = new (std::nothrow) contentsData(lpsRowData);
		}

		void SortListData::InitializeNode(
			ULONG cProps,
			_In_opt_ LPSPropValue lpProps,
			_In_opt_ LPSBinary lpEntryID,
			_In_opt_ LPSBinary lpInstanceKey,
			ULONG bSubfolders,
			ULONG ulContainerFlags)
		{
			Clean();
			cSourceProps = cProps;
			lpSourceProps = lpProps;
			m_lpData = new (std::nothrow) NodeData(lpEntryID, lpInstanceKey, bSubfolders, ulContainerFlags);
		}

		void SortListData::InitializeNode(_In_ LPSRow lpsRow)
		{
			Clean();
			if (!lpsRow) return;

			LPSBinary lpEIDBin = nullptr; // don't free
			LPSBinary lpInstanceBin = nullptr; // don't free

			auto lpEID = PpropFindProp(lpsRow->lpProps, lpsRow->cValues, PR_ENTRYID);
			if (lpEID) lpEIDBin = &lpEID->Value.bin;

			auto lpInstance = PpropFindProp(lpsRow->lpProps, lpsRow->cValues, PR_INSTANCE_KEY);
			if (lpInstance) lpInstanceBin = &lpInstance->Value.bin;

			const auto lpSubfolders = PpropFindProp(lpsRow->lpProps, lpsRow->cValues, PR_SUBFOLDERS);

			const auto lpContainerFlags = PpropFindProp(lpsRow->lpProps, lpsRow->cValues, PR_CONTAINER_FLAGS);

			InitializeNode(
				lpsRow->cValues,
				lpsRow->lpProps, // pass on props to be archived in node
				lpEIDBin,
				lpInstanceBin,
				lpSubfolders ? static_cast<ULONG>(lpSubfolders->Value.b) : MAPI_E_NOT_FOUND,
				lpContainerFlags ? lpContainerFlags->Value.ul : MAPI_E_NOT_FOUND);
		}

		void SortListData::InitializePropList(_In_ ULONG ulPropTag)
		{
			Clean();
			bItemFullyLoaded = true;
			m_lpData = new (std::nothrow) propListData(ulPropTag);
		}

		void SortListData::InitializeMV(_In_ const _SPropValue* lpProp, ULONG iProp)
		{
			Clean();
			m_lpData = new (std::nothrow) mvPropData(lpProp, iProp);
		}

		void SortListData::InitializeMV(_In_opt_ const _SPropValue* lpProp)
		{
			Clean();
			m_lpData = new (std::nothrow) mvPropData(lpProp);
		}

		void SortListData::InitializeRes(_In_opt_ const _SRestriction* lpOldRes)
		{
			Clean();
			bItemFullyLoaded = true;
			m_lpData = new (std::nothrow) resData(lpOldRes);
		}

		void SortListData::InitializeComment(_In_opt_ const _SPropValue* lpOldProp)
		{
			Clean();
			bItemFullyLoaded = true;
			m_lpData = new (std::nothrow) commentData(lpOldProp);
		}

		void SortListData::InitializeBinary(_In_opt_ LPSBinary lpOldBin)
		{
			Clean();
			m_lpData = new (std::nothrow) binaryData(lpOldBin);
		}
	} // namespace sortlistdata
} // namespace controls