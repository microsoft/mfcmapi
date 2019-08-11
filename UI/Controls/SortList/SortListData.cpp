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

		void SortListData::setSortText(const std::wstring& _sortText) { sortText = strings::wstringToLower(_sortText); }

		// Sets data from the LPSRow into the SortListData structure
		// Assumes the structure is either an existing structure or a new one which has been memset to 0
		// If it's an existing structure - we need to free up some memory
		// SORTLIST_CONTENTS
		void InitContents(SortListData* data, _In_ LPSRow lpsRowData)
		{
			if (!data) return;

			if (lpsRowData)
			{
				data->Init(new (std::nothrow) contentsData(lpsRowData), lpsRowData->cValues, lpsRowData->lpProps);
			}
			else
			{
				data->Clean();
			}
		}

		void InitNode(
			SortListData* data,
			ULONG cProps,
			_In_opt_ LPSPropValue lpProps,
			_In_opt_ LPSBinary lpEntryID,
			_In_opt_ LPSBinary lpInstanceKey,
			ULONG bSubfolders,
			ULONG ulContainerFlags)
		{
			if (!data) return;
			data->Init(
				new (std::nothrow) NodeData(lpEntryID, lpInstanceKey, bSubfolders, ulContainerFlags), cProps, lpProps);
		}

		void InitNode(SortListData* data, _In_ LPSRow lpsRow)
		{
			if (!data) return;

			if (!lpsRow)
			{
				data->Clean();
				return;
			}

			LPSBinary lpEIDBin = nullptr; // don't free
			LPSBinary lpInstanceBin = nullptr; // don't free

			auto lpEID = PpropFindProp(lpsRow->lpProps, lpsRow->cValues, PR_ENTRYID);
			if (lpEID) lpEIDBin = &lpEID->Value.bin;

			auto lpInstance = PpropFindProp(lpsRow->lpProps, lpsRow->cValues, PR_INSTANCE_KEY);
			if (lpInstance) lpInstanceBin = &lpInstance->Value.bin;

			const auto lpSubfolders = PpropFindProp(lpsRow->lpProps, lpsRow->cValues, PR_SUBFOLDERS);

			const auto lpContainerFlags = PpropFindProp(lpsRow->lpProps, lpsRow->cValues, PR_CONTAINER_FLAGS);

			InitNode(
				data,
				lpsRow->cValues,
				lpsRow->lpProps, // pass on props to be archived in node
				lpEIDBin,
				lpInstanceBin,
				lpSubfolders ? static_cast<ULONG>(lpSubfolders->Value.b) : MAPI_E_NOT_FOUND,
				lpContainerFlags ? lpContainerFlags->Value.ul : MAPI_E_NOT_FOUND);
		}

		void InitPropList(SortListData* data, _In_ ULONG ulPropTag)
		{
			if (!data) return;
			data->Init(new (std::nothrow) propListData(ulPropTag), true);
		}

		void InitMV(SortListData* data, _In_ const _SPropValue* lpProp, ULONG iProp)
		{
			if (!data) return;
			data->Init(new (std::nothrow) mvPropData(lpProp, iProp));
		}

		void InitMV(SortListData* data, _In_opt_ const _SPropValue* lpProp)
		{
			if (!data) return;
			data->Init(new (std::nothrow) mvPropData(lpProp));
		}

		void InitRes(SortListData* data, _In_opt_ const _SRestriction* lpOldRes)
		{
			if (!data) return;
			data->Init(new (std::nothrow) resData(lpOldRes), true);
		}

		void InitComment(SortListData* data, _In_opt_ const _SPropValue* lpOldProp)
		{
			if (!data) return;
			data->Init(new (std::nothrow) commentData(lpOldProp), true);
		}

		void InitBinary(SortListData* data, _In_opt_ LPSBinary lpOldBin)
		{
			if (!data) return;
			data->Init(new (std::nothrow) binaryData(lpOldBin));
		}
	} // namespace sortlistdata
} // namespace controls