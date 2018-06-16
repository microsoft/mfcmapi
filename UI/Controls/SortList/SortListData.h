#pragma once
#include <UI/Controls/SortList/Data.h>
namespace controls
{
	namespace sortlistdata
	{
		class ContentsData;
		class NodeData;
		class PropListData;
		class MVPropData;
		class ResData;
		class CommentData;
		class BinaryData;

		class SortListData
		{
		public:
			SortListData();
			~SortListData();
			void Clean();
			void InitializeContents(_In_ LPSRow lpsRowData);
			void InitializeNode(
				ULONG cProps,
				_In_opt_ LPSPropValue lpProps,
				_In_opt_ LPSBinary lpEntryID,
				_In_opt_ LPSBinary lpInstanceKey,
				ULONG bSubfolders,
				ULONG ulContainerFlags);
			void InitializeNode(_In_ LPSRow lpsRow);
			void InitializePropList(_In_ ULONG ulPropTag);
			void InitializeMV(_In_ const _SPropValue* lpProp, ULONG iProp);
			void InitializeMV(_In_opt_ const _SPropValue* lpProp);
			void InitializeRes(_In_ const _SRestriction* lpOldRes);
			void InitializeComment(_In_ const _SPropValue* lpOldProp);
			void InitializeBinary(_In_ LPSBinary lpOldBin);

			std::wstring szSortText;
			ULARGE_INTEGER ulSortValue{};

			ContentsData* Contents() const;
			NodeData* Node() const;
			PropListData* Prop() const;
			MVPropData* MV() const;
			ResData* Res() const;
			CommentData* Comment() const;
			BinaryData* Binary() const;

			ULONG cSourceProps;
			LPSPropValue
				lpSourceProps; // Stolen from lpsRowData in SortListData::InitializeContents - free with MAPIFreeBuffer
			bool bItemFullyLoaded;

		private:
			IData* m_lpData;
		};
	}
}