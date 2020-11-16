#pragma once
#include <UI/Controls/SortList/SortHeader.h>
#include <core/sortlistdata/sortListData.h>

namespace controls::sortlistctrl
{
	// This enum maps to icons.bmp
	// Reorder at your own risk!!!
	enum class sortIcon
	{
		siDefault = 0,
		nodeCollapsed,
		nodeExpanded,
		ptUnspecified,
		ptNull,
		ptI2,
		ptLong,
		ptR4,
		ptDouble,
		ptCurrency,
		ptAppTime,
		ptError,
		ptBoolean,
		ptObject,
		ptI8,
		ptString8,
		ptUnicode,
		ptSysTime,
		ptClsid,
		ptBinary,
		mvI2,
		mvLong,
		mvR4,
		mvDouble,
		mvCurrency,
		mvAppTime,
		mvSysTime,
		mvString8,
		mvBinary,
		mvUnicode,
		mvClsid,
		mvI8,
		sRestriction,
		actions,
		mapiStore,
		mapiAddrBook,
		mapiFolder,
		mapiAbcont,
		mapiMessage,
		mapiMailUser,
		mapiAttach,
		mapiDistList,
		mapiProfSect,
		mapiStatus,
		mapiSession,
		mapiFormInfo,
	};

	struct TypeIcon
	{
		ULONG objType{};
		sortIcon image{};
	};

	class CSortListCtrl : public CListCtrl
	{
	public:
		CSortListCtrl();
		~CSortListCtrl();

		STDMETHODIMP_(ULONG) AddRef();
		STDMETHODIMP_(ULONG) Release();

		// Exported manipulation functions
		void Create(_In_ CWnd* pCreateParent, ULONG ulFlags, UINT nID, bool bImages);
		void AutoSizeColumns(bool bMinWidth);
		void DeleteAllColumns(bool bShutdown = false);
		void SetSelectedItem(int iItem);
		void SortClickedColumn();
		_Check_return_ sortlistdata::sortListData* InsertRow(int iRow, const std::wstring& szText) const;
		void SetItemText(int nItem, int nSubItem, const std::wstring& lpszText);
		std::wstring GetItemText(_In_ int nItem, _In_ int nSubItem) const;
		void AllowEscapeClose() noexcept;
		int InsertColumnW(_In_ int nCol, const std::wstring& columnHeading) noexcept;

	protected:
		void MySetRedraw(bool bRedraw);
		_Check_return_ sortlistdata::sortListData*
		InsertRow(int iRow, const std::wstring& szText, int iIndent, sortIcon iImage) const;
		void FakeClickColumn(int iColumn, bool bSortUp);

		// protected since derived classes need to call the base implementation
		LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;

	private:
		// Overrides from base class
		_Check_return_ UINT OnGetDlgCode();

		void OnColumnClick(int iColumn);
		void OnDeleteAllItems(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult) noexcept;
		void OnDeleteItem(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
		void AutoSizeColumn(int iColumn, int iMaxWidth, int iMinWidth);
		void OnCustomDraw(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult) noexcept;
		_Check_return_ static int CALLBACK
		MyCompareProc(_In_ LPARAM lParam1, _In_ LPARAM lParam2, _In_ LPARAM lParamSort) noexcept;

		LONG m_cRef;
		int m_iRedrawCount;
		CImageList m_ImageList;
		bool m_bHaveSorted;
		bool m_bHeaderSubclassed;
		CSortHeader m_cSortHeader;
		int m_iClickedColumn;
		bool m_bSortUp;
		int m_iItemCurHover;
		bool m_bAllowEscapeClose;

		DECLARE_MESSAGE_MAP()
	};
} // namespace controls::sortlistctrl