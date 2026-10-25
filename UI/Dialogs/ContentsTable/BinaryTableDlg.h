#pragma once
#include <UI/Dialogs/ContentsTable/ContentsTableDlg.h>
#include <core/mapi/mapiBinaryFunctions.h>

namespace dialog
{
	_Check_return_ HRESULT
	DisplayBinaryTable(_In_ LPMAPIPROP message, ULONG propTag, bool entryIDs, bool isAB, _In_ CBaseDialog* parent);
	_Check_return_ HRESULT DisplayBinaryTable(const SPropValue& property, bool entryIDs, _In_ CBaseDialog* parent);

	class CBinaryTableDlg : public CContentsTableDlg
	{
	public:
		CBinaryTableDlg(
			std::shared_ptr<cache::CMapiObjects> objects,
			_In_opt_ LPMAPIPROP message,
			ULONG propTag,
			bool entryIDs,
			_In_ LPMAPITABLE table,
			std::vector<mapi::binary::BinaryTableRow> members,
			bool isAB = false);
		~CBinaryTableDlg();

	private:
		BOOL OnInitDialog() override;
		_Check_return_ LPMAPIPROP OpenItemProp(int item, modifyType modify) override;
		void OnDisplayItem() override;
		void OnRefreshView() override;

		LPMAPIPROP m_message{};
		bool m_entryIDs{};
		ULONG m_propTag{};
		std::vector<mapi::binary::BinaryTableRow> m_members;
	};
} // namespace dialog
