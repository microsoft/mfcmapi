#pragma once
#include <UI/Dialogs/Editors/Editor.h>
#include <core/mapi/cache/mapiObjects.h>

namespace dialog::editor
{
	class CRestrictEditor : public CEditor
	{
	public:
		CRestrictEditor(
			_In_ CWnd* pParentWnd,
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			_In_opt_ LPVOID lpAllocParent,
			_In_opt_ const _SRestriction* lpRes);
		~CRestrictEditor();

		_Check_return_ LPSRestriction DetachModifiedSRestriction() noexcept;

	private:
		void OnEditAction1() override;
		HRESULT EditCompare(const _SRestriction* lpSourceRes);
		HRESULT EditAndOr(const _SRestriction* lpSourceRes);
		HRESULT EditRestrict(const _SRestriction* lpSourceRes);
		HRESULT EditCombined(const _SRestriction* lpSourceRes);
		HRESULT EditBitmask(const _SRestriction* lpSourceRes);
		HRESULT EditSize(const _SRestriction* lpSourceRes);
		HRESULT EditExist(const _SRestriction* lpSourceRes);
		HRESULT EditSubrestriction(const _SRestriction* lpSourceRes);
		HRESULT EditComment(const _SRestriction* lpSourceRes);
		BOOL OnInitDialog() override;
		_Check_return_ ULONG HandleChange(UINT nID) override;
		void OnOK() override;

		_Check_return_ const _SRestriction* GetSourceRes() const noexcept;

		// source variables
		const _SRestriction* m_lpRes;
		LPVOID m_lpAllocParent;

		// output variable
		LPSRestriction m_lpOutputRes;
		bool m_bModified;
		std::shared_ptr<cache::CMapiObjects> m_lpMapiObjects{};
	};

	class CCriteriaEditor : public CEditor
	{
	public:
		CCriteriaEditor(
			_In_ CWnd* pParentWnd,
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			_In_ const _SRestriction* lpRes,
			_In_ LPENTRYLIST lpEntryList,
			ULONG ulSearchState);
		~CCriteriaEditor();

		_Check_return_ LPSRestriction DetachModifiedSRestriction() noexcept;
		_Check_return_ LPENTRYLIST DetachModifiedEntryList() noexcept;
		_Check_return_ ULONG GetSearchFlags() const noexcept;
		_Check_return_ bool DoListEdit(ULONG ulListNum, int iItem, _In_ sortlistdata::sortListData* lpData) override;

	private:
		void OnEditAction1() override;
		BOOL OnInitDialog() override;
		_Check_return_ ULONG HandleChange(UINT nID) override;
		void InitListFromEntryList(ULONG ulListNum, _In_ const SBinaryArray* lpEntryList) const;
		void OnOK() override;

		_Check_return_ const _SRestriction* GetSourceRes() const noexcept;

		const _SRestriction* m_lpSourceRes;
		LPSRestriction m_lpNewRes;

		LPENTRYLIST m_lpSourceEntryList;
		LPENTRYLIST m_lpNewEntryList{};

		ULONG m_ulNewSearchFlags;
		std::shared_ptr<cache::CMapiObjects> m_lpMapiObjects{};
	};
} // namespace dialog::editor