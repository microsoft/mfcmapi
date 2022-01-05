#pragma once
#include <UI/Dialogs/Editors/Editor.h>
#include <core/mapi/cache/mapiObjects.h>

namespace dialog::editor
{
	class RestrictEditor : public CEditor
	{
	public:
		RestrictEditor(
			_In_ CWnd* pParentWnd,
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			_In_opt_ LPVOID lpAllocParent,
			_In_opt_ const _SRestriction* lpRes);
		~RestrictEditor();

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
} // namespace dialog::editor