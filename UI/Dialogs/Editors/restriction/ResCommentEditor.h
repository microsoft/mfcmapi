#pragma once
#include <UI/Dialogs/Editors/Editor.h>

namespace dialog::editor
{
	// This class is only invoked by RestrictEditor. RestrictEditor always passes an alloc parent.
	// So all memory detached from this class is owned by a parent and must not be freed manually
	static std::wstring COMMENTCLASS = L"ResCommentEditor"; // STRING_OK
	class ResCommentEditor : public CEditor
	{
	public:
		ResCommentEditor(
			_In_ CWnd* pParentWnd,
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			_In_ const _SRestriction* lpRes,
			_In_ LPVOID lpAllocParent);

		_Check_return_ LPSRestriction DetachModifiedSRestriction() noexcept;
		_Check_return_ LPSPropValue DetachModifiedSPropValue() noexcept;
		_Check_return_ ULONG GetSPropValueCount() const noexcept;
		_Check_return_ bool DoListEdit(ULONG ulListNum, int iItem, _In_ sortlistdata::sortListData* lpData) override;

	private:
		void OnEditAction1() override;
		BOOL OnInitDialog() override;
		void InitListFromPropArray(ULONG ulListNum, ULONG cProps, _In_count_(cProps) const _SPropValue* lpProps) const;
		_Check_return_ const _SRestriction* GetSourceRes() const noexcept;
		void OnOK() override;

		LPVOID m_lpAllocParent;
		const _SRestriction* m_lpSourceRes;

		LPSRestriction m_lpNewCommentRes;
		LPSPropValue m_lpNewCommentProp;
		ULONG m_ulNewCommentProp{};
		std::shared_ptr<cache::CMapiObjects> m_lpMapiObjects{};
	};
} // namespace dialog::editor