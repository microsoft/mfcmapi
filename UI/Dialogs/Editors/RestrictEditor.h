#pragma once
#include <UI/Dialogs/Editors/Editor.h>

class CRestrictEditor : public CEditor
{
public:
	CRestrictEditor(
		_In_ CWnd* pParentWnd,
		_In_opt_ LPVOID lpAllocParent,
		_In_opt_ const _SRestriction* lpRes);
	virtual ~CRestrictEditor();

	_Check_return_ LPSRestriction DetachModifiedSRestriction();

private:
	void OnEditAction1() override;
	HRESULT EditCompare(LPSRestriction lpSourceRes);
	HRESULT EditAndOr(LPSRestriction lpSourceRes);
	HRESULT EditRestrict(LPSRestriction lpSourceRes);
	HRESULT EditCombined(LPSRestriction lpSourceRes);
	HRESULT EditBitmask(LPSRestriction lpSourceRes);
	HRESULT EditSize(LPSRestriction lpSourceRes);
	HRESULT EditExist(LPSRestriction lpSourceRes);
	HRESULT EditSubrestriction(LPSRestriction lpSourceRes);
	HRESULT EditComment(LPSRestriction lpSourceRes);
	BOOL OnInitDialog() override;
	_Check_return_ ULONG HandleChange(UINT nID) override;
	void OnOK() override;

	_Check_return_ LPSRestriction GetSourceRes() const;

	// source variables
	const _SRestriction* m_lpRes;
	LPVOID m_lpAllocParent;

	// output variable
	LPSRestriction m_lpOutputRes;
	bool m_bModified;
};


class CCriteriaEditor : public CEditor
{
public:
	CCriteriaEditor(
		_In_ CWnd* pParentWnd,
		_In_ LPSRestriction lpRes,
		_In_ LPENTRYLIST lpEntryList,
		ULONG ulSearchState);
	virtual ~CCriteriaEditor();

	_Check_return_ LPSRestriction DetachModifiedSRestriction();
	_Check_return_ LPENTRYLIST DetachModifiedEntryList();
	_Check_return_ ULONG GetSearchFlags() const;
	_Check_return_ bool DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData) override;

private:
	void OnEditAction1() override;
	BOOL OnInitDialog() override;
	_Check_return_ ULONG HandleChange(UINT nID) override;
	void InitListFromEntryList(ULONG ulListNum, _In_ LPENTRYLIST lpEntryList) const;
	void OnOK() override;

	_Check_return_ LPSRestriction GetSourceRes() const;

	LPSRestriction m_lpSourceRes;
	LPSRestriction m_lpNewRes;

	LPENTRYLIST m_lpSourceEntryList;
	LPENTRYLIST m_lpNewEntryList{};

	ULONG m_ulNewSearchFlags;
};