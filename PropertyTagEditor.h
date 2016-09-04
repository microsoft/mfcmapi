#pragma once
#include "Editor.h"

#define NOSKIPFIELD ((ULONG)0xffffffff)

class CPropertyTagEditor : public CEditor
{
public:
	CPropertyTagEditor(
		UINT uidTitle,
		UINT uidPrompt,
		ULONG ulPropTag,
		bool bIsAB,
		_In_opt_ LPMAPIPROP lpMAPIProp,
		_In_ CWnd* pParentWnd);
	virtual ~CPropertyTagEditor();

	_Check_return_ ULONG GetPropertyTag() const;

private:
	_Check_return_ ULONG HandleChange(UINT nID) override;
	void OnEditAction1() override;
	void OnEditAction2() override;
	BOOL OnInitDialog() override;
	void PopulateFields(ULONG ulSkipField);
	_Check_return_ ULONG GetSelectedPropType();
	void LookupNamedProp(ULONG ulSkipField, bool bCreate);
	_Check_return_ wstring GetDropStringUseControl(ULONG iControl);
	_Check_return_ int GetDropDownSelection(ULONG iControl);
	void InsertDropString(ULONG iControl, int iRow, _In_ wstring szText);
	void SetDropDownSelection(ULONG i, _In_ wstring szText);

	ULONG m_ulPropTag;
	bool m_bIsAB;
	LPMAPIPROP m_lpMAPIProp;
};

class CPropertySelector : public CEditor
{
public:
	CPropertySelector(
		bool bIncludeABProps,
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ CWnd* pParentWnd);
	virtual ~CPropertySelector();

	_Check_return_ ULONG GetPropertyTag() const;

private:
	BOOL OnInitDialog() override;
	_Check_return_ bool DoListEdit(ULONG ulListNum, int iItem, _In_ SortListData* lpData) override;
	void OnOK() override;
	_Check_return_ SortListData* GetSelectedListRowData(ULONG iControl);

	ULONG m_ulPropTag;
	bool m_bIncludeABProps;
	LPMAPIPROP m_lpMAPIProp;
};