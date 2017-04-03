#pragma once
#include <Dialogs/Editors/Editor.h>

#define NOSKIPFIELD ((ULONG)0xffffffff)

class CSearchEditor : public CEditor
{
public:
	enum SearchFields
	{
		PROPNAME,
		SEARCHTERM,
		FUZZYLEVEL,
		FINDROW,
	};

	CSearchEditor(
		ULONG ulPropTag,
		_In_opt_ LPMAPIPROP lpMAPIProp,
		_In_ CWnd* pParentWnd);
	virtual ~CSearchEditor();

	_Check_return_ LPSRestriction GetRestriction();

private:
	_Check_return_ ULONG HandleChange(UINT nID) override;
	void OnEditAction1() override;
	void PopulateFields(ULONG ulSkipField) const;
	_Check_return_ ULONG GetSelectedFuzzyLevel() const;

	LPMAPIPROP m_lpMAPIProp;
	ULONG m_ulPropTag;
	ULONG m_ulFuzzyLevel;
};