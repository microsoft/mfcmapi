#pragma once
#include "SmartViewParser.h"

struct FolderFieldDefinitionCommon
{
	GUID PropSetGuid;
	DWORD fcapm;
	DWORD dwString;
	DWORD dwBitmap;
	DWORD dwDisplay;
	DWORD iFmt;
	WORD wszFormulaLength;
	wstring wszFormula;
};

struct FolderFieldDefinitionA
{
	DWORD FieldType;
	WORD FieldNameLength;
	string FieldName;
	FolderFieldDefinitionCommon Common;
};

struct FolderFieldDefinitionW
{
	DWORD FieldType;
	WORD FieldNameLength;
	wstring FieldName;
	FolderFieldDefinitionCommon Common;
};

class FolderUserFieldStream : public SmartViewParser
{
public:
	FolderUserFieldStream();

private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	FolderFieldDefinitionCommon BinToFolderFieldDefinitionCommon();

	DWORD m_FolderUserFieldsAnsiCount;
	vector<FolderFieldDefinitionA> m_FieldDefinitionsA;
	DWORD m_FolderUserFieldsUnicodeCount;
	vector<FolderFieldDefinitionW> m_FieldDefinitionsW;
};