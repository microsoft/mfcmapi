#pragma once
#include <Interpret/SmartView/SmartViewParser.h>

namespace smartview
{
	struct FolderFieldDefinitionCommon
	{
		GUID PropSetGuid;
		DWORD fcapm;
		DWORD dwString;
		DWORD dwBitmap;
		DWORD dwDisplay;
		DWORD iFmt;
		WORD wszFormulaLength;
		std::wstring wszFormula;
	};

	struct FolderFieldDefinitionA
	{
		DWORD FieldType;
		WORD FieldNameLength;
		std::string FieldName;
		FolderFieldDefinitionCommon Common;
	};

	struct FolderFieldDefinitionW
	{
		DWORD FieldType;
		WORD FieldNameLength;
		std::wstring FieldName;
		FolderFieldDefinitionCommon Common;
	};

	class FolderUserFieldStream : public SmartViewParser
	{
	public:
		FolderUserFieldStream();

	private:
		void Parse() override;
		_Check_return_ std::wstring ToStringInternal() override;

		FolderFieldDefinitionCommon BinToFolderFieldDefinitionCommon();

		DWORD m_FolderUserFieldsAnsiCount;
		std::vector<FolderFieldDefinitionA> m_FieldDefinitionsA;
		DWORD m_FolderUserFieldsUnicodeCount;
		std::vector<FolderFieldDefinitionW> m_FieldDefinitionsW;
	};
}