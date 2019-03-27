#pragma once
#include <core/smartview/smartViewParser.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	struct FolderFieldDefinitionCommon
	{
		std::shared_ptr<blockT<GUID>> PropSetGuid = emptyT<GUID>();
		std::shared_ptr<blockT<DWORD>> fcapm = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> dwString = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> dwBitmap = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> dwDisplay = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> iFmt = emptyT<DWORD>();
		std::shared_ptr<blockT<WORD>> wszFormulaLength = emptyT<WORD>();
		std::shared_ptr<blockStringW> wszFormula = emptySW();

		void parse(const std::shared_ptr<binaryParser>& parser);
	};

	struct FolderFieldDefinitionA
	{
		std::shared_ptr<blockT<DWORD>> FieldType = emptyT<DWORD>();
		std::shared_ptr<blockT<WORD>> FieldNameLength = emptyT<WORD>();
		std::shared_ptr<blockStringA> FieldName = emptySA();
		FolderFieldDefinitionCommon Common;

		FolderFieldDefinitionA(const std::shared_ptr<binaryParser>& parser);
	};

	struct FolderFieldDefinitionW
	{
		std::shared_ptr<blockT<DWORD>> FieldType = emptyT<DWORD>();
		std::shared_ptr<blockT<WORD>> FieldNameLength = emptyT<WORD>();
		std::shared_ptr<blockStringW> FieldName = emptySW();
		FolderFieldDefinitionCommon Common;

		FolderFieldDefinitionW(const std::shared_ptr<binaryParser>& parser);
	};

	class FolderUserFieldStream : public smartViewParser
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<DWORD>> m_FolderUserFieldsAnsiCount = emptyT<DWORD>();
		std::vector<std::shared_ptr<FolderFieldDefinitionA>> m_FieldDefinitionsA;
		std::shared_ptr<blockT<DWORD>> m_FolderUserFieldsUnicodeCount = emptyT<DWORD>();
		std::vector<std::shared_ptr<FolderFieldDefinitionW>> m_FieldDefinitionsW;
	};
} // namespace smartview