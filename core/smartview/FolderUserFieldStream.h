#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	class FolderFieldDefinition : public block
	{
	public:
		FolderFieldDefinition::FolderFieldDefinition(bool _unicode) : unicode(_unicode) {}

	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<DWORD>> FieldType = emptyT<DWORD>();
		std::shared_ptr<blockT<WORD>> FieldNameLength = emptyT<WORD>();
		std::shared_ptr<blockStringA> FieldNameA = emptySA();
		std::shared_ptr<blockStringW> FieldNameW = emptySW();
		std::shared_ptr<blockT<GUID>> PropSetGuid = emptyT<GUID>();
		std::shared_ptr<blockT<DWORD>> fcapm = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> dwString = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> dwBitmap = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> dwDisplay = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> iFmt = emptyT<DWORD>();
		std::shared_ptr<blockT<WORD>> wszFormulaLength = emptyT<WORD>();
		std::shared_ptr<blockStringW> wszFormula = emptySW();

		bool unicode{};
	};

	class FolderUserFieldStream : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<DWORD>> m_FolderUserFieldsAnsiCount = emptyT<DWORD>();
		std::vector<std::shared_ptr<FolderFieldDefinition>> m_FieldDefinitionsA;
		std::shared_ptr<blockT<DWORD>> m_FolderUserFieldsUnicodeCount = emptyT<DWORD>();
		std::vector<std::shared_ptr<FolderFieldDefinition>> m_FieldDefinitionsW;
	};
} // namespace smartview