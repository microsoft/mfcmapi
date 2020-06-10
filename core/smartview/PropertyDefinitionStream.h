#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	struct PackedUnicodeString
	{
		std::shared_ptr<blockT<BYTE>> cchLength = emptyT<BYTE>();
		std::shared_ptr<blockT<WORD>> cchExtendedLength = emptyT<WORD>();
		std::shared_ptr<blockStringW> szCharacters = emptySW();

		void parse(const std::shared_ptr<binaryParser>& parser);
		std::shared_ptr<block> toBlock(_In_ const std::wstring& szFieldName);
	};

	struct PackedAnsiString
	{
		std::shared_ptr<blockT<BYTE>> cchLength = emptyT<BYTE>();
		std::shared_ptr<blockT<WORD>> cchExtendedLength = emptyT<WORD>();
		std::shared_ptr<blockStringA> szCharacters = emptySA();

		void parse(const std::shared_ptr<binaryParser>& parser);
		std::shared_ptr<block> toBlock(_In_ const std::wstring& szFieldName);
	};

	struct SkipBlock
	{
		std::shared_ptr<blockT<DWORD>> dwSize = emptyT<DWORD>();
		std::shared_ptr<blockBytes> lpbContent = emptyBB();
		PackedUnicodeString lpbContentText;

		SkipBlock(const std::shared_ptr<binaryParser>& parser, DWORD iSkip);
	};

	struct FieldDefinition
	{
		std::shared_ptr<blockT<DWORD>> dwFlags = emptyT<DWORD>();
		std::shared_ptr<blockT<WORD>> wVT = emptyT<WORD>();
		std::shared_ptr<blockT<DWORD>> dwDispid = emptyT<DWORD>();
		std::shared_ptr<blockT<WORD>> wNmidNameLength = emptyT<WORD>();
		std::shared_ptr<blockStringW> szNmidName = emptySW();
		PackedAnsiString pasNameANSI;
		PackedAnsiString pasFormulaANSI;
		PackedAnsiString pasValidationRuleANSI;
		PackedAnsiString pasValidationTextANSI;
		PackedAnsiString pasErrorANSI;
		std::shared_ptr<blockT<DWORD>> dwInternalType = emptyT<DWORD>();
		std::vector<std::shared_ptr<SkipBlock>> psbSkipBlocks;

		FieldDefinition(const std::shared_ptr<binaryParser>& parser, WORD version);
	};

	class PropertyDefinitionStream : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<WORD>> m_wVersion = emptyT<WORD>();
		std::shared_ptr<blockT<DWORD>> m_dwFieldDefinitionCount = emptyT<DWORD>();
		std::vector<std::shared_ptr<FieldDefinition>> m_pfdFieldDefinitions;
	};
} // namespace smartview