#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockStringW.h>

namespace smartview
{
	struct PackedUnicodeString
	{
		blockT<BYTE> cchLength;
		blockT<WORD> cchExtendedLength;
		blockStringW szCharacters;

		void parse(std::shared_ptr<binaryParser>& parser);
	};

	struct PackedAnsiString
	{
		blockT<BYTE> cchLength;
		blockT<WORD> cchExtendedLength;
		blockStringA szCharacters;

		void parse(std::shared_ptr<binaryParser>& parser);
	};

	struct SkipBlock
	{
		blockT<DWORD> dwSize;
		blockBytes lpbContent;
		PackedUnicodeString lpbContentText;

		SkipBlock(std::shared_ptr<binaryParser>& parser, DWORD iSkip);
	};

	struct FieldDefinition
	{
		blockT<DWORD> dwFlags;
		blockT<WORD> wVT;
		blockT<DWORD> dwDispid;
		blockT<WORD> wNmidNameLength;
		blockStringW szNmidName;
		PackedAnsiString pasNameANSI;
		PackedAnsiString pasFormulaANSI;
		PackedAnsiString pasValidationRuleANSI;
		PackedAnsiString pasValidationTextANSI;
		PackedAnsiString pasErrorANSI;
		blockT<DWORD> dwInternalType;
		DWORD dwSkipBlockCount{};
		std::vector<std::shared_ptr<SkipBlock>> psbSkipBlocks;

		FieldDefinition(std::shared_ptr<binaryParser>& parser, WORD version);
	};

	class PropertyDefinitionStream : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<WORD> m_wVersion;
		blockT<DWORD> m_dwFieldDefinitionCount;
		std::vector<std::shared_ptr<FieldDefinition>> m_pfdFieldDefinitions;
	};
} // namespace smartview