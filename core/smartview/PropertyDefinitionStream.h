#pragma once
#include <core/smartview/SmartViewParser.h>

namespace smartview
{
	struct PackedUnicodeString
	{
		blockT<BYTE> cchLength;
		blockT<WORD> cchExtendedLength;
		blockStringW szCharacters;
	};

	struct PackedAnsiString
	{
		blockT<BYTE> cchLength;
		blockT<WORD> cchExtendedLength;
		blockStringA szCharacters;
	};

	struct SkipBlock
	{
		blockT<DWORD> dwSize;
		blockBytes lpbContent;
		PackedUnicodeString lpbContentText;
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
		std::vector<SkipBlock> psbSkipBlocks;
	};

	class PropertyDefinitionStream : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<WORD> m_wVersion;
		blockT<DWORD> m_dwFieldDefinitionCount;
		std::vector<FieldDefinition> m_pfdFieldDefinitions;
	};
} // namespace smartview