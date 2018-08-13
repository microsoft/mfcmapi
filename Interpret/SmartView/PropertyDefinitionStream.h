#pragma once
#include <Interpret/SmartView/SmartViewParser.h>

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
		DWORD dwSkipBlockCount;
		std::vector<SkipBlock> psbSkipBlocks;
	};

	class PropertyDefinitionStream : public SmartViewParser
	{
	public:
		PropertyDefinitionStream();

	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<WORD> m_wVersion;
		blockT<DWORD> m_dwFieldDefinitionCount;
		std::vector<FieldDefinition> m_pfdFieldDefinitions;
	};
}