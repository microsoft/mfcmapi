#pragma once
#include "SmartViewParser.h"

struct PackedUnicodeString
{
	BYTE cchLength;
	WORD cchExtendedLength;
	std::wstring szCharacters;
};

struct PackedAnsiString
{
	BYTE cchLength;
	WORD cchExtendedLength;
	std::string szCharacters;
};

struct SkipBlock
{
	DWORD dwSize;
	std::vector<BYTE> lpbContent;
};

struct FieldDefinition
{
	DWORD dwFlags;
	WORD wVT;
	DWORD dwDispid;
	WORD wNmidNameLength;
	std::wstring szNmidName;
	PackedAnsiString pasNameANSI;
	PackedAnsiString pasFormulaANSI;
	PackedAnsiString pasValidationRuleANSI;
	PackedAnsiString pasValidationTextANSI;
	PackedAnsiString pasErrorANSI;
	DWORD dwInternalType;
	DWORD dwSkipBlockCount;
	std::vector<SkipBlock> psbSkipBlocks;
};

class PropertyDefinitionStream : public SmartViewParser
{
public:
	PropertyDefinitionStream();

private:
	void Parse() override;
	_Check_return_ std::wstring ToStringInternal() override;

	WORD m_wVersion;
	DWORD m_dwFieldDefinitionCount;
	std::vector<FieldDefinition> m_pfdFieldDefinitions;
};