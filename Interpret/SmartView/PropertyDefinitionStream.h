#pragma once
#include "SmartViewParser.h"

struct PackedUnicodeString
{
	BYTE cchLength;
	WORD cchExtendedLength;
	wstring szCharacters;
};

struct PackedAnsiString
{
	BYTE cchLength;
	WORD cchExtendedLength;
	string szCharacters;
};

struct SkipBlock
{
	DWORD dwSize;
	vector<BYTE> lpbContent;
};

struct FieldDefinition
{
	DWORD dwFlags;
	WORD wVT;
	DWORD dwDispid;
	WORD wNmidNameLength;
	wstring szNmidName;
	PackedAnsiString pasNameANSI;
	PackedAnsiString pasFormulaANSI;
	PackedAnsiString pasValidationRuleANSI;
	PackedAnsiString pasValidationTextANSI;
	PackedAnsiString pasErrorANSI;
	DWORD dwInternalType;
	DWORD dwSkipBlockCount;
	vector<SkipBlock> psbSkipBlocks;
};

class PropertyDefinitionStream : public SmartViewParser
{
public:
	PropertyDefinitionStream();

private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	WORD m_wVersion;
	DWORD m_dwFieldDefinitionCount;
	vector<FieldDefinition> m_pfdFieldDefinitions;
};