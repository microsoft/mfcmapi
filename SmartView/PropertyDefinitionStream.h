#pragma once
#include "SmartViewParser.h"

struct PackedUnicodeString
{
	BYTE cchLength;
	WORD cchExtendedLength;
	LPWSTR szCharacters;
};

struct PackedAnsiString
{
	BYTE cchLength;
	WORD cchExtendedLength;
	LPSTR szCharacters;
};

struct SkipBlock
{
	DWORD dwSize;
	BYTE* lpbContent;
};

struct FieldDefinition
{
	DWORD dwFlags;
	WORD wVT;
	DWORD dwDispid;
	WORD wNmidNameLength;
	LPWSTR szNmidName;
	PackedAnsiString pasNameANSI;
	PackedAnsiString pasFormulaANSI;
	PackedAnsiString pasValidationRuleANSI;
	PackedAnsiString pasValidationTextANSI;
	PackedAnsiString pasErrorANSI;
	DWORD dwInternalType;
	DWORD dwSkipBlockCount;
	SkipBlock* psbSkipBlocks;
};

class PropertyDefinitionStream : public SmartViewParser
{
public:
	PropertyDefinitionStream();
	~PropertyDefinitionStream();

private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	WORD m_wVersion;
	DWORD m_dwFieldDefinitionCount;
	FieldDefinition* m_pfdFieldDefinitions;
};