#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	class PackedUnicodeString : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<BYTE>> cchLength = emptyT<BYTE>();
		std::shared_ptr<blockT<WORD>> cchExtendedLength = emptyT<WORD>();
		std::shared_ptr<blockStringW> szCharacters = emptySW();
	};

	class PackedAnsiString : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<BYTE>> cchLength = emptyT<BYTE>();
		std::shared_ptr<blockT<WORD>> cchExtendedLength = emptyT<WORD>();
		std::shared_ptr<blockStringA> szCharacters = emptySA();
	};

	class SkipBlock : public block
	{
	public:
		SkipBlock(DWORD _iSkip) : iSkip(_iSkip){};

	private:
		void parse() override;
		void parseBlocks() override;

		DWORD iSkip{};
		std::shared_ptr<blockT<DWORD>> dwSize = emptyT<DWORD>();
		std::shared_ptr<blockBytes> lpbContent = emptyBB();
		std::shared_ptr<PackedUnicodeString> lpbContentText;
	};

	class FieldDefinition : public block
	{
	public:
		FieldDefinition(WORD _version) : version(_version){};

	private:
		void parse() override;
		void parseBlocks() override;

		WORD version{};

		std::shared_ptr<blockT<DWORD>> dwFlags = emptyT<DWORD>();
		std::shared_ptr<blockT<WORD>> wVT = emptyT<WORD>();
		std::shared_ptr<blockT<DWORD>> dwDispid = emptyT<DWORD>();
		std::shared_ptr<blockT<WORD>> wNmidNameLength = emptyT<WORD>();
		std::shared_ptr<blockStringW> szNmidName = emptySW();
		std::shared_ptr<PackedAnsiString> pasNameANSI;
		std::shared_ptr<PackedAnsiString> pasFormulaANSI;
		std::shared_ptr<PackedAnsiString> pasValidationRuleANSI;
		std::shared_ptr<PackedAnsiString> pasValidationTextANSI;
		std::shared_ptr<PackedAnsiString> pasErrorANSI;
		std::shared_ptr<blockT<DWORD>> dwInternalType = emptyT<DWORD>();
		std::vector<std::shared_ptr<SkipBlock>> psbSkipBlocks;
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