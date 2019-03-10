#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockStringW.h>

namespace smartview
{
	struct VerbData
	{
		blockT<DWORD> VerbType;
		blockT<BYTE> DisplayNameCount;
		blockStringA DisplayName;
		blockT<BYTE> MsgClsNameCount;
		blockStringA MsgClsName;
		blockT<BYTE> Internal1StringCount;
		blockStringA Internal1String;
		blockT<BYTE> DisplayNameCountRepeat;
		blockStringA DisplayNameRepeat;
		blockT<DWORD> Internal2;
		blockT<BYTE> Internal3;
		blockT<DWORD> fUseUSHeaders;
		blockT<DWORD> Internal4;
		blockT<DWORD> SendBehavior;
		blockT<DWORD> Internal5;
		blockT<DWORD> ID;
		blockT<DWORD> Internal6;

		VerbData(std::shared_ptr<binaryParser>& parser);
	};

	struct VerbExtraData
	{
		blockT<BYTE> DisplayNameCount;
		blockStringW DisplayName;
		blockT<BYTE> DisplayNameCountRepeat;
		blockStringW DisplayNameRepeat;

		VerbExtraData(std::shared_ptr<binaryParser>& parser);
	};

	class VerbStream : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<WORD> m_Version;
		blockT<DWORD> m_Count;
		std::vector<std::shared_ptr<VerbData>> m_lpVerbData;
		blockT<WORD> m_Version2;
		std::vector<std::shared_ptr<VerbExtraData>> m_lpVerbExtraData;
	};
} // namespace smartview