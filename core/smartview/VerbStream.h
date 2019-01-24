#pragma once
#include <core/smartview/SmartViewParser.h>

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
	};

	struct VerbExtraData
	{
		blockT<BYTE> DisplayNameCount;
		blockStringW DisplayName;
		blockT<BYTE> DisplayNameCountRepeat;
		blockStringW DisplayNameRepeat;
	};

	class VerbStream : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<WORD> m_Version;
		blockT<DWORD> m_Count;
		std::vector<VerbData> m_lpVerbData;
		blockT<WORD> m_Version2;
		std::vector<VerbExtraData> m_lpVerbExtraData;
	};
} // namespace smartview