#include "stdafx.h"
#include "BinaryParser.h"

static wstring CLASS = L"CBinaryParser";

CBinaryParser::CBinaryParser()
{
	m_cbBin = 0;
	m_lpBin = nullptr;
	m_lpCur = nullptr;
	m_lpEnd = nullptr;
}

CBinaryParser::CBinaryParser(size_t cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	Init(cbBin, lpBin);
}

void CBinaryParser::Init(size_t cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	DebugPrintEx(DBGSmartView, CLASS, L"Init", L"cbBin = 0x%08X = %u\n", static_cast<int>(cbBin), static_cast<UINT>(cbBin));
	m_cbBin = cbBin;
	m_lpBin = lpBin;
	m_lpCur = lpBin;
	m_lpEnd = lpBin + cbBin;
}

bool CBinaryParser::Empty() const
{
	return m_cbBin == 0 || m_lpBin == nullptr;
}

void CBinaryParser::Advance(size_t cbAdvance)
{
	m_lpCur += cbAdvance;
}

void CBinaryParser::Rewind()
{
	DebugPrintEx(DBGSmartView, CLASS, L"Rewind", L"Rewinding to the beginning of the stream\n");
	m_lpCur = m_lpBin;
}

size_t CBinaryParser::GetCurrentOffset() const
{
	return m_lpCur - m_lpBin;
}

LPBYTE CBinaryParser::GetCurrentAddress() const
{
	return m_lpCur;
}

void CBinaryParser::SetCurrentOffset(size_t stOffset)
{
	DebugPrintEx(DBGSmartView, CLASS, L"SetCurrentOffset", L"Setting offset 0x%08X = %u bytes.\n", static_cast<int>(stOffset), static_cast<UINT>(stOffset));
	m_lpCur = m_lpBin + stOffset;
}

// If we're before the end of the buffer, return the count of remaining bytes
// If we're at or past the end of the buffer, return 0
// If we're before the beginning of the buffer, return 0
size_t CBinaryParser::RemainingBytes() const
{
	if (m_lpCur < m_lpBin || m_lpCur > m_lpEnd) return 0;
	return m_lpEnd - m_lpCur;
}

bool CBinaryParser::CheckRemainingBytes(size_t cbBytes) const
{
	if (!m_lpCur)
	{
		DebugPrintEx(DBGSmartView, CLASS, L"CheckRemainingBytes", L"Current offset does not exist!\n");
		return false;
	}

	auto cbRemaining = RemainingBytes();
	if (cbBytes > cbRemaining)
	{
		DebugPrintEx(DBGSmartView, CLASS, L"CheckRemainingBytes", L"Bytes requested (0x%08X = %u) > remaining bytes (0x%08X = %u)\n",
			static_cast<int>(cbBytes), static_cast<UINT>(cbBytes),
			static_cast<int>(cbRemaining), static_cast<UINT>(cbRemaining));
		DebugPrintEx(DBGSmartView, CLASS, L"CheckRemainingBytes", L"Total Bytes: 0x%08X = %u\n", static_cast<int>(m_cbBin), static_cast<UINT>(m_cbBin));
		DebugPrintEx(DBGSmartView, CLASS, L"CheckRemainingBytes", L"Current offset: 0x%08X = %d\n", static_cast<int>(m_lpCur - m_lpBin), static_cast<int>(m_lpCur - m_lpBin));
		return false;
	}

	return true;
}

void CBinaryParser::GetBYTE(_Out_ BYTE* pBYTE)
{
	if (!pBYTE) return;
	*pBYTE = NULL;
	if (!CheckRemainingBytes(sizeof(BYTE))) return;
	*pBYTE = *static_cast<BYTE*>(m_lpCur);
	m_lpCur += sizeof(BYTE);
}

void CBinaryParser::GetWORD(_Out_ WORD* pWORD)
{
	if (!pWORD) return;
	*pWORD = NULL;
	if (!CheckRemainingBytes(sizeof(WORD))) return;
	*pWORD = *reinterpret_cast<WORD*>(m_lpCur);
	m_lpCur += sizeof(WORD);
}

void CBinaryParser::GetDWORD(_Out_ DWORD* pDWORD)
{
	if (!pDWORD) return;
	*pDWORD = NULL;
	if (!CheckRemainingBytes(sizeof(DWORD))) return;
	*pDWORD = *reinterpret_cast<DWORD*>(m_lpCur);
	m_lpCur += sizeof(DWORD);
}

void CBinaryParser::GetLARGE_INTEGER(_Out_ LARGE_INTEGER* pLARGE_INTEGER)
{
	if (!pLARGE_INTEGER) return;
	*pLARGE_INTEGER = LARGE_INTEGER();
	if (!CheckRemainingBytes(sizeof(LARGE_INTEGER))) return;
	*pLARGE_INTEGER = *reinterpret_cast<LARGE_INTEGER*>(m_lpCur);
	m_lpCur += sizeof(LARGE_INTEGER);
}

void CBinaryParser::GetBYTESNoAlloc(size_t cbBytes, size_t cbMaxBytes, _In_count_(cbBytes) LPBYTE pBYTES)
{
	if (!cbBytes || !pBYTES || !CheckRemainingBytes(cbBytes)) return;
	if (cbBytes > cbMaxBytes) return;
	memset(pBYTES, 0, sizeof(BYTE)* cbBytes);
	memcpy(pBYTES, m_lpCur, cbBytes);
	m_lpCur += cbBytes;
}

string CBinaryParser::GetStringA(size_t cchChar)
{
	if (cchChar == -1)
	{
		auto hRes = StringCchLengthA(reinterpret_cast<LPCSTR>(m_lpCur), (m_lpEnd - m_lpCur) / sizeof CHAR, &cchChar);
		if (FAILED(hRes)) return "";
		cchChar += 1;
	}

	if (!cchChar || !CheckRemainingBytes(sizeof CHAR * cchChar)) return "";
	string ret(reinterpret_cast<LPCSTR>(m_lpCur), cchChar);
	m_lpCur += sizeof CHAR * cchChar;
	return ret;
}

wstring CBinaryParser::GetStringW(size_t cchChar)
{
	if (cchChar == -1)
	{
		auto hRes = StringCchLengthW(reinterpret_cast<LPCWSTR>(m_lpCur), (m_lpEnd - m_lpCur) / sizeof WCHAR, &cchChar);
		if (FAILED(hRes)) return emptystring;
		cchChar += 1;
	}

	if (!cchChar || !CheckRemainingBytes(sizeof WCHAR * cchChar)) return emptystring;
	wstring ret(reinterpret_cast<LPCWSTR>(m_lpCur), cchChar);
	m_lpCur += sizeof WCHAR * cchChar;
	return ret;
}

vector<BYTE> CBinaryParser::GetBYTES(size_t cbBytes, size_t cbMaxBytes)
{
	if (!cbBytes || !CheckRemainingBytes(cbBytes)) return vector<BYTE>();
	if (cbMaxBytes != -1 && cbBytes > cbMaxBytes) return vector<BYTE>();
	vector<BYTE> ret(reinterpret_cast<LPBYTE>(m_lpCur), reinterpret_cast<LPBYTE>(m_lpCur + cbBytes));
	m_lpCur += cbBytes;
	return ret;
}

vector<BYTE> CBinaryParser::GetRemainingData()
{
	return GetBYTES(RemainingBytes());
}