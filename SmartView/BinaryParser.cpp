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

void CBinaryParser::GetBYTES(size_t cbBytes, size_t cbMaxBytes, _Out_ LPBYTE* ppBYTES)
{
	if (ppBYTES) *ppBYTES = nullptr;
	if (!cbBytes || !ppBYTES || !CheckRemainingBytes(cbBytes)) return;
	if (cbBytes > cbMaxBytes) return;
	*ppBYTES = new BYTE[cbBytes];
	if (*ppBYTES)
	{
		memset(*ppBYTES, 0, sizeof(BYTE)* cbBytes);
		memcpy(*ppBYTES, m_lpCur, cbBytes);
	}

	m_lpCur += cbBytes;
}

void CBinaryParser::GetBYTESNoAlloc(size_t cbBytes, size_t cbMaxBytes, _In_count_(cbBytes) LPBYTE pBYTES)
{
	if (!cbBytes || !pBYTES || !CheckRemainingBytes(cbBytes)) return;
	if (cbBytes > cbMaxBytes) return;
	memset(pBYTES, 0, sizeof(BYTE)* cbBytes);
	memcpy(pBYTES, m_lpCur, cbBytes);
	m_lpCur += cbBytes;
}

// cchChar is the length of the source string, NOT counting the NULL terminator
void CBinaryParser::GetStringA(size_t cchChar, _Deref_out_opt_z_ LPSTR* ppStr)
{
	if (!ppStr) return;
	*ppStr = nullptr;
	if (!cchChar) return;
	if (!CheckRemainingBytes(sizeof(CHAR)* cchChar)) return;
	*ppStr = new CHAR[cchChar + 1];
	if (*ppStr)
	{
		memset(*ppStr, 0, sizeof(CHAR)* cchChar);
		memcpy(*ppStr, m_lpCur, sizeof(CHAR)* cchChar);
		(*ppStr)[cchChar] = NULL;
	}

	m_lpCur += sizeof(CHAR)* cchChar;
}

// cchChar is the length of the source string, NOT counting the NULL terminator
void CBinaryParser::GetStringW(size_t cchWChar, _Deref_out_opt_z_ LPWSTR* ppStr)
{
	if (!ppStr) return;
	*ppStr = nullptr;
	if (!cchWChar) return;
	if (!CheckRemainingBytes(sizeof(WCHAR)* cchWChar)) return;
	*ppStr = new WCHAR[cchWChar + 1];
	if (*ppStr)
	{
		memset(*ppStr, 0, sizeof(WCHAR)* cchWChar);
		memcpy(*ppStr, m_lpCur, sizeof(WCHAR)* cchWChar);
		(*ppStr)[cchWChar] = NULL;
	}

	m_lpCur += sizeof(WCHAR)* cchWChar;
}

// No size specified - assume the NULL terminator is in the stream, but don't read off the end
void CBinaryParser::GetStringA(_Deref_out_opt_z_ LPSTR* ppStr)
{
	if (!ppStr) return;
	*ppStr = nullptr;
	size_t cchChar = NULL;
	auto hRes = StringCchLengthA(reinterpret_cast<LPSTR>(m_lpCur), (m_lpEnd - m_lpCur) / sizeof(CHAR), &cchChar);

	if (FAILED(hRes)) return;

	// With string length in hand, we defer to our other implementation
	// Add 1 for the NULL terminator
	GetStringA(cchChar + 1, ppStr);
}

// No size specified - assume the NULL terminator is in the stream, but don't read off the end
void CBinaryParser::GetStringW(_Deref_out_opt_z_ LPWSTR* ppStr)
{
	if (!ppStr) return;
	*ppStr = nullptr;

	size_t cchChar = NULL;
	auto hRes = StringCchLengthW(reinterpret_cast<LPWSTR>(m_lpCur), (m_lpEnd - m_lpCur) / sizeof(WCHAR), &cchChar);

	if (FAILED(hRes)) return;

	// With string length in hand, we defer to our other implementation
	// Add 1 for the NULL terminator
	GetStringW(cchChar + 1, ppStr);
}

size_t CBinaryParser::GetRemainingData(_Out_ LPBYTE* ppRemainingBYTES)
{
	if (!ppRemainingBYTES) return 0;
	*ppRemainingBYTES = nullptr;

	auto cbBytes = RemainingBytes();
	if (cbBytes > 0)
	{
		GetBYTES(cbBytes, cbBytes, ppRemainingBYTES);
	}

	return cbBytes;
}