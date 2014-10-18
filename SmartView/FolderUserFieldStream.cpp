#include "stdafx.h"
#include "..\stdafx.h"
#include "FolderUserFieldStream.h"
#include "..\String.h"
#include "..\InterpretProp2.h"
#include "..\ParseProperty.h"
#include "..\ExtraPropTags.h"

void BinToFolderFieldDefinitionCommon(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_ size_t* lpcbBytesRead, _Out_ FolderFieldDefinitionCommon* pffdcFolderFieldDefinitionCommon);

FolderUserFieldStream::FolderUserFieldStream(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	m_FolderUserFieldsAnsi = { 0 };
	m_FolderUserFieldsUnicode = { 0 };
}

FolderUserFieldStream::~FolderUserFieldStream()
{
	ULONG i = 0;
	if (m_FolderUserFieldsAnsi.FieldDefinitionCount && m_FolderUserFieldsAnsi.FieldDefinitions)
	{
		for (i = 0; i < m_FolderUserFieldsAnsi.FieldDefinitionCount; i++)
		{
			delete[] m_FolderUserFieldsAnsi.FieldDefinitions[i].FieldName;
			delete[] m_FolderUserFieldsAnsi.FieldDefinitions[i].Common.wszFormula;
		}
	}

	delete[] m_FolderUserFieldsAnsi.FieldDefinitions;

	if (m_FolderUserFieldsUnicode.FieldDefinitionCount && m_FolderUserFieldsUnicode.FieldDefinitions)
	{
		for (i = 0; i < m_FolderUserFieldsUnicode.FieldDefinitionCount; i++)
		{
			delete[] m_FolderUserFieldsUnicode.FieldDefinitions[i].FieldName;
			delete[] m_FolderUserFieldsUnicode.FieldDefinitions[i].Common.wszFormula;
		}
	}

	delete[] m_FolderUserFieldsUnicode.FieldDefinitions;
}

void FolderUserFieldStream::Parse()
{
	if (!m_lpBin) return;

	m_Parser.GetDWORD(&m_FolderUserFieldsAnsi.FieldDefinitionCount);

	if (m_FolderUserFieldsAnsi.FieldDefinitionCount && m_FolderUserFieldsAnsi.FieldDefinitionCount < _MaxEntriesSmall)
		m_FolderUserFieldsAnsi.FieldDefinitions = new FolderFieldDefinitionA[m_FolderUserFieldsAnsi.FieldDefinitionCount];

	if (m_FolderUserFieldsAnsi.FieldDefinitions)
	{
		memset(m_FolderUserFieldsAnsi.FieldDefinitions, 0, sizeof(FolderFieldDefinitionA)*m_FolderUserFieldsAnsi.FieldDefinitionCount);
		ULONG i = 0;

		for (i = 0; i < m_FolderUserFieldsAnsi.FieldDefinitionCount; i++)
		{
			m_Parser.GetDWORD(&m_FolderUserFieldsAnsi.FieldDefinitions[i].FieldType);
			m_Parser.GetWORD(&m_FolderUserFieldsAnsi.FieldDefinitions[i].FieldNameLength);

			if (m_FolderUserFieldsAnsi.FieldDefinitions[i].FieldNameLength &&
				m_FolderUserFieldsAnsi.FieldDefinitions[i].FieldNameLength < _MaxEntriesSmall)
			{
				m_Parser.GetStringA(
					m_FolderUserFieldsAnsi.FieldDefinitions[i].FieldNameLength,
					&m_FolderUserFieldsAnsi.FieldDefinitions[i].FieldName);
			}

			size_t cbBytesRead = 0;
			BinToFolderFieldDefinitionCommon(
				(ULONG)m_Parser.RemainingBytes(),
				m_lpBin + m_Parser.GetCurrentOffset(),
				&cbBytesRead,
				&m_FolderUserFieldsAnsi.FieldDefinitions[i].Common);
			m_Parser.Advance(cbBytesRead);
		}
	}

	m_Parser.GetDWORD(&m_FolderUserFieldsUnicode.FieldDefinitionCount);

	if (m_FolderUserFieldsUnicode.FieldDefinitionCount && m_FolderUserFieldsUnicode.FieldDefinitionCount < _MaxEntriesSmall)
		m_FolderUserFieldsUnicode.FieldDefinitions = new FolderFieldDefinitionW[m_FolderUserFieldsUnicode.FieldDefinitionCount];

	if (m_FolderUserFieldsUnicode.FieldDefinitions)
	{
		memset(m_FolderUserFieldsUnicode.FieldDefinitions, 0, sizeof(FolderFieldDefinitionA)*m_FolderUserFieldsUnicode.FieldDefinitionCount);
		ULONG i = 0;

		for (i = 0; i < m_FolderUserFieldsUnicode.FieldDefinitionCount; i++)
		{
			m_Parser.GetDWORD(&m_FolderUserFieldsUnicode.FieldDefinitions[i].FieldType);
			m_Parser.GetWORD(&m_FolderUserFieldsUnicode.FieldDefinitions[i].FieldNameLength);

			if (m_FolderUserFieldsUnicode.FieldDefinitions[i].FieldNameLength &&
				m_FolderUserFieldsUnicode.FieldDefinitions[i].FieldNameLength < _MaxEntriesSmall)
			{
				m_Parser.GetStringW(
					m_FolderUserFieldsUnicode.FieldDefinitions[i].FieldNameLength,
					&m_FolderUserFieldsUnicode.FieldDefinitions[i].FieldName);
			}

			size_t cbBytesRead = 0;
			BinToFolderFieldDefinitionCommon(
				(ULONG)m_Parser.RemainingBytes(),
				m_lpBin + m_Parser.GetCurrentOffset(),
				&cbBytesRead,
				&m_FolderUserFieldsUnicode.FieldDefinitions[i].Common);
			m_Parser.Advance(cbBytesRead);
		}
	}
}

_Check_return_ LPWSTR FolderUserFieldStream::ToString()
{
	Parse();

	wstring szFolderUserFieldStream;
	wstring szTmp;

	szFolderUserFieldStream = formatmessage(
		IDS_FIELDHEADER,
		m_FolderUserFieldsAnsi.FieldDefinitionCount);

	if (m_FolderUserFieldsAnsi.FieldDefinitionCount && m_FolderUserFieldsAnsi.FieldDefinitions)
	{
		ULONG i = 0;
		for (i = 0; i < m_FolderUserFieldsAnsi.FieldDefinitionCount; i++)
		{
			wstring szGUID = GUIDToWstring(&m_FolderUserFieldsAnsi.FieldDefinitions[i].Common.PropSetGuid);
			wstring szFieldType = InterpretFlags(flagFolderType, m_FolderUserFieldsAnsi.FieldDefinitions[i].FieldType);
			wstring szFieldcap = InterpretFlags(flagFieldCap, m_FolderUserFieldsAnsi.FieldDefinitions[i].Common.fcapm);

			szTmp = formatmessage(
				IDS_FIELDANSIFIELD,
				i,
				m_FolderUserFieldsAnsi.FieldDefinitions[i].FieldType, szFieldType.c_str(),
				m_FolderUserFieldsAnsi.FieldDefinitions[i].FieldNameLength,
				m_FolderUserFieldsAnsi.FieldDefinitions[i].FieldName,
				szGUID.c_str(),
				m_FolderUserFieldsAnsi.FieldDefinitions[i].Common.fcapm, szFieldcap.c_str(),
				m_FolderUserFieldsAnsi.FieldDefinitions[i].Common.dwString,
				m_FolderUserFieldsAnsi.FieldDefinitions[i].Common.dwBitmap,
				m_FolderUserFieldsAnsi.FieldDefinitions[i].Common.dwDisplay,
				m_FolderUserFieldsAnsi.FieldDefinitions[i].Common.iFmt,
				m_FolderUserFieldsAnsi.FieldDefinitions[i].Common.wszFormulaLength,
				m_FolderUserFieldsAnsi.FieldDefinitions[i].Common.wszFormula);
			szFolderUserFieldStream += szTmp;
		}
	}

	szTmp = formatmessage(
		IDS_FIELDUNICODEHEADER,
		m_FolderUserFieldsUnicode.FieldDefinitionCount);
	szFolderUserFieldStream += szTmp;

	if (m_FolderUserFieldsUnicode.FieldDefinitionCount && m_FolderUserFieldsUnicode.FieldDefinitions)
	{
		ULONG i = 0;
		for (i = 0; i < m_FolderUserFieldsUnicode.FieldDefinitionCount; i++)
		{
			wstring szGUID = GUIDToWstring(&m_FolderUserFieldsUnicode.FieldDefinitions[i].Common.PropSetGuid);
			wstring szFieldType = InterpretFlags(flagFolderType, m_FolderUserFieldsUnicode.FieldDefinitions[i].FieldType);
			wstring szFieldcap = InterpretFlags(flagFieldCap, m_FolderUserFieldsUnicode.FieldDefinitions[i].Common.fcapm);

			szTmp = formatmessage(
				IDS_FIELDUNICODEFIELD,
				i,
				m_FolderUserFieldsUnicode.FieldDefinitions[i].FieldType, szFieldType.c_str(),
				m_FolderUserFieldsUnicode.FieldDefinitions[i].FieldNameLength,
				m_FolderUserFieldsUnicode.FieldDefinitions[i].FieldName,
				szGUID.c_str(),
				m_FolderUserFieldsUnicode.FieldDefinitions[i].Common.fcapm, szFieldcap.c_str(),
				m_FolderUserFieldsUnicode.FieldDefinitions[i].Common.dwString,
				m_FolderUserFieldsUnicode.FieldDefinitions[i].Common.dwBitmap,
				m_FolderUserFieldsUnicode.FieldDefinitions[i].Common.dwDisplay,
				m_FolderUserFieldsUnicode.FieldDefinitions[i].Common.iFmt,
				m_FolderUserFieldsUnicode.FieldDefinitions[i].Common.wszFormulaLength,
				m_FolderUserFieldsUnicode.FieldDefinitions[i].Common.wszFormula);
			szFolderUserFieldStream += szTmp;
		}
	}

	szFolderUserFieldStream += JunkDataToString();

	return wstringToLPWSTR(szFolderUserFieldStream);
}

void BinToFolderFieldDefinitionCommon(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_ size_t* lpcbBytesRead, _Out_ FolderFieldDefinitionCommon* pffdcFolderFieldDefinitionCommon)
{
	if (!lpBin || !lpcbBytesRead || !pffdcFolderFieldDefinitionCommon) return;
	*pffdcFolderFieldDefinitionCommon = FolderFieldDefinitionCommon();

	CBinaryParser Parser(cbBin, lpBin);

	Parser.GetBYTESNoAlloc(sizeof(GUID), sizeof(GUID), (LPBYTE)&pffdcFolderFieldDefinitionCommon->PropSetGuid);
	Parser.GetDWORD(&pffdcFolderFieldDefinitionCommon->fcapm);
	Parser.GetDWORD(&pffdcFolderFieldDefinitionCommon->dwString);
	Parser.GetDWORD(&pffdcFolderFieldDefinitionCommon->dwBitmap);
	Parser.GetDWORD(&pffdcFolderFieldDefinitionCommon->dwDisplay);
	Parser.GetDWORD(&pffdcFolderFieldDefinitionCommon->iFmt);
	Parser.GetWORD(&pffdcFolderFieldDefinitionCommon->wszFormulaLength);
	if (pffdcFolderFieldDefinitionCommon->wszFormulaLength &&
		pffdcFolderFieldDefinitionCommon->wszFormulaLength < _MaxEntriesLarge)
	{
		Parser.GetStringW(
			pffdcFolderFieldDefinitionCommon->wszFormulaLength,
			&pffdcFolderFieldDefinitionCommon->wszFormula);
	}

	*lpcbBytesRead = Parser.GetCurrentOffset();
}