#include "stdafx.h"
#include "FolderUserFieldStream.h"
#include "String.h"
#include "InterpretProp2.h"
#include "ExtraPropTags.h"

FolderUserFieldStream::FolderUserFieldStream()
{
	m_FolderUserFieldsAnsi = { 0 };
	m_FolderUserFieldsUnicode = { 0 };
}

FolderUserFieldStream::~FolderUserFieldStream()
{
	if (m_FolderUserFieldsAnsi.FieldDefinitionCount && m_FolderUserFieldsAnsi.FieldDefinitions)
	{
		for (DWORD i = 0; i < m_FolderUserFieldsAnsi.FieldDefinitionCount; i++)
		{
			delete[] m_FolderUserFieldsAnsi.FieldDefinitions[i].FieldName;
			delete[] m_FolderUserFieldsAnsi.FieldDefinitions[i].Common.wszFormula;
		}
	}

	delete[] m_FolderUserFieldsAnsi.FieldDefinitions;

	if (m_FolderUserFieldsUnicode.FieldDefinitionCount && m_FolderUserFieldsUnicode.FieldDefinitions)
	{
		for (DWORD i = 0; i < m_FolderUserFieldsUnicode.FieldDefinitionCount; i++)
		{
			delete[] m_FolderUserFieldsUnicode.FieldDefinitions[i].FieldName;
			delete[] m_FolderUserFieldsUnicode.FieldDefinitions[i].Common.wszFormula;
		}
	}

	delete[] m_FolderUserFieldsUnicode.FieldDefinitions;
}

void FolderUserFieldStream::Parse()
{
	m_Parser.GetDWORD(&m_FolderUserFieldsAnsi.FieldDefinitionCount);

	if (m_FolderUserFieldsAnsi.FieldDefinitionCount && m_FolderUserFieldsAnsi.FieldDefinitionCount < _MaxEntriesSmall)
		m_FolderUserFieldsAnsi.FieldDefinitions = new FolderFieldDefinitionA[m_FolderUserFieldsAnsi.FieldDefinitionCount];

	if (m_FolderUserFieldsAnsi.FieldDefinitions)
	{
		memset(m_FolderUserFieldsAnsi.FieldDefinitions, 0, sizeof(FolderFieldDefinitionA)*m_FolderUserFieldsAnsi.FieldDefinitionCount);

		for (DWORD i = 0; i < m_FolderUserFieldsAnsi.FieldDefinitionCount; i++)
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

			BinToFolderFieldDefinitionCommon(
				&m_FolderUserFieldsAnsi.FieldDefinitions[i].Common);
		}
	}

	m_Parser.GetDWORD(&m_FolderUserFieldsUnicode.FieldDefinitionCount);

	if (m_FolderUserFieldsUnicode.FieldDefinitionCount && m_FolderUserFieldsUnicode.FieldDefinitionCount < _MaxEntriesSmall)
		m_FolderUserFieldsUnicode.FieldDefinitions = new FolderFieldDefinitionW[m_FolderUserFieldsUnicode.FieldDefinitionCount];

	if (m_FolderUserFieldsUnicode.FieldDefinitions)
	{
		memset(m_FolderUserFieldsUnicode.FieldDefinitions, 0, sizeof(FolderFieldDefinitionA)*m_FolderUserFieldsUnicode.FieldDefinitionCount);

		for (DWORD i = 0; i < m_FolderUserFieldsUnicode.FieldDefinitionCount; i++)
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

			BinToFolderFieldDefinitionCommon(
				&m_FolderUserFieldsUnicode.FieldDefinitions[i].Common);
		}
	}
}

void FolderUserFieldStream::BinToFolderFieldDefinitionCommon(_Out_ FolderFieldDefinitionCommon* pffdcFolderFieldDefinitionCommon)
{
	*pffdcFolderFieldDefinitionCommon = FolderFieldDefinitionCommon();

	m_Parser.GetBYTESNoAlloc(sizeof(GUID), sizeof(GUID), reinterpret_cast<LPBYTE>(&pffdcFolderFieldDefinitionCommon->PropSetGuid));
	m_Parser.GetDWORD(&pffdcFolderFieldDefinitionCommon->fcapm);
	m_Parser.GetDWORD(&pffdcFolderFieldDefinitionCommon->dwString);
	m_Parser.GetDWORD(&pffdcFolderFieldDefinitionCommon->dwBitmap);
	m_Parser.GetDWORD(&pffdcFolderFieldDefinitionCommon->dwDisplay);
	m_Parser.GetDWORD(&pffdcFolderFieldDefinitionCommon->iFmt);
	m_Parser.GetWORD(&pffdcFolderFieldDefinitionCommon->wszFormulaLength);
	if (pffdcFolderFieldDefinitionCommon->wszFormulaLength &&
		pffdcFolderFieldDefinitionCommon->wszFormulaLength < _MaxEntriesLarge)
	{
		m_Parser.GetStringW(
			pffdcFolderFieldDefinitionCommon->wszFormulaLength,
			&pffdcFolderFieldDefinitionCommon->wszFormula);
	}
}

_Check_return_ wstring FolderUserFieldStream::ToStringInternal()
{
	wstring szFolderUserFieldStream;
	wstring szTmp;

	szFolderUserFieldStream = formatmessage(
		IDS_FIELDHEADER,
		m_FolderUserFieldsAnsi.FieldDefinitionCount);

	if (m_FolderUserFieldsAnsi.FieldDefinitionCount && m_FolderUserFieldsAnsi.FieldDefinitions)
	{
		for (DWORD i = 0; i < m_FolderUserFieldsAnsi.FieldDefinitionCount; i++)
		{
			auto szGUID = GUIDToString(&m_FolderUserFieldsAnsi.FieldDefinitions[i].Common.PropSetGuid);
			auto szFieldType = InterpretFlags(flagFolderType, m_FolderUserFieldsAnsi.FieldDefinitions[i].FieldType);
			auto szFieldcap = InterpretFlags(flagFieldCap, m_FolderUserFieldsAnsi.FieldDefinitions[i].Common.fcapm);

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
		for (DWORD i = 0; i < m_FolderUserFieldsUnicode.FieldDefinitionCount; i++)
		{
			auto szGUID = GUIDToString(&m_FolderUserFieldsUnicode.FieldDefinitions[i].Common.PropSetGuid);
			auto szFieldType = InterpretFlags(flagFolderType, m_FolderUserFieldsUnicode.FieldDefinitions[i].FieldType);
			auto szFieldcap = InterpretFlags(flagFieldCap, m_FolderUserFieldsUnicode.FieldDefinitions[i].Common.fcapm);

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

	return szFolderUserFieldStream;
}